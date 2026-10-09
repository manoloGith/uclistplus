VERSION 5.00
Begin VB.UserControl ucScrollbar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   ToolboxBitmap   =   "ucScrollbar.ctx":0000
   Begin VB.Timer TimSmoothChange 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   30
      Top             =   1890
   End
End
Attribute VB_Name = "ucScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucScrollbar.ctl
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  None
' Last revision: 11/12/2019
' Version:       1.0.9b (Shagratt 2019)
' Version:       1.0.5 (Jason James Newland 2007)
' Version:       1.0.4 (Carles P.V 2005)
'----------------------------------------------------------------------------------------
'
' (*) 1. Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'     2. pvCheckEnvironment() and pvIsLuna() routines by Paul Caton
'
'     3. Flat button fxs code extracted from (see pvDrawFlatButton() routine):
'        Special flat Cool Scrollbars version 1.2 by James Brown
'        http://www.catch22.net/tuts/coolscroll.asp
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.0.1: - Flat style *properly* painted:
'              * Hot thumb appearance = Pressed thumb appearance.
'              * Pressed/hot buttons using correct system colors.
'              Is there a default? For example, ListView with flat-scrollbars flag set,
'              preserves pressed buttons with 1-pixel edge using 'shadow' color and
'              their background is filled using color black instead of 'dark shadow'.
'   * 1.0.2: - Added Refresh method: only for custom-draw purposes.
'   * 1.0.3: - Fixed control on m_bHasTrack and m_bHasNullTrack flags.
'   * 1.0.4: - Fixed thumb rendering (classic style). DrawFrameControl->DrawEdge.
'   * 1.0.5: - Added theme support for Vista in theme mode style (JJN)
'   * 1.0.6: (Shagratt)
'            - Added MouseWhell support on Hover (no focus needed)
'            - Added MouseWhell support over another object (on demand)
'            - Added Public Methods to create object on demand (by code)
'            - Added Smooth Scrolling (can be adjusted.value=1 Disabled)
'            - MouseWheel scroll value can be adjusted
'            - Changed Thumb size formula
'            - Continous TL/BR button scrolling
'            - Thumb scrolling raise Scroll and Change events
'            - Changed default props: Flat, scroll rates,etc.
'            - Changed default props: Flat, scroll rates,etc.
'            - Updated to Paul Caton Self Sub v1.6
'   * 1.0.7: - Added Horizontal Scroll with Shift Pressed (.AttachHorizontalScrollBar() )
'              Thanks dseaman@vbforums for the idea!
'   * 1.0.7b:- Changed dragging thumb to smooth scroll instead of direct to value (smooth drawing too)
'   * 1.0.8: - Changed dragging the thumb is direct visually (only value change smooth)
'            - Fixed lot of errors on thumb draws/calculation introduced in 1.0.7b
'   * 1.0.9: - Fixed Themed Style
'            - Change Themed hover on arrows to emphasize both buttons
'   * 1.0.9b:- Fixed thumb formula (finally!)
'            - Exposed ZOrder
'----------------------------------------------------------------------------------------
'
' Notes:
'
'   * Restriction: Max >= Min
'   * Restriction: TabStop not supported
'----------------------------------------------------------------------------------------
'
' Known issues:
'========================================================================================

Option Explicit

Private Const VERSION_INFO As String = "1.0.9"

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum


Private Type tSubData                                                         'Subclass data type
    hWnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpFn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

'========================================================================================
' UserControl API declarations
'========================================================================================
Private Const SM_CXVSCROLL  As Long = 2
Private Const SM_CYHSCROLL  As Long = 3
Private Const SM_CYVSCROLL  As Long = 20
Private Const SM_CXHSCROLL  As Long = 21
Private Const SM_SWAPBUTTON As Long = 23

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Const SPI_GETKEYBOARDPREF  As Long = 68

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Type Rect
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As Rect) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
'Shagratt 29/11/19
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As KeyCodeConstants) As Integer
Private Const KEY_DOWN As Integer = &H8000

Private ExtWheelHwnd&

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLUP       As Long = &H0
Private Const DFCS_SCROLLDOWN     As Long = &H1
Private Const DFCS_SCROLLLEFT     As Long = &H2
Private Const DFCS_SCROLLRIGHT    As Long = &H3
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000
Private Const DFCS_MONO           As Long = &H8000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const BDR_RAISED As Long = &H5
Private Const BF_RECT    As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const COLOR_BTNFACE     As Long = 15
Private Const COLOR_3DSHADOW    As Long = 16
Private Const COLOR_BTNTEXT     As Long = 18
Private Const COLOR_3DHIGHLIGHT As Long = 20
Private Const COLOR_3DDKSHADOW  As Long = 21

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetTextColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const WHITE_BRUSH As Long = 0
Private Const BLACK_BRUSH As Long = 4

Private Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
    
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreatePatternBrush Lib "GDI32" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
 
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Type PAINTSTRUCT
    hdc             As Long
    fErase          As Long
    rcPaint         As Rect
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

Private Const WM_SIZE           As Long = &H5
Private Const WM_PAINT          As Long = &HF
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_CANCELMODE     As Long = &H1F
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const MK_LBUTTON        As Long = &H1
Private Const WM_MOUSEWHEEL = &H20A    '522
 
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long

' Class name
Private Const SB_THEME As String = "Scrollbar"

' [UxThemeSCROLLBARParts]
Private Const SBP_ARROWBTN = 1
Private Const SBP_THUMBBTNHORZ = 2
Private Const SBP_THUMBBTNVERT = 3
Private Const SBP_LOWERTRACKHORZ = 4
Private Const SBP_UPPERTRACKHORZ = 5
Private Const SBP_LOWERTRACKVERT = 6
Private Const SBP_UPPERTRACKVERT = 7
Private Const SBP_GRIPPERHORZ = 8
Private Const SBP_GRIPPERVERT = 9
Private Const SBP_SIZEBOX = 10

' [UxThemeARROWBTNStates]
Private Const ABS_UPNORMAL = 1
Private Const ABS_UPHOT = 2
Private Const ABS_UPPRESSED = 3
Private Const ABS_UPDISABLED = 4
Private Const ABS_DOWNNORMAL = 5
Private Const ABS_DOWNHOT = 6
Private Const ABS_DOWNPRESSED = 7
Private Const ABS_DOWNDISABLED = 8
Private Const ABS_LEFTNORMAL = 9
Private Const ABS_LEFTHOT = 10
Private Const ABS_LEFTPRESSED = 11
Private Const ABS_LEFTDISABLED = 12
Private Const ABS_RIGHTNORMAL = 13
Private Const ABS_RIGHTHOT = 14
Private Const ABS_RIGHTPRESSED = 15
Private Const ABS_RIGHTDISABLED = 16

' [UxThemeHorzScrollStates]
Private Const HSS_NORMAL = 1
Private Const HSS_HOT = 2
Private Const HSS_PUSHED = 3
Private Const HSS_DISABLED = 4

' [UxThemeHorzThumbStates]
Private Const HTS_NORMAL = 1
Private Const HTS_HOT = 2
Private Const HTS_PUSHED = 3
Private Const HTS_DISABLED = 4

' [UxThemeVertScrollStates]
Private Const VSS_NORMAL = 1
Private Const VSS_HOT = 2
Private Const VSS_PUSHED = 3
Private Const VSS_DISABLED = 4

' [UxThemeVertThumbStates]
Private Const VTS_NORMAL = 1
Private Const VTS_HOT = 2
Private Const VTS_PUSHED = 3
Private Const VTS_DISABLED = 4

'========================================================================================
' UserControl enums., variables and constants
'========================================================================================
'-- Public enums.:
Public Enum sbOrientationCts
    [oVertical] = 0
    [oHorizontal] = 1
End Enum

Public Enum sbStyleCts
    [sClassic] = 0
    [sFlat] = 1
    [sThemed] = 2
    [sCustomDraw] = 3
End Enum

Public Enum sbOnPaintPartCts
    [ppTLButton] = 0
    [ppBRButton] = 1
    [ppTLTrack] = 2
    [ppBRTrack] = 3
    [ppNullTrack] = 4
    [ppThumb] = 5
End Enum

Public Enum sbOnPaintPartStateCts
    [ppsNormal] = 0
    [ppsPressed] = 1
    [ppsHot] = 2
    [ppsDisabled] = 3
End Enum

'-- Private enums.:
Private Enum eFlatButtonStateCts
    [fbsNormal] = 0
    [fbsSelected] = 1
    [fbsHot] = 2
End Enum

'-- Private constants:
Private Const HT_NOTHING          As Long = 0
Private Const HT_TLBUTTON         As Long = 1
Private Const HT_BRBUTTON         As Long = 2
Private Const HT_TLTRACK          As Long = 3
Private Const HT_BRTRACK          As Long = 4
Private Const HT_THUMB            As Long = 5

Private Const TIMERID_CHANGE1     As Long = 1
Private Const TIMERID_CHANGE2     As Long = 2
Private Const TIMERID_HOT         As Long = 3

Private Const CHANGEDELAY_MIN     As Long = 0
Private Const CHANGEFREQUENCY_MIN As Long = 25
Private Const TIMERDT_HOT         As Long = 25

Private Const THUMBSIZE_MIN       As Long = 8
Private Const GRIPPERSIZE_MIN     As Long = 16

'-- Private variables:
Private m_bHasTrack               As Boolean
Private m_bHasNullTrack           As Boolean
Private m_uRctNullTrack           As Rect

Private m_uRctTLButton            As Rect
Private m_uRctBRButton            As Rect
Private m_uRctTLTrack             As Rect
Private m_uRctBRTrack             As Rect
Private m_uRctThumb               As Rect
Private m_lThumbOffset            As Long
Private m_uRctDrag                As Rect

Private m_bTLButtonPressed        As Boolean
Private m_bBRButtonPressed        As Boolean
Private m_bTLTrackPressed         As Boolean
Private m_bBRTrackPressed         As Boolean
Private m_bThumbPressed           As Boolean

Private m_bTLButtonHot            As Boolean
Private m_bBRButtonHot            As Boolean
Private m_bThumbHot               As Boolean

Private m_lAbsRange               As Long
Private m_lThumbPos               As Long
Private m_lThumbSize              As Long
Private m_eHitTest                As Long
Private m_eHitTestHot             As Long
Private m_x                       As Long
Private m_y                       As Long
Private m_lValueStartDrag         As Long

Private m_hPatternBrush           As Long

'-- Property variables:
Private m_lChangeDelay            As Long
Private m_lChangeFrequency        As Long
Private m_lMax                    As Long
Private m_lMin                    As Long
Private m_lValue                  As Long
Private m_lSmallChange            As Long
Private m_lLargeChange            As Long
Private m_eOrientation            As sbOrientationCts
Private m_eStyle                  As sbStyleCts
Private m_bShowButtons            As Boolean

Private m_bIsXP                   As Boolean ' RO
Private m_bIsLuna                 As Boolean ' RO

' Variable to hold 'DisableMouseWheelSupport' property value
Private m_bDisableMouseWheelSupport As Boolean
Private m_TargetValue As Long 'For SmoothScroll
Private m_SinSmoothScrollFactor As Long 'For SmoothScroll
' Variable to hold 'WheelChange' property value
Private m_LonWheelChange As Long
Private m_ucScrollbarH As ucScrollbar


'-- Default property values:
Private Const ENABLED_DEF         As Boolean = True
Private Const MIN_DEF             As Long = 0
Private Const MAX_DEF             As Long = 100
Private Const VALUE_DEF           As Long = MIN_DEF
Private Const SMALLCHANGE_DEF     As Long = 10 '1
Private Const LARGECHANGE_DEF     As Long = 40 '10
Private Const CHANGEDELAY_DEF     As Long = 250
Private Const CHANGEFREQUENCY_DEF As Long = 75
Private Const ORIENTATION_DEF     As Long = [oVertical]
Private Const STYLE_DEF           As Long = [sFlat]
Private Const SHOWBUTTONS_DEF     As Boolean = True

'-- Events:
Public Event Change()
Public Event Scroll()
Public Event ThemeChanged()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnPaint(ByVal lHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal eState As sbOnPaintPartStateCts)

'========================================================================================
' UserControl initialization/termination
'========================================================================================
Private Sub UserControl_Initialize()
    Call pvCreatePatternBrush
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    '-- Stop subclassing
    Call SubclassEnd
    '-- In any case...
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    '-- Clean up
    Call DeleteObject(m_hPatternBrush)
End Sub

'========================================================================================
' Only in design-mode
'========================================================================================
Private Sub UserControl_Resize()
    On Error Resume Next
    If (Ambient.UserMode = False) Then
        Call pvOnSize
    End If
End Sub

Private Sub UserControl_Paint()
    If (Ambient.UserMode = False) Then
        Call pvOnPaint(UserControl.hdc)
    End If
End Sub

Public Property Get Parent() As Object
    Set Parent = UserControl.Extender.Parent
End Property
Public Property Set Parent(ByVal ObjValue As Object)
    Set UserControl.Extender.Parent = ObjValue
End Property

Public Property Get SmoothScrollFactor() As Long
    SmoothScrollFactor = m_SinSmoothScrollFactor
End Property
Public Property Let SmoothScrollFactor(ByVal SinValue As Long)
    m_SinSmoothScrollFactor = SinValue
    PropertyChanged "SmoothScrollFactor"
End Property

Public Property Get Extender() As Object
    Set Extender = UserControl.Extender
End Property

Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property
Public Property Let Visible(ByVal bValue As Boolean)
    UserControl.Extender.Visible = bValue
End Property

Public Property Get Top() As Long
    Top = UserControl.Extender.Top
End Property
Public Property Let Top(ByVal LonValue As Long)
    UserControl.Extender.Top = LonValue
End Property

Public Property Get Left() As Long
    Left = UserControl.Extender.Left
End Property
Public Property Let Left(ByVal LonValue As Long)
    UserControl.Extender.Left = LonValue
End Property

Public Property Get DisableMouseWheelSupport() As Boolean
    DisableMouseWheelSupport = m_bDisableMouseWheelSupport
End Property
Public Property Let DisableMouseWheelSupport(ByVal bValue As Boolean)
    m_bDisableMouseWheelSupport = bValue
    PropertyChanged "DisableMouseWheelSupport"
End Property

'========================================================================================
' Methods
'========================================================================================
Public Sub Refresh()
    '-- Force a complete paint
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Sub

Public Function ZOrder(Optional LonValue As Long = 0)
    UserControl.Extender.ZOrder LonValue
End Function


'*2-
Public Sub AttachHorizontalScrollBar(ucScrollbarH As ucScrollbar)
    
    'Sanity check controls
    If (m_eOrientation = oHorizontal) Then
        Debug.Print UserControl.Extender.Name & " AttachHorizontalScrollBar() can only be called on Vertical Scrollbars"
        Exit Sub
    End If
    
    'Disable
    If (ucScrollbarH Is Nothing) Then
        Set m_ucScrollbarH = ucScrollbarH
        Exit Sub
    End If
    
    If (ucScrollbarH.Orientation = oVertical) Then
        Debug.Print UserControl.Extender.Name & " AttachHorizontalScrollBar() parameter was another Vertical Scrollbar, you need to pass a Horizontal SB."
        Exit Sub
    End If
        
    Set m_ucScrollbarH = ucScrollbarH
    
End Sub

'======================================
'Do a scrolling top/left
'(depending on the type of scrollbar)
'======================================
Public Sub WheelScrollTopLeft()
    m_eHitTest = HT_NOTHING
    Call pvScrollPosDec(m_LonWheelChange, True)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, 150)
End Sub

'======================================
'Do a scrolling bottom/right
'(depending on the type of scrollbar)
'======================================
Public Sub WheelScrollBotRight()
    m_eHitTest = HT_NOTHING
    Call pvScrollPosInc(m_LonWheelChange, True)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, 150)
End Sub

'========================================================================================
' Messages response
'========================================================================================
Private Sub pvOnSize()
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Sub

Private Sub pvOnPaint(ByVal lHdc As Long)
On Error GoTo Err:
    Dim lfHorz As Long
    '
    lfHorz = -CLng(m_eOrientation = [oHorizontal])
    Select Case True
        Case m_eStyle = [sClassic] Or (m_eStyle = [sThemed] And m_bIsLuna = False)
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonPressed) Then
                    Call DrawFrameControl(lHdc, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_FLAT Or DFCS_PUSHED)
                Else
                    Call DrawFrameControl(lHdc, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz))
                End If
                If (m_bBRButtonPressed) Then
                    Call DrawFrameControl(lHdc, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_FLAT Or DFCS_PUSHED)
                Else
                    Call DrawFrameControl(lHdc, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz))
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call FillRect(lHdc, m_uRctTLTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lHdc, m_uRctTLTrack, m_hPatternBrush)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call FillRect(lHdc, m_uRctBRTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lHdc, m_uRctBRTrack, m_hPatternBrush)
                    End If
                    '-- Thumb
                    Call FillRect(lHdc, m_uRctThumb, GetSysColorBrush(COLOR_BTNFACE))
                    Call DrawEdge(lHdc, m_uRctThumb, BDR_RAISED, BF_RECT)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lHdc, m_uRctNullTrack, m_hPatternBrush)
                End If
            Else
                '-- Draw all disabled
                Call DrawFrameControl(lHdc, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_INACTIVE)
                Call DrawFrameControl(lHdc, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_INACTIVE)
                If (m_bHasTrack) Then
                    Call FillRect(lHdc, m_uRctTLTrack, m_hPatternBrush)
                    Call FillRect(lHdc, m_uRctBRTrack, m_hPatternBrush)
                    Call DrawFrameControl(lHdc, m_uRctThumb, 0, 0)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lHdc, m_uRctNullTrack, m_hPatternBrush)
                End If
            End If
        
        Case m_eStyle = [sFlat]
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonHot) Then
                    Call pvDrawFlatButton(lHdc, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsHot])
                Else
                    If (m_bTLButtonPressed) Then
                        Call pvDrawFlatButton(lHdc, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsSelected])
                      Else
                        Call pvDrawFlatButton(lHdc, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsNormal])
                    End If
                End If
                If (m_bBRButtonHot) Then
                    Call pvDrawFlatButton(lHdc, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsHot])
                Else
                    If (m_bBRButtonPressed) Then
                        Call pvDrawFlatButton(lHdc, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsSelected])
                    Else
                        Call pvDrawFlatButton(lHdc, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsNormal])
                    End If
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call FillRect(lHdc, m_uRctTLTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lHdc, m_uRctTLTrack, m_hPatternBrush)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call FillRect(lHdc, m_uRctBRTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lHdc, m_uRctBRTrack, m_hPatternBrush)
                    End If
                    '-- Thumb
                    If (m_bThumbHot) Then
                        Call FillRect(lHdc, m_uRctThumb, GetSysColorBrush(COLOR_3DSHADOW))
                    Else
                        If (m_bThumbPressed) Then
                            Call FillRect(lHdc, m_uRctThumb, GetSysColorBrush(COLOR_3DSHADOW))
                        Else
                            Call DrawFrameControl(lHdc, m_uRctThumb, 0, DFCS_FLAT)
                        End If
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lHdc, m_uRctNullTrack, m_hPatternBrush)
                End If
            Else
                '-- Draw all disabled
                Call DrawFrameControl(lHdc, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_FLAT Or DFCS_INACTIVE)
                Call DrawFrameControl(lHdc, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_FLAT Or DFCS_INACTIVE)
                If (m_bHasTrack) Then
                    Call FillRect(lHdc, m_uRctTLTrack, m_hPatternBrush)
                    Call FillRect(lHdc, m_uRctBRTrack, m_hPatternBrush)
                    Call DrawFrameControl(lHdc, m_uRctThumb, 0, DFCS_FLAT)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lHdc, m_uRctNullTrack, m_hPatternBrush)
                End If
            End If
        Case m_eStyle = [sThemed]
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonHot) Or (m_bBRButtonHot) Then
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_UPHOT + (8 * lfHorz), m_uRctTLButton)
                Else
                    If (m_bTLButtonPressed) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_UPPRESSED + (8 * lfHorz), m_uRctTLButton)
                    Else
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_UPNORMAL + (8 * lfHorz), m_uRctTLButton)
                    End If
                End If
                If (m_bTLButtonHot) Or (m_bBRButtonHot) Then
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_DOWNHOT + (8 * lfHorz), m_uRctBRButton)
                Else
                    If (m_bBRButtonPressed) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_DOWNPRESSED + (8 * lfHorz), m_uRctBRButton)
                    Else
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_DOWNNORMAL + (8 * lfHorz), m_uRctBRButton)
                    End If
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_PUSHED, m_uRctTLTrack)
                    Else
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctTLTrack)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_LOWERTRACKVERT - (2 * lfHorz), HSS_PUSHED, m_uRctBRTrack)
                    Else
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_LOWERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctBRTrack)
                    End If
                    '-- Thumb
                    If (m_bThumbHot) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_HOT, m_uRctThumb)
                        If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                            Call pvDrawThemePart(lHdc, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_HOT, m_uRctThumb)
                        End If
                    Else
                        If (m_bThumbPressed) Then
                            Call pvDrawThemePart(lHdc, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_PUSHED, m_uRctThumb)
                            If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                                Call pvDrawThemePart(lHdc, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_PUSHED, m_uRctThumb)
                            End If
                        Else
                            Call pvDrawThemePart(lHdc, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_NORMAL, m_uRctThumb)
                            If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                               Call pvDrawThemePart(lHdc, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_NORMAL, m_uRctThumb)
                            End If
                        End If
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctNullTrack)
                End If
            Else
                '-- Draw all disabled
                Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_UPDISABLED + (8 * lfHorz), m_uRctTLButton)
                Call pvDrawThemePart(lHdc, SB_THEME, SBP_ARROWBTN, ABS_DOWNDISABLED + (8 * lfHorz), m_uRctBRButton)
                Call pvDrawThemePart(lHdc, SB_THEME, SBP_UPPERTRACKVERT + (2 * lfHorz), HSS_DISABLED, m_uRctTLTrack)
                If (m_bHasTrack) Then
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_LOWERTRACKVERT + (2 * lfHorz), HSS_DISABLED, m_uRctBRTrack)
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_DISABLED, m_uRctThumb)
                    If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                        Call pvDrawThemePart(lHdc, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_DISABLED, m_uRctThumb)
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call pvDrawThemePart(lHdc, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_DISABLED, m_uRctNullTrack)
                End If
            End If
        Case m_eStyle = [sCustomDraw]
            With m_uRctTLButton
                RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppTLButton], IIf(m_bTLButtonHot, [ppsHot], IIf(m_bTLButtonPressed, [ppsPressed], [ppsNormal])))
            End With
            With m_uRctBRButton
                RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppBRButton], IIf(m_bBRButtonHot, [ppsHot], IIf(m_bBRButtonPressed, [ppsPressed], [ppsNormal])))
            End With
            If (m_bHasTrack) Then
                With m_uRctTLTrack
                    RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppTLTrack], IIf(m_bTLTrackPressed, [ppsPressed], [ppsNormal]))
                End With
                With m_uRctBRTrack
                    RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppBRTrack], IIf(m_bBRTrackPressed, [ppsPressed], [ppsNormal]))
                End With
                With m_uRctThumb
                    RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppThumb], IIf(m_bThumbHot, [ppsHot], IIf(m_bThumbPressed, [ppsPressed], [ppsNormal])))
                End With
            End If
            If (m_bHasNullTrack) Then
                With m_uRctNullTrack
                    RaiseEvent OnPaint(lHdc, .X1, .Y1, .X2, .Y2, [ppNullTrack], [ppsNormal])
                End With
            End If
    End Select

Err:
End Sub

Private Sub pvOnMouseDown(ByVal wParam As Long, ByVal lParam As Long)
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        Call pvMakePoints(lParam, m_x, m_y)
        m_eHitTest = pvHitTest(m_x, m_y)
        Select Case m_eHitTest
            Case HT_THUMB
                Select Case m_eOrientation
                    Case [oVertical]
                        m_lThumbOffset = m_uRctThumb.Y1 - m_y
                    Case [oHorizontal]
                        m_lThumbOffset = m_uRctThumb.X1 - m_x
                End Select
                m_bThumbPressed = True
                m_bThumbHot = False
                m_lValueStartDrag = m_lValue
                Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
            Case HT_TLBUTTON
                m_bTLButtonPressed = True
                m_bTLButtonHot = False
                Call pvScrollPosDec(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE2)
                Call pvSetTimer(TIMERID_CHANGE2, m_lChangeDelay)
            Case HT_BRBUTTON
                m_bBRButtonPressed = True
                m_bBRButtonHot = False
                Call pvScrollPosInc(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE2)
                Call pvSetTimer(TIMERID_CHANGE2, m_lChangeDelay)
            Case HT_TLTRACK
                m_bTLTrackPressed = True
                Call pvScrollPosDec(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            Case HT_BRTRACK
                m_bBRTrackPressed = True
                Call pvScrollPosInc(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
        End Select
    End If
End Sub

Private Sub pvOnMouseMove(ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo Err:
Dim auxl&
    Dim lValuePrev As Long, lThumbPosPrev As Long, bPressed As Boolean, bHot As Boolean
    '
    Call pvMakePoints(lParam, m_x, m_y)
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        Select Case m_eHitTest
            Case HT_THUMB
                lValuePrev = m_lValue
                lThumbPosPrev = m_lThumbPos
                If (PtInRect(m_uRctDrag, m_x, m_y)) Then
                    Select Case m_eOrientation
                        Case [oVertical]
                            m_lThumbPos = m_y + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.Y2) Then
                                m_lThumbPos = m_uRctTLButton.Y2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.Y1) Then
                                m_lThumbPos = m_uRctBRButton.Y1 - m_lThumbSize
                            End If
                        Case [oHorizontal]
                            m_lThumbPos = m_x + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.X2) Then
                                m_lThumbPos = m_uRctTLButton.X2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.X1) Then
                                m_lThumbPos = m_uRctBRButton.X1 - m_lThumbSize
                            End If
                    End Select
                    '*2-
                    'Shagratt 08/12/19
                    'For instant replace all and set m_lValue=pvGetScrollPos()
                    auxl& = pvGetScrollPos()
                    If (auxl& <> lValuePrev) Then
                        If (TimSmoothChange.Enabled) Then TimSmoothChange.Enabled = False
                        m_TargetValue = auxl&
                        TimSmoothChange.Enabled = True
                    End If
                Else
                    'Shagratt removed from original Carles P.V. code
                    'This is when you drag the thumb outside the Scrollbar area... not needed
                    'm_lValue = m_lValueStartDrag
                    'm_lThumbPos = pvGetThumbPos()
                End If
                If (m_lThumbPos <> lThumbPosPrev) Then
                    Call pvSizeTrack
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    If (m_lValue <> lValuePrev) Then
                        RaiseEvent Scroll
                    End If
                End If
            Case HT_TLBUTTON
                If (m_bTLButtonPressed) Then
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                End If
            Case HT_BRBUTTON
                If (m_bBRButtonPressed) Then
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                End If
        End Select
    Else
        m_eHitTestHot = pvHitTest(m_x, m_y)
        Select Case m_eHitTestHot
            Case HT_TLBUTTON
                bHot = (PtInRect(m_uRctTLButton, m_x, m_y) <> 0)
                If (m_bTLButtonHot Xor bHot) Then
                    m_bTLButtonHot = True
                    m_bBRButtonHot = False
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            Case HT_BRBUTTON
                bHot = (PtInRect(m_uRctBRButton, m_x, m_y) <> 0)
                If (m_bBRButtonHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = True
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            Case HT_THUMB
                bHot = (PtInRect(m_uRctThumb, m_x, m_y) <> 0)
                If (m_bThumbHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = False
                    m_bThumbHot = True
                    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
        End Select
    End If
Err:
End Sub

Private Sub pvOnMouseUp()
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
'    If (m_eHitTest = HT_THUMB) Then
'        If (m_lValue <> m_lValueStartDrag) Then
'            RaiseEvent Change
'        End If
'    End If
    'Shagratt this is to keep the thumb on the position where we release the mouse
    If (m_eHitTest <> 5) Then
        m_eHitTest = HT_NOTHING
    End If
    m_bTLButtonPressed = False
    m_bBRButtonPressed = False
    m_bThumbPressed = False
    m_bTLTrackPressed = False
    m_bBRTrackPressed = False
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Sub

Private Sub pvOnTimer(ByVal wParam As Long)
    Dim uPt As POINTAPI
    Select Case wParam
        Case TIMERID_CHANGE1
            Call pvKillTimer(TIMERID_CHANGE1)
            Call pvSetTimer(TIMERID_CHANGE2, m_lChangeFrequency)
        Case TIMERID_CHANGE2
            Select Case m_eHitTest
                Case HT_TLBUTTON
                    'Shagratt 30/11/19
                    pvScrollPosDec (m_lSmallChange)
                    Call pvKillTimer(TIMERID_CHANGE2)
                    Call pvSetTimer(TIMERID_CHANGE2, 50)
                Case HT_BRBUTTON
                    'Shagratt 30/11/19
                    pvScrollPosInc (m_lSmallChange)
                    Call pvKillTimer(TIMERID_CHANGE2)
                    Call pvSetTimer(TIMERID_CHANGE2, 50)
                Case HT_TLTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos > m_y) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                            Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos > m_x) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                            Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                            End If
                    End Select
                Case HT_BRTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos + m_lThumbSize < m_y) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                            Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos + m_lThumbSize < m_x) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                            Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                            End If
                    End Select
           End Select
        Case TIMERID_HOT
            Call GetCursorPos(uPt)
            Call ScreenToClient(hWnd, uPt)
            Select Case True
                Case m_bTLButtonHot
                    If (PtInRect(m_uRctTLButton, uPt.X, uPt.Y) = 0) Then
                        m_bTLButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    End If
                Case m_bBRButtonHot
                    If (PtInRect(m_uRctBRButton, uPt.X, uPt.Y) = 0) Then
                        m_bBRButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    End If
                Case m_bThumbHot
                    If (PtInRect(m_uRctThumb, uPt.X, uPt.Y) = 0) Then
                        m_bThumbHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
                    End If
            End Select
                        
    End Select
End Sub

Private Sub pvOnSysColorChange()
    '-- Repaint all
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Sub

Private Sub pvOnThemeChanged()
    '-- Check OS
    Call pvCheckEnvironment
    RaiseEvent ThemeChanged
    '-- Repaint all
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Sub

'========================================================================================
' Private
'========================================================================================
'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------
Private Sub pvSizeButtons()
    Dim uRct As Rect, lButtonSize As Long
    '
    Call GetClientRect(hWnd, uRct)
    m_bHasTrack = False
    m_bHasNullTrack = False
    Select Case m_eOrientation
        Case [oVertical]
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CYVSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .Y2) Then
                    If (2 * lButtonSize < .Y2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                        Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, 0, lButtonSize, .X2, .Y2 - lButtonSize)
                    Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, .Y2 \ 2)
                        Call SetRect(m_uRctBRButton, 0, .Y2 \ 2 + (.Y2 Mod 2), .X2, .Y2)
                        m_bHasNullTrack = CBool(.Y2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, 0, .Y2 \ 2, .X2, .Y2 \ 2 + 1)
                        End If
                    End If
                Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                    Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                End If
            End With
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 250, 25)
        Case [oHorizontal]
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CXHSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .X2) Then
                    If (2 * lButtonSize < .X2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, lButtonSize, 0, .X2 - lButtonSize, .Y2)
                    Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2 \ 2, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 \ 2 + (.X2 Mod 2), 0, .X2, .Y2)
                        m_bHasNullTrack = CBool(.X2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, .X2 \ 2, 0, .X2 \ 2 + 1, .Y2)
                        End If
                    End If
                Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                    Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                End If
            End With
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 25, 250)
    End Select
    '-- No track: avoid pvSizeTrack() calcs.
    If (m_bHasTrack = False) Then
        Call SetRectEmpty(m_uRctTLTrack)
        Call SetRectEmpty(m_uRctBRTrack)
        Call SetRectEmpty(m_uRctThumb)
    End If
End Sub

'*1-
Private Sub pvSizeTrack()
    If (m_bHasTrack) Then
        '-- Tracks and thumbs exist
        Select Case m_eOrientation
            Case [oVertical]
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, 0, m_uRctTLButton.Y2, m_uRctTLButton.X2, m_lThumbPos)
                Call SetRect(m_uRctBRTrack, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.X2, m_uRctBRButton.Y1)
                Call SetRect(m_uRctThumb, 0, m_lThumbPos, m_uRctBRButton.X2, m_lThumbPos + m_lThumbSize)
            Case [oHorizontal]
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, m_uRctTLButton.X2, 0, m_lThumbPos, m_uRctTLButton.Y2)
                Call SetRect(m_uRctBRTrack, m_lThumbPos + m_lThumbSize, 0, m_uRctBRButton.X1, m_uRctBRButton.Y2)
                Call SetRect(m_uRctThumb, m_lThumbPos, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.Y2)
        End Select
    End If
End Sub

Private Function pvGetThumbSize() As Long
    On Error Resume Next
    'Shagratt At last found the problem with my formula :)
    If (m_lLargeChange > m_lAbsRange) Then
        m_lLargeChange = m_lAbsRange
    End If
    Select Case m_eOrientation
        Case [oVertical]
            'Original: pvGetThumbSize = (m_uRctBRButton.y1 - m_uRctTLButton.y2) \ (m_lAbsRange \ m_lLargeChange + 1)
            pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * ((m_lLargeChange / m_lAbsRange))
            ' Sugerencia de JB
            If pvGetThumbSize < (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) / 10 Then pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) / 10
            If pvGetThumbSize > (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * 0.7 Then pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * 0.7
            ' Fin Sugerencia de JB
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
        Case [oHorizontal]
            'Original: pvGetThumbSize = (m_uRctBRButton.x1 - m_uRctTLButton.x2) \ (m_lAbsRange \ m_lLargeChange + 1)
            pvGetThumbSize = CLng((m_uRctBRButton.X1 - m_uRctTLButton.X2) * ((m_lLargeChange / m_lAbsRange)))
            ' Sugerencia de JB
            If pvGetThumbSize < (m_uRctBRButton.X1 - m_uRctTLButton.X2) / 10 Then pvGetThumbSize = (m_uRctBRButton.X1 - m_uRctTLButton.X2) / 10
            If pvGetThumbSize > (m_uRctBRButton.X1 - m_uRctTLButton.X2) * 0.7 Then pvGetThumbSize = (m_uRctBRButton.X1 - m_uRctTLButton.X2) * 0.7
            ' Fin Sugerencia de JB
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
    End Select
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Controling value
'----------------------------------------------------------------------------------------
'Animate to value
Public Function ScrollToValue(lTarget&)
    m_eHitTest = HT_NOTHING
    m_TargetValue = lTarget&
    TimSmoothChange.Enabled = True
End Function

Public Property Get TargetValue()
    TargetValue = m_TargetValue
End Property

Private Sub TimSmoothChange_Timer()
Dim SmoothStep&, vstep&
    If (m_lValue = m_TargetValue) Then
        TimSmoothChange.Enabled = False
        Exit Sub
    End If

    vstep& = m_TargetValue - m_lValue
    SmoothStep& = Abs(vstep) * (m_SinSmoothScrollFactor / 100) '0.3
    If (SmoothStep& < 1) Then SmoothStep& = 1
    
    If (vstep& > 0) Then
        pvScrollPosInc2 SmoothStep, True
    Else
        pvScrollPosDec2 SmoothStep, True
    End If
End Sub

Private Function pvScrollPosDec(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    'If it was animating the scroll with add to the final value, and not the actual value
    If (TimSmoothChange.Enabled) Then
        m_TargetValue = m_TargetValue - lSteps
    Else
        m_TargetValue = m_lValue - lSteps
    End If
    If (m_TargetValue < m_lMin) Then m_TargetValue = m_lMin
    TimSmoothChange.Enabled = True
End Function

Private Function pvScrollPosDec2(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    Dim bChange As Boolean, lValuePrev As Long

    lValuePrev = m_lValue
    m_lValue = m_lValue - lSteps
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
        m_TargetValue = m_lValue
    End If
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    pvScrollPosDec2 = bChange
End Function


Private Function pvScrollPosInc(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    'If it was animating the scroll with add to the final value, and not the actual value
    If (TimSmoothChange.Enabled) Then
        m_TargetValue = m_TargetValue + lSteps
    Else
        m_TargetValue = m_lValue + lSteps
    End If
    If (m_TargetValue > m_lMax) Then m_TargetValue = m_lMax
    TimSmoothChange.Enabled = True
End Function

Private Function pvScrollPosInc2(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    Dim bChange As Boolean, lValuePrev As Long
    '
    lValuePrev = m_lValue
    m_lValue = m_lValue + lSteps
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
        m_TargetValue = m_lValue
    End If
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    pvScrollPosInc2 = bChange
End Function

'----------------------------------------------------------------------------------------
' Positioning thumb from Value (not thumb position!)
'----------------------------------------------------------------------------------------
Private Function pvGetThumbPos() As Long
    On Error Resume Next
    Select Case m_eOrientation
        Case [oVertical]
            'Shagratt if dragged from the thumb set intantly the thumb position
            If (m_eHitTest = 5) Then
                pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_TargetValue - m_lMin)
            Else
                pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
            End If
        Case [oHorizontal]
            If (m_eHitTest = 5) Then
                pvGetThumbPos = m_uRctTLButton.X2 + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) / m_lAbsRange * (m_TargetValue - m_lMin)
            Else
                pvGetThumbPos = m_uRctTLButton.X2 + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
            End If
    End Select
    
    On Error GoTo 0
End Function

'Get value from the thumb pos
Private Function pvGetScrollPos() As Long
    On Error Resume Next
    Select Case m_eOrientation
        Case [oVertical]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.Y2) / (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) * m_lAbsRange
        Case [oHorizontal]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.X2) / (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) * m_lAbsRange
    End Select
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Hit-Test
'----------------------------------------------------------------------------------------
Private Function pvHitTest(ByVal X As Long, ByVal Y As Long) As Long
    Select Case True
        Case PtInRect(m_uRctTLButton, X, Y)
            pvHitTest = HT_TLBUTTON
        Case PtInRect(m_uRctBRButton, X, Y)
            pvHitTest = HT_BRBUTTON
        Case PtInRect(m_uRctTLTrack, X, Y)
            pvHitTest = HT_TLTRACK
        Case PtInRect(m_uRctBRTrack, X, Y)
            pvHitTest = HT_BRTRACK
        Case PtInRect(m_uRctThumb, X, Y)
            pvHitTest = HT_THUMB
    End Select
End Function

Private Sub pvMakePoints(ByVal lPoint As Long, X As Long, Y As Long)
    If (lPoint And &H8000&) Then
        X = &H8000 Or (lPoint And &H7FFF&)
    Else
        X = lPoint And &HFFFF&
    End If
    If (lPoint And &H80000000) Then
        Y = (lPoint \ &H10000) - 1
    Else
        Y = lPoint \ &H10000
    End If
End Sub

'----------------------------------------------------------------------------------------
' Timing
'----------------------------------------------------------------------------------------
Private Sub pvSetTimer(ByVal lTimerID As Long, ByVal ldT As Long)
    Call SetTimer(UserControl.hWnd, lTimerID, ldT, 0)
End Sub

Private Sub pvKillTimer(ByVal lTimerID As Long)
    Call KillTimer(UserControl.hWnd, lTimerID)
    m_eHitTestHot = HT_NOTHING
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------
Private Sub pvDrawFlatButton(ByVal hdc As Long, uRct As Rect, ByVal lfArrowDirection As Long, ByVal eState As eFlatButtonStateCts)
    Dim uRctMem As Rect, hDCMem1 As Long, hDCMem2 As Long
    Dim hBmp1 As Long, hBmp2 As Long, hBmpOld1 As Long, hBmpOld2 As Long
    Dim clrBkOld As Long, clrTextOld As Long
    '
    With uRct
        '-- Monochrome bitmap to convert the arrow to black/white mask
        hDCMem1 = CreateCompatibleDC(hdc)
        hBmp1 = CreateBitmap(.X2 - .X1, .Y2 - .Y1, 1, 1, ByVal 0)
        hBmpOld1 = SelectObject(hDCMem1, hBmp1)
        '-- Normal bitmap to draw the arrow into
        hDCMem2 = CreateCompatibleDC(hdc)
        hBmp2 = CreateCompatibleBitmap(hdc, .X2 - .X1, .Y2 - .Y1)
        hBmpOld2 = SelectObject(hDCMem2, hBmp2)
        '-- Draw frame normaly
        Call CopyRect(uRctMem, uRct)
        Call OffsetRect(uRctMem, -.X1, -.Y1)
        Call DrawFrameControl(hDCMem2, uRctMem, DFC_SCROLL, DFCS_FLAT Or lfArrowDirection)
        Select Case eState
            Case [fbsNormal]
                '-- Nothing to do
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            Case [fbsSelected]
                '-- Invert
                Call InvertRect(hDCMem2, uRctMem)
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            Case [fbsHot]
                '-- Mask glyph
                Call SetBkColor(hDCMem2, GetSysColor(COLOR_BTNTEXT))
                Call BitBlt(hDCMem1, 0, 0, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
                clrBkOld = SetBkColor(hdc, GetSysColor(COLOR_3DHIGHLIGHT))
                clrTextOld = SetTextColor(hdc, GetSysColor(COLOR_3DSHADOW))
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem1, 0, 0, vbSrcCopy)
                Call SetBkColor(hdc, clrBkOld)
                Call SetTextColor(hdc, clrTextOld)
        End Select
    End With
    '-- Clean up
    Call DeleteObject(SelectObject(hDCMem1, hBmpOld1))
    Call DeleteObject(SelectObject(hDCMem2, hBmpOld2))
    Call DeleteDC(hDCMem1)
    Call DeleteDC(hDCMem2)
End Sub

Private Function pvDrawThemePart(ByVal lHdc As Long, ByVal sClass As String, ByVal lPart As Long, ByVal lState As Long, lpRect As Rect) As Boolean
    Dim hTheme As Long
    On Error GoTo Catch
    '
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
    If (hTheme <> 0) Then
        pvDrawThemePart = (DrawThemeBackground(hTheme, lHdc, lPart, lState, lpRect, lpRect) = 0)
    End If
Catch:
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------
'-- Creating a pattern bitmap (track)
Private Sub pvCreatePatternBrush()
    Dim hBitmap As Long, nPattern(1 To 8) As Integer
    '
    '-- Brush pattern (8x8)
    nPattern(1) = &HAA
    nPattern(2) = &H55
    nPattern(3) = &HAA
    nPattern(4) = &H55
    nPattern(5) = &HAA
    nPattern(6) = &H55
    nPattern(7) = &HAA
    nPattern(8) = &H55
    '-- Create brush from bitmap
    hBitmap = CreateBitmap(8, 8, 1, 1, nPattern(1))
    m_hPatternBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

'-- Checking environment and Windows theming
Private Sub pvCheckEnvironment()
    'modified by Jason James Newland 2007
    'comented by Leandro Ascierto 2021
    
'    Dim uOSV As OSVERSIONINFO
'    '
'    m_bIsXP = False
'    m_bIsLuna = False
'
'    With uOSV
'        .dwOSVersionInfoSize = Len(uOSV)
'        Call GetVersionEx(uOSV)
'        If (.dwPlatformId = 2) Then
'            If (.dwMajorVersion >= 5) Then    ' NT based
'                If (.dwMinorVersion > 0) Then ' XP
'                    m_bIsXP = True
'                    m_bIsLuna = pvIsLuna()
'                End If
'            End If
'        End If
'    End With
'
'
    m_bIsXP = True
    m_bIsLuna = True
End Sub

Private Function pvIsLuna() As Boolean
    'modified by Jason James Newland 2007
    Dim hLib   As Long, lPos  As Long, sTheme As String, sName As String
    '-- Be sure that the theme dll is present
    hLib = LoadLibrary("uxtheme.dll")
    If (hLib <> 0) Then
        '-- Get the theme file name
        sTheme = String$(255, 0)
        Call GetCurrentThemeName(StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0)
        lPos = InStr(1, sTheme, Chr$(0))
        If (lPos > 0) Then
            '-- Get the canonical theme name
            sTheme = Left$(sTheme, lPos - 1)
            sName = String$(255, 0)
            Call GetThemeDocumentationProperty(StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName))
            lPos = InStr(1, sName, Chr$(0))
            If (lPos > 0) Then
                '-- Is it Luna or Areo?
                sName = Left$(sName, lPos - 1)
                pvIsLuna = IIf(LenB(sName) <> 0, True, False)
            End If
        End If
        Call FreeLibrary(hLib)
    End If
End Function

'========================================================================================
' UserControl persistent properties
'========================================================================================
Private Sub UserControl_InitProperties()
    '-- Initialization default values
    Let m_lChangeDelay = CHANGEDELAY_DEF
    Let m_lChangeFrequency = CHANGEFREQUENCY_DEF
    Let m_lMin = MIN_DEF
    Let m_lMax = MAX_DEF
    Let m_lValue = VALUE_DEF
    Let m_lSmallChange = SMALLCHANGE_DEF
    Let m_lLargeChange = LARGECHANGE_DEF
    Let m_eOrientation = ORIENTATION_DEF
    Let m_eStyle = STYLE_DEF
    Let m_bShowButtons = SHOWBUTTONS_DEF
    m_bDisableMouseWheelSupport = False
    
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    m_SinSmoothScrollFactor = 15
    m_LonWheelChange = 3000
End Sub

'*1
'=============================================
'Shagratt: Needed cause if we add from code
' the readproperties is never called
'=============================================
Public Sub AddedFromCodeINIT_AND_DEFAULTS()
   
    '-- Check OS and Luna theme
    Call pvCheckEnvironment
    
    Call SubclassingStart
    
End Sub

'Shagratt 29/11/19
Public Function TrackMouseWheelOnHwnd(lHwnd&)
    'Keep the handle so we can check if its over its updated position
    ExtWheelHwnd& = lHwnd
End Function

'Shagratt 29/11/19
Public Function TrackMouseWheelOnHwndStop()
    ExtWheelHwnd& = 0
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    '-- Bag properties
    With PropBag
        '-- Read inherently-stored properties
        Let UserControl.Enabled = .ReadProperty("Enabled", ENABLED_DEF)
        '-- Read 'memory' properties
        Let m_lMin = .ReadProperty("Min", MIN_DEF)
        Let m_lMax = .ReadProperty("Max", MAX_DEF)
        Let m_lValue = .ReadProperty("Value", VALUE_DEF)
        Let m_lSmallChange = .ReadProperty("SmallChange", SMALLCHANGE_DEF)
        Let m_lLargeChange = .ReadProperty("LargeChange", LARGECHANGE_DEF)
        Let m_lChangeDelay = .ReadProperty("ChangeDelay", CHANGEDELAY_DEF)
        Let m_lChangeFrequency = .ReadProperty("ChangeFrequency", CHANGEFREQUENCY_DEF)
        Let m_eOrientation = .ReadProperty("Orientation", ORIENTATION_DEF)
        Let m_eStyle = .ReadProperty("Style", STYLE_DEF)
        Let m_bShowButtons = .ReadProperty("ShowButtons", SHOWBUTTONS_DEF)
        Let m_bDisableMouseWheelSupport = .ReadProperty("DisableMouseWheelSupport", False)
    End With
    '-- Initialize rectangles

    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    '-- Run-time?

    If (Ambient.UserMode) Then
        '-- Check OS and Luna theme
        Call pvCheckEnvironment
        
        Call SubclassingStart
    
    End If
    m_SinSmoothScrollFactor = PropBag.ReadProperty("SmoothScrollFactor", 30)
    m_LonWheelChange = PropBag.ReadProperty("WheelChange", 3000)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, ENABLED_DEF)
        Call .WriteProperty("Min", m_lMin, MIN_DEF)
        Call .WriteProperty("Max", m_lMax, MAX_DEF)
        Call .WriteProperty("Value", m_lValue, VALUE_DEF)
        Call .WriteProperty("SmallChange", m_lSmallChange, SMALLCHANGE_DEF)
        Call .WriteProperty("LargeChange", m_lLargeChange, LARGECHANGE_DEF)
        Call .WriteProperty("ChangeDelay", m_lChangeDelay, CHANGEDELAY_DEF)
        Call .WriteProperty("ChangeFrequency", m_lChangeFrequency, CHANGEFREQUENCY_DEF)
        Call .WriteProperty("Orientation", m_eOrientation, ORIENTATION_DEF)
        Call .WriteProperty("Style", m_eStyle, STYLE_DEF)
        Call .WriteProperty("ShowButtons", m_bShowButtons, SHOWBUTTONS_DEF)
        Call .WriteProperty("DisableMouseWheelSupport", m_bDisableMouseWheelSupport, False)
    End With
    Call PropBag.WriteProperty("SmoothScrollFactor", m_SinSmoothScrollFactor, 30)
    Call PropBag.WriteProperty("WheelChange", m_LonWheelChange, 3000)
End Sub

Public Property Get WheelChange() As Long
    WheelChange = m_LonWheelChange
End Property
Public Property Let WheelChange(ByVal LonValue As Long)
    m_LonWheelChange = LonValue
    PropertyChanged "WheelChange"
End Property


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enable As Boolean)
    UserControl.Enabled = New_Enable
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property
Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < m_lMin) Then
        New_Max = m_lMin
    End If
    m_lMax = New_Max
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get Min() As Long
    Min = m_lMin
End Property
Public Property Let Min(ByVal New_Min As Long)
    If (New_Min > m_lMax) Then
        New_Min = m_lMax
    End If
    m_lMin = New_Min
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get Value() As Long
    Value = m_lValue
End Property
Public Property Let Value(ByVal New_Value As Long)
    Dim lValuePrev As Long
    '
    If (New_Value < m_lMin) Then
        New_Value = m_lMin
    ElseIf (New_Value > m_lMax) Then
        New_Value = m_lMax
    End If
    lValuePrev = m_lValue
    m_lValue = New_Value
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
    '
    If (m_lValue <> lValuePrev) Then
        RaiseEvent Change
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_lSmallChange
End Property
Public Property Let SmallChange(ByVal New_SmallChange As Long)
    If (New_SmallChange < 1) Then
        New_SmallChange = 1
    End If
    m_lSmallChange = New_SmallChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_lLargeChange
End Property
Public Property Let LargeChange(ByVal New_LargeChange As Long)
    If (New_LargeChange < 1) Then
        New_LargeChange = 1
    End If
    m_lLargeChange = New_LargeChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get ChangeDelay() As Long
    ChangeDelay = m_lChangeDelay
End Property
Public Property Let ChangeDelay(ByVal New_ChangeDelay As Long)
    If (New_ChangeDelay < CHANGEDELAY_MIN) Then
        New_ChangeDelay = CHANGEDELAY_MIN
    End If
    m_lChangeDelay = New_ChangeDelay
End Property

Public Property Get ChangeFrequency() As Long
    ChangeFrequency = m_lChangeFrequency
End Property

Public Property Let ChangeFrequency(ByVal New_ChangeFrequency As Long)
    If (New_ChangeFrequency < CHANGEFREQUENCY_MIN) Then
        New_ChangeFrequency = CHANGEFREQUENCY_MIN
    End If
    m_lChangeFrequency = New_ChangeFrequency
End Property

Public Property Get Orientation() As sbOrientationCts
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As sbOrientationCts)
    If (New_Orientation < [oVertical]) Then
        New_Orientation = [oVertical]
    ElseIf (New_Orientation > [oHorizontal]) Then
        New_Orientation = [oHorizontal]
    End If
    m_eOrientation = New_Orientation
    Call pvOnSize
End Property

Public Property Get Style() As sbStyleCts
    Style = m_eStyle
End Property

Public Property Let Style(ByVal New_Style As sbStyleCts)
    If (New_Style < [sClassic]) Then
        New_Style = [sClassic]
    ElseIf (New_Style > [sCustomDraw]) Then
        New_Style = [sCustomDraw]
    End If
    m_eStyle = New_Style
    Call InvalidateRect(UserControl.hWnd, ByVal 0, 0)
End Property

Public Property Get ShowButtons() As Boolean
    ShowButtons = m_bShowButtons
End Property

Public Property Let ShowButtons(ByVal New_ShowButtons As Boolean)
    m_bShowButtons = New_ShowButtons
    Call pvOnSize
End Property

'// Runtime read only
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IsXP() As Boolean
    IsXP = m_bIsXP
End Property

Public Property Get IsThemed() As Boolean
    IsThemed = m_bIsLuna
End Property

'========================================================================================
' About
'========================================================================================
Public Sub About()
    Call VBA.MsgBox("ucScrollbar " & VERSION_INFO & " - Shagratt MOD (Original by Carles P.V. 2005)", , "About")
End Sub

'Replace with favourite subclassing
Private Sub SubclassingStart()
    
    'Paul Caton
    Call sc_Subclass(UserControl.hWnd) 'Optional pass and object with ObjPtr(pic(0))
    Call sc_AddMsg(UserControl.hWnd, WM_PAINT, [MSG_BEFORE])
    Call sc_AddMsg(UserControl.hWnd, WM_SIZE, [MSG_BEFORE])
    Call sc_AddMsg(UserControl.hWnd, WM_CANCELMODE)
    Call sc_AddMsg(UserControl.hWnd, WM_MOUSEMOVE)
    Call sc_AddMsg(UserControl.hWnd, WM_LBUTTONDOWN)
    Call sc_AddMsg(UserControl.hWnd, WM_LBUTTONUP)
    Call sc_AddMsg(UserControl.hWnd, WM_LBUTTONDBLCLK)
    Call sc_AddMsg(UserControl.hWnd, WM_TIMER)
    Call sc_AddMsg(UserControl.hWnd, WM_SYSCOLORCHANGE)
    Call sc_AddMsg(UserControl.hWnd, WM_MOUSEWHEEL, [MSG_BEFORE])
    If (m_bIsXP) Then
        Call sc_AddMsg(UserControl.hWnd, WM_THEMECHANGED)
    End If
    
    'To receive mousewheel and process when is over control
    Call sc_Subclass(UserControl.ContainerHwnd)
    Call sc_AddMsg(UserControl.ContainerHwnd, WM_MOUSEWHEEL, MSG_BEFORE)

End Sub

Private Sub SubclassEnd()
    'Paul Caton
    sc_Terminate
End Sub

'-SelfSub code------------------------------------------------------------------------------------
'==================================================================================================
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'* v1.4 Optional IDE protection added
'*      User-defined callback parameter added
'*      All user routines that pass in a hWnd get additional validation
'*      End removed from zError.......................................................... 20060411
'* v1.5 Added nOrdinal parameter to sc_Subclass
'*      Switched machine-code array from Currency to Long................................ 20060412
'* v1.6 Added an optional callback target object
'*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
'*************************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  'zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
  '  zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal newValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = newValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  'zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'*-
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lHwnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Object)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
On Error GoTo Err
Dim uPS As PAINTSTRUCT

    Select Case lHwnd
        Case UserControl.hWnd
            Select Case uMsg
            Case WM_PAINT
                Call BeginPaint(lHwnd, uPS)
                Call pvOnPaint(uPS.hdc)
                Call EndPaint(lHwnd, uPS)
                bHandled = True: lReturn = 0
            Case WM_SIZE
                Call pvOnSize
                bHandled = True: lReturn = 0
            Case WM_LBUTTONDOWN
                Call pvOnMouseDown(wParam, lParam)
            Case WM_MOUSEMOVE
                Call pvOnMouseMove(wParam, lParam)
            Case WM_LBUTTONUP, WM_CANCELMODE
                Call pvOnMouseUp
            Case WM_LBUTTONDBLCLK
                Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            Case WM_TIMER
                Call pvOnTimer(wParam)
            Case WM_SYSCOLORCHANGE
                Call pvOnSysColorChange
            Case WM_THEMECHANGED
                Call pvOnThemeChanged
            'Case WM_MOUSEWHEEL
            '    Not needed. We get the MouseWheel from the parent. This way we dont need focus on the scrollbar
            End Select
    Case Else
    
        'sLog UserControl.hwnd & "," & lhWnd & "-[" & Hex(uMsg) & "] "

        If (uMsg = WM_MOUSEWHEEL) Then
            Dim PT As POINTAPI, R As Rect
            GetCursorPos PT
            GetWindowRect UserControl.hWnd, R
            'If Mouse Wheel is used over the scrollbar (any part) process it
            If (PtInRect(R, PT.X, PT.Y)) Then
                'sLog UserControl.hwnd & " IN OVER"
                bHandled = True: lReturn = 0
            
            'If its tracking Mouse Wheel over another object
            ElseIf (ExtWheelHwnd& <> 0) Then
                GetWindowRect ExtWheelHwnd&, R
                'GetWindowRect UserControl.Parent.hwnd, r
                'Check if mouse if over the obj tracking area
                If (PtInRect(R, PT.X, PT.Y)) Then
                    'sLog UserControl.hwnd & " EXT OVER EXT"
                    bHandled = True: lReturn = 0
                End If
            End If
            
            If (bHandled) Then
            
                'Have a Horizontal Scrollbar attached? check for SHIFT key pressed
                If Not (m_ucScrollbarH Is Nothing) Then
                    If GetKeyState(vbKeyShift) And KEY_DOWN Then
                        'Relay the scroll to the Horizontal scrollbar
                        If (wParam > 0) Then
                            m_ucScrollbarH.WheelScrollTopLeft
                        Else
                            m_ucScrollbarH.WheelScrollBotRight
                        End If
                        
                        Exit Sub
                    End If
                End If
            
                If (wParam > 0) Then
                    'Top/Left ...scroll up
                    WheelScrollTopLeft
                Else
                    'Bottom/Right ...scroll down
                    WheelScrollBotRight
                End If
            End If
            
        End If
    End Select
    
Err:
End Sub







