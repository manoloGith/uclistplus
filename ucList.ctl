VERSION 5.00
Begin VB.UserControl ucList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "ucList.ctx":0000
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   1320
   End
   Begin Proyecto1.ucScrollbar ucScrollbar1 
      Height          =   3855
      Left            =   4440
      Top             =   720
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   6800
      SmoothScrollFactor=   75
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "ucList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------------
'Name: ucListPlus
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Publish Date: 7/Jul/2021
'Last Update: 31/12/ 2021
'Version: 1.0.4
'--------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As Rect, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mpath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type COLORMATRIX
    M(0 To 4, 0 To 4)           As Single
End Type

Private Const SmoothingModeAntiAlias    As Long = 4
Private Const UnitPixel                 As Long = 2&
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const CombineModeIntersect      As Long = &H1

Private Const GWL_HWNDPARENT    As Long = -8
Private Const HWND_TOPMOST      As Long = -1
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_SHOWWINDOW    As Long = &H40


Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_VCENTER        As Long = &H4
Private Const DT_WORD_ELLIPSIS  As Long = &H40000
Private Const DT_CENTER         As Long = &H1
Private Const CLR_NONE          As Long = &HFFFFFFFF
Private Const DT_RIGHT          As Long = &H2
Private Const DT_CALCRECT       As Long = &H400

Private Const WM_HOTKEY         As Long = &H312
Private Const WM_MOUSELEAVE     As Long = &H2A3&
Private Const WM_CHAR           As Long = &H102
Private Const WM_IME_SETCONTEXT As Long = &H281&

Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const EM_SETSEL As Long = &HB1

'Private Const WM_KILLFOCUS      As Long = &H8&
'Private Const WM_ACTIVATE       As Long = &H6&
'Private Const WM_CAPTURECHANGED As Long = &H215&
'Private Const WM_WINDOWPOSCHANGING As Long = &H46
'Private Const WM_WINDOWPOSCHANGED As Long = &H47&
'Private Const WM_MOUSEFIRST As Long = &H200
'Private Const WM_SHOWWINDOW As Long = &H18&

'Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const GA_ROOT As Long = 2

Private Type tItem
    Text        As String
    SubText     As String
    Icon        As Long
    Data        As Variant
    Separator   As Boolean
    Checked     As Boolean
    IsChild     As Boolean
    IsGroup     As Boolean
    Disabled    As Boolean
    Visible     As Boolean
    SubMenu     As ucList
    TextColor   As Long
    SubTextColor As Long
    BackColor   As Long
    ItemBold    As Boolean
End Type

Public Enum eListDesign
    GoogleDesign
    MicrosoftDesign
End Enum

Public Enum eListStyle
    StyleDropDown
    StyleListBox
    StyleMenu
End Enum

Public Enum eSortOrder
    SortAscendant
    SortDescendant
End Enum

Public Enum ListPlusOLEDropModeConstants
    OLEDropModeNone = vbOLEDropNone
    OLEDropModeManual = vbOLEDropManual
End Enum

Public Event IndexChange()
Public Event ItemClick()
Public Event HideList()
Public Event ScrollBottom()
Public Event PostAutoComplete()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PostPaint(ByVal hdc As Long)
Public Event ItemCustomDraw(Index As Long, ByVal hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Dim m_ListStyle As eListStyle
Dim m_SubTitleVisible As Boolean
Dim m_ShowFilter As Boolean
Dim m_ItemHeight As Long
Dim m_MaxListHeight As Long
Dim m_MaxVisibleItems As Long
Dim m_ListWidth As Long
Dim m_ListDesign As eListDesign
Dim m_MultiSelect  As Boolean
Dim m_UseMonocromeIcons As Boolean
Dim m_BorderRadius As Long
Dim m_BorderColor As OLE_COLOR
Dim m_DrawSeparatorLine As Boolean
Dim m_ImagesRadius As Long
Dim m_SelectionRadius As Long
Dim m_FindTextColor As OLE_COLOR
Dim m_SelectionColor As OLE_COLOR
Dim m_LineColor As OLE_COLOR
Dim m_BorderVisible As Boolean
Dim m_Left
Dim m_Index         As Long
Dim DpiFactor       As Single
Dim cShadow         As ClsShadow
Dim hRgn            As Long
Dim BR              As Long 'BorderRadius
Dim m_Items()       As tItem
Dim mListCount       As Long
Dim SeparatorsCount As Long
Dim m_HotItem       As Long
Dim IH              As Long 'ItemHeight
Dim GroupIsOpen     As Boolean
Dim m_SegoeFont     As StdFont
Dim m_Wingdings2    As StdFont
Dim m_Wingdings3     As StdFont

Dim ListHeight      As Long
Dim ListTop         As Long
'Dim StrFilter       As String
Dim lHeadHeight     As Long
Dim m_PhWnd         As Long
Dim hAnsestor       As Long
Dim bShow           As Boolean
Dim lMargin         As Long
Dim lCheckSize      As Long
Dim bMouseDown      As Boolean
Dim bFiltered       As Boolean
Dim cImageList      As Collection
Dim ImgLstWidth     As Long
Dim ImgLstHeight    As Long
Dim isSegoeFontInstaled As Boolean
Dim PrevFocus       As Long
Dim cSubClass       As clsSubClass
Dim bKillFocus      As Boolean
Dim m_BtnTop        As Long
Dim m_BtnLeft       As Long
Dim isWinChild         As Boolean
Dim bHasFocus       As Boolean
Dim bHideSubMenu As Boolean
Dim bShowSubMenu As Boolean
Dim TickCount       As Long
Public ParentMenu   As ucList
Public ChildMenu As ucList
Private SubMenuVisible As Boolean
Dim bTimerSkipRoutines As Boolean
Dim VisibleItemsCount As Long
Dim m_OffsetTextBox As Rect
Private WithEvents Frm As Form
Attribute Frm.VB_VarHelpID = -1
Private SrcTextBox As Long
Attribute SrcTextBox.VB_VarHelpID = -1

Public Sub SetKeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Public Sub SetKeyDown(KeyCode As Integer)
    Call Text1_KeyDown(KeyCode, 0)
End Sub

Public Function AutoCompleteTextBox(oTextBox As Long, Optional OffsetLeft As Long, Optional OffsetTop As Long, Optional OffsetRight As Long, Optional OffsetBottom As Long)
    
    If SrcTextBox <> 0 Then
        cSubClass.ssc_UnSubclass SrcTextBox
    End If
    
    SrcTextBox = oTextBox
    
    If oTextBox = 0 Then Exit Function
    
    Text1.Text = vbNullString
    
    SetRect m_OffsetTextBox, OffsetLeft, OffsetTop, OffsetRight, OffsetBottom
    With cSubClass
        If .ssc_Subclass(SrcTextBox, , , Me) Then
            .ssc_AddMsg SrcTextBox, WM_CHAR, MSG_AFTER
        End If
    End With
End Function

Private Sub SrcTextBox_Change()
    Dim Rect As Rect, i As Long
    
    If bTimerSkipRoutines = False Then
        GetWindowRect SrcTextBox, Rect
        If bShow = False Then
            For i = 0 To mListCount - 1
                If Not m_Items(i).IsGroup Then
                    If InStr(UCase(m_Items(i).Text), UCase(TextBox_GetText(SrcTextBox))) Then
                        With m_OffsetTextBox
                            Me.ListWidth = ((Rect.Right + .Right) - (Rect.Left + .Left)) / DpiFactor
                            ShowList Rect.Left + .Left, Rect.Bottom + .Top, 0, Rect.Top + .Bottom
                        End With
                        GoTo Filter
                    End If
                End If
            Next
            Exit Sub
        Else
            If Len(TextBox_GetText(SrcTextBox)) = 0 Then HideList
        End If
        
Filter:
        Text1.Text = TextBox_GetText(SrcTextBox)
        If VisibleItemsCount = 0 Then
            HideList
        End If
    End If

End Sub

Public Sub AddSubMenu(SubMenuTitle As String, GList As ucList, Optional ByVal Icon As Long = -1, Optional ByVal ItemData As Variant)
    ReDim Preserve m_Items(mListCount)
    
    With m_Items(mListCount)
        .Text = SubMenuTitle
        .Icon = Icon
        .Data = ItemData
        .Visible = True
        .TextColor = CLR_NONE
        .SubTextColor = CLR_NONE
        .BackColor = CLR_NONE
        Set .SubMenu = GList
    End With

    mListCount = mListCount + 1

End Sub

Public Sub OpenGroup(GropuTitle As String, Optional ByVal Icon As Long = -1, Optional ByVal ItemData As Variant)
    ReDim Preserve m_Items(mListCount)
    
    With m_Items(mListCount)
        .Text = GropuTitle
        .Icon = Icon
        .Data = ItemData
        .IsGroup = True
        .Visible = True
        .TextColor = CLR_NONE
        .SubTextColor = CLR_NONE
        .BackColor = CLR_NONE
    End With

    mListCount = mListCount + 1
    GroupIsOpen = True
End Sub

Public Sub CloseGroup()
    GroupIsOpen = False
End Sub

Public Function RemoveItem(Index As Long)
    Dim i As Long
    'If Index < 0 Or Index > mListCount Then Exit Function
    If m_Index = Index Then m_Index = m_Index - 1
    mListCount = mListCount - 1

    For i = Index To mListCount - 1
        m_Items(i) = m_Items(i + 1)
    Next
    
 
    
    
    
    ReDim Preserve m_Items(mListCount)
    
End Function

Public Function AddItem(Optional ByVal Text As String, Optional ByVal Icon As Long = -1, Optional ByVal ItemData As Variant, Optional ByVal SubText As String, Optional Checked As Boolean, Optional Disabled As Boolean, Optional ByVal Index As Long = -1)
    Dim i As Long
    ReDim Preserve m_Items(mListCount)
    
    If Index > -1 Then
        If Index > mListCount Then
            Index = mListCount
        Else
            For i = mListCount To Index + 1 Step -1
                m_Items(i) = m_Items(i - 1)
            Next
        End If
    Else
        Index = mListCount
    End If
    
    With m_Items(Index)
        .Text = Text
        .Icon = Icon
        .Data = ItemData
        .SubText = SubText
        .Separator = Text = "-"
        .Checked = Checked
        .IsChild = GroupIsOpen
        .Disabled = Disabled
        .Visible = True
        .TextColor = CLR_NONE
        .SubTextColor = CLR_NONE
        .BackColor = CLR_NONE
        If .Separator Then SeparatorsCount = SeparatorsCount + 1
        mListCount = mListCount + 1
    End With
End Function

Public Function GetItemTop(ByVal Index As Long) As Long
    Dim i As Long, lTop As Long
    
    If m_ShowFilter Then
        Text1.Height = UserControl.Font.Size * 1.5 * DpiFactor
        lTop = Text1.Height + lMargin * 4
    Else
        lTop = lMargin
    End If

    lTop = lTop - ucScrollbar1.Value
    
    If Index < 0 Or Index > mListCount - 1 Then
        GetItemTop = lTop
        Exit Function
    End If
    
    For i = 0 To mListCount - 1
        With m_Items(i)
            If Index = i Then
                GetItemTop = lTop
                Exit Function
            End If
            If .Separator Then
                lTop = lTop + IH / 2
            Else
                lTop = lTop + IH
            End If
        End With
    Next
End Function

'*1
Private Sub Draw()
    Dim i As Long
    Dim lTop As Long, lLeft As Long
    Dim Rect As Rect
    Dim lWidth As Long
    Dim TextHeight As Long
    Dim oColor As OLE_COLOR
    Dim HotItembckColor As Long
    Dim HighlightTextColor As Long
    Dim NormalTextColor As Long
    Dim mFontSize As Single
    Dim sText As String, sShortcut As String
    Dim lPos As String
    Dim PrevForeColor As Long
    
    
    Cls
    
    If IsDarkColor(m_SelectionColor) Then
        HighlightTextColor = vbWhite
    Else
        HighlightTextColor = vbBlack
    End If
    
    If IsDarkColor(UserControl.BackColor) Then
        HotItembckColor = ShiftColor(UserControl.BackColor, vbWhite, 200)
        NormalTextColor = vbWhite
    Else
        HotItembckColor = ShiftColor(UserControl.BackColor, vbBlack, 240)
        NormalTextColor = vbBlack
    End If
    
    m_LineColor = ShiftColor(NormalTextColor, UserControl.BackColor, 50)
    
    TextHeight = UserControl.TextHeight("Àj")
    
    lWidth = UserControl.ScaleWidth
    If ucScrollbar1.Visible Then
        lWidth = ucScrollbar1.Left - 2 * DpiFactor
    Else
        If m_SelectionRadius > 0 Then
            lWidth = lWidth - lMargin
            lLeft = lMargin
        End If
    End If
    
    lTop = ListTop - ucScrollbar1.Value '  (m_BorderRadius / 2) - ucScrollbar1.Value
    
    For i = 0 To mListCount - 1
        With m_Items(i)
            If .Visible Then
                If lTop + IH > 0 Then
                
                    SetRect Rect, 0, lTop, lWidth, lTop + IH
                    

                                        
                    If i = m_HotItem Or i = m_Index Then
                        If bMouseDown Or i = m_Index Then
                            oColor = m_SelectionColor
                        Else
                            If .BackColor <> CLR_NONE Then
                                If IsDarkColor(.BackColor) Then
                                    oColor = ShiftColor(.BackColor, vbWhite, 150)
                                Else
                                    oColor = ShiftColor(.BackColor, vbBlack, 200)
                                End If
                            Else
                                oColor = HotItembckColor
                            End If
                        End If
                        
                        If m_SelectionRadius = 0 Then
                            FillRectangle hdc, Rect, oColor
                        Else
                            RoundRectPlus hdc, Rect.Left + lMargin / 2, Rect.Top, _
                                Rect.Right - Rect.Left - lMargin / 2, Rect.Bottom - Rect.Top, _
                                RGBtoARGB(oColor, 100), 0, m_SelectionRadius * DpiFactor
                        End If
                    Else
                        If .BackColor <> CLR_NONE Then
                            FillRectangle hdc, Rect, .BackColor
                        End If
                    End If

                    If .Separator Then
                        DrawLine hdc, lLeft, lTop + IH / 4, lWidth, lTop + IH / 4, m_LineColor
                        lTop = lTop + IH / 2
                    Else

                        If .TextColor <> CLR_NONE Then
                            UserControl.ForeColor = .TextColor
                        Else
                            If .Disabled Then
                                UserControl.ForeColor = ShiftColor(NormalTextColor, UserControl.BackColor, 120)
                            Else
                                If i = m_HotItem And bMouseDown Or i = m_Index Then
                                    
                                    UserControl.ForeColor = HighlightTextColor ' vbHighlightText
                                Else
                                    UserControl.ForeColor = NormalTextColor ' vbWindowText
                                End If
                            End If
                        End If
                        
                        If Not .SubMenu Is Nothing Then
                            DrawArrow hdc, i = m_HotItem And bMouseDown Or i = m_Index, Rect.Right - lCheckSize, lTop + IH / 2 - lCheckSize / 2, Rect.Right, lTop + IH / 2 + lCheckSize / 2
                        End If
                        
                        If m_UseMonocromeIcons Then
                            oColor = UserControl.ForeColor
                        Else
                            oColor = CLR_NONE
                        End If
                        
                        If m_DrawSeparatorLine Then
                            DrawLine hdc, 0, lTop + IH, lWidth, lTop + IH, m_LineColor
                        End If
        
                        If .IsGroup Then
                        
                            Rect.Left = lMargin * 2
                            
                            If cImageList.Count And .Icon > -1 Then
                                ImageListDraw hdc, .Icon + 1, Rect.Left, lTop + IH / 2 - ImgLstHeight / 2, ImgLstWidth, ImgLstHeight, oColor, .Disabled
                                Rect.Left = Rect.Left + ImgLstWidth + lMargin
                            End If
                            
                            UserControl.FontBold = True
                            DrawText hdc, StrPtr(.Text), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
                            UserControl.FontBold = False
                            RaiseEvent ItemCustomDraw(i, hdc, 0, lTop, lWidth, IH)
                        Else
                            Rect.Left = lMargin * 2
                            If .IsChild Then Rect.Left = Rect.Left + lMargin * 2
                            
                            If m_MultiSelect Then
                                If .SubMenu Is Nothing Then
                                    DrawCheckBox hdc, .Checked, i = m_HotItem And bMouseDown Or i = m_Index, Rect.Left, lTop + IH / 2 - lCheckSize / 2, Rect.Left + lCheckSize, lTop + IH / 2 + lCheckSize / 2
                                End If
                                Rect.Left = Rect.Left + lCheckSize + lMargin
                            End If
                            
                            If cImageList.Count Then
                                If .Icon > -1 Then ImageListDraw hdc, .Icon + 1, Rect.Left, lTop + IH / 2 - ImgLstHeight / 2, ImgLstWidth, ImgLstHeight, oColor, .Disabled
                                Rect.Left = Rect.Left + ImgLstWidth + lMargin
                            End If
                            
                            lPos = InStr(.Text, vbTab)
                            If lPos Then
                                sText = Left$(.Text, lPos)
                                sShortcut = Mid$(.Text, lPos)
                            Else
                                sText = .Text
                            End If
                            
                            Rect.Right = Rect.Right - lMargin - lCheckSize / 2
                            If Len(.SubText) And m_SubTitleVisible Then
                                Rect.Bottom = Rect.Bottom - TextHeight
                                DrawText hdc, StrPtr(sText), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
                                
                                If .SubTextColor <> CLR_NONE Then
                                     UserControl.ForeColor = .SubTextColor
                                Else
                                    If i <> m_Index Then UserControl.ForeColor = ShiftColor(UserControl.ForeColor, UserControl.BackColor, 120)
                                End If
                                If lPos Then DrawText hdc, StrPtr(sShortcut), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT
                                Rect.Bottom = Rect.Bottom + TextHeight * 2
                            
                                mFontSize = UserControl.Font.Size
                                UserControl.Font.Size = mFontSize - 1
                                DrawText hdc, StrPtr(.SubText), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
                                UserControl.Font.Size = mFontSize
                            Else
                                If .ItemBold Then UserControl.FontBold = True
                                DrawText hdc, StrPtr(sText), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
                                If lPos Then
                                    PrevForeColor = UserControl.ForeColor
                                    If .SubTextColor <> CLR_NONE Then
                                        UserControl.ForeColor = .SubTextColor
                                    Else
                                        UserControl.ForeColor = ShiftColor(PrevForeColor, UserControl.BackColor, 120)
                                    End If
                                    DrawText hdc, StrPtr(sShortcut), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT
                                    UserControl.ForeColor = PrevForeColor
                                End If
                                
                                If Len(Text1.Text) And i <> m_Index Then
                                    Dim TempStr As String, TempRect As Rect
                                    lPos = InStr(UCase(sText), UCase(Text1.Text))
                                    
                                    If lPos Then
                                        TempStr = Left$(sText, lPos - 1)
                                        UserControl.ForeColor = m_FindTextColor
                                        TempRect = Rect
                                        DrawText hdc, StrPtr(TempStr), -1, TempRect, DT_SINGLELINE Or DT_VCENTER Or DT_CALCRECT
                                        
                                        TempStr = Mid$(sText, lPos, Len(Text1.Text))
                                        Rect.Left = Rect.Left + TempRect.Right - TempRect.Left
                                        DrawText hdc, StrPtr(TempStr), -1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
        
                                    End If
                                End If
                                If .ItemBold Then UserControl.FontBold = False
                            End If

                        End If
                        RaiseEvent ItemCustomDraw(i, hdc, 0, lTop, lWidth, IH)
                        lTop = lTop + IH
                
                    End If
                Else
                    If .Separator Then
                        lTop = lTop + IH / 2
                    Else
                        lTop = lTop + IH
                    End If
                End If
            End If
            
            If lTop > UserControl.ScaleHeight Then Exit For
        End With
    Next
    
    'Header
    SetRect Rect, 0, 0, UserControl.ScaleWidth, ListTop
    FillRectangle hdc, Rect, UserControl.BackColor
    'optional
    If ucScrollbar1.Value > 0 Then
        SetRect Rect, 0, ListTop, UserControl.ScaleWidth, lMargin / 2
        FillRectGradient hdc, Rect, RGBtoARGB(UserControl.BackColor, 100), 0
    End If
    
    'Footer
    SetRect Rect, 0, UserControl.ScaleHeight - lMargin, UserControl.ScaleWidth, UserControl.ScaleHeight
    FillRectangle hdc, Rect, UserControl.BackColor
    'optional
    If ucScrollbar1.Value < ucScrollbar1.Max Then
        SetRect Rect, 0, UserControl.ScaleHeight - lMargin * 1.5, UserControl.ScaleWidth, lMargin / 2
        FillRectGradient hdc, Rect, 0, RGBtoARGB(UserControl.BackColor, 100)
    End If
    
    If m_ShowFilter Then
        If isSegoeFontInstaled Then
            RoundRectPlus hdc, Text1.Left - lMargin, Text1.Top - lMargin, Text1.Width + lMargin * 2 + lCheckSize, Text1.Height + lMargin * 2, RGBtoARGB(UserControl.BackColor, 100), RGBtoARGB(m_SelectionColor, 100), BR / 2
            SetRect Rect, Text1.Left + Text1.Width, Text1.Top - lMargin, UserControl.ScaleWidth, Text1.Top + Text1.Height + lMargin
            DrawGlass UserControl.hdc, Rect, NormalTextColor
        Else
            RoundRectPlus hdc, Text1.Left - lMargin, Text1.Top - lMargin, Text1.Width + lMargin * 2 + lCheckSize, Text1.Height + lMargin * 2, RGBtoARGB(UserControl.BackColor, 100), RGBtoARGB(m_SelectionColor, 100), BR / 2
        End If
    End If
    
    If isWinChild Then
        'oColor = IIf(bHasFocus, m_SelectionColor, m_LineColor)
        'RoundRectPlus hdc, DpiFactor / 2, DpiFactor / 2, UserControl.ScaleWidth - DpiFactor * 2, UserControl.ScaleHeight - DpiFactor * 2, 0, RGBtoARGB(oColor, 100), BR
        
    End If
    
    'RaiseEvent PostPaint(hdc)
  
    UserControl.Refresh
End Sub


Private Sub FillRectGradient(hdc As Long, Rect As Rect, ByVal Color1 As Long, Color2 As Long)
    Dim hBrush As Long, hGraphics As Long
    GdipCreateFromHDC hdc, hGraphics
    GdipCreateLineBrushFromRectWithAngleI Rect, Color1, Color2, 90, 0, &H3, hBrush
    GdipFillRectangleI hGraphics, hBrush, Rect.Left, Rect.Top, Rect.Right, Rect.Bottom
    GdipDeleteBrush hBrush
    GdipDeleteGraphics hGraphics
End Sub

Private Sub FillRectangle(hdc As Long, Rect As Rect, ByVal Color As OLE_COLOR)
    Dim hBrush As Long
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    hBrush = CreateSolidBrush(Color)
    FillRect hdc, Rect, hBrush
    Call DeleteObject(hBrush)
End Sub

Private Sub DrawGlass(hdc, Rect As Rect, Color As Long)
    Dim iFont As iFont
    Dim hPrevFont As Long
    Dim PrevForeColor As Long
    
    Set iFont = m_SegoeFont
    m_SegoeFont.Size = 18 / 1.5
    hPrevFont = SelectObject(hdc, iFont.hFont)
    PrevForeColor = UserControl.ForeColor
    UserControl.ForeColor = Color
    DrawText hdc, StrPtr(ChrW2(&HE094)), 1, Rect, DT_SINGLELINE Or DT_VCENTER
    Call SelectObject(hdc, hPrevFont)
    UserControl.ForeColor = PrevForeColor
End Sub

Private Sub DrawArrow(hdc As Long, ByVal bHighlight As Boolean, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    Dim iFont As iFont
    Dim hPrevFont As Long
    Dim Rect As Rect
    'Dim PrevForeColor As Long
    Dim sChr As String

    If isSegoeFontInstaled Then
        m_SegoeFont.Size = UserControl.Font.Size
        Set iFont = m_SegoeFont
        SetRect Rect, X1, Y1, X2, Y2
        sChr = ChrW2(&HE970)
    Else
        m_Wingdings3.Size = UserControl.Font.Size
        Set iFont = m_Wingdings3
        sChr = ChrW2(&H7D)
    End If
    
    hPrevFont = SelectObject(hdc, iFont.hFont)
    SetRect Rect, X1, Y1, X2, Y2
    DrawText hdc, StrPtr(sChr), 1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    Call SelectObject(hdc, hPrevFont)
    
End Sub

Private Sub DrawCheckBox(hdc As Long, Value As Boolean, ByVal bHighlight As Boolean, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    Dim iFont As iFont
    Dim hPrevFont As Long
    Dim Rect As Rect
    Dim PrevForeColor As Long
    Dim sChr As String
    
    PrevForeColor = UserControl.ForeColor
    
    If isSegoeFontInstaled Then
        m_SegoeFont.Size = 12
        Set iFont = m_SegoeFont
        If Value Then
            If bHighlight = False Then UserControl.ForeColor = m_SelectionColor
            sChr = ChrW2(&HE005)
        Else
            sChr = ChrW2(&HE003)
        End If
    Else
        m_Wingdings2.Size = 14
        Set iFont = m_Wingdings2
        If Value Then
            If bHighlight = False Then UserControl.ForeColor = m_SelectionColor
            sChr = ChrW2(&H52)
        Else
            sChr = ChrW2(&HA3)
        End If
    End If

    SetRect Rect, X1, Y1, X2, Y2
    hPrevFont = SelectObject(hdc, iFont.hFont)
    DrawText hdc, StrPtr(sChr), 1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    UserControl.ForeColor = PrevForeColor
    Call SelectObject(hdc, hPrevFont)
    
End Sub

Public Function GetWindowDpi() As Single
    GetWindowDpi = DpiFactor
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function
'*2
Public Sub ShowList(Left As Long, Top As Long, Optional BtnLeft As Long, Optional BtnTop As Long)
    Dim i As Long, WinHeight As Long
    Dim ScreenHeight As Long, ScreenWidth As Long
    Dim oCtl As Object
    
    If isWinChild Then Exit Sub
    If bShow Then HideList

    bHideSubMenu = False
    bShowSubMenu = False
    Timer1.Tag = -1
    m_HotItem = -1
    bKillFocus = False
    m_BtnTop = 0
    m_BtnLeft = 0
    
    PrevFocus = GetFocus
    m_PhWnd = UserControl.ContainerHwnd
    hAnsestor = GetAncestor(m_PhWnd, GA_ROOT)
    If SrcTextBox = 0 Then
        PutFocus UserControl.hWnd
    End If
    Text1.Text = vbNullString
    
    On Error Resume Next '---------Find the form to capture the events Unload and Resize
        'SubClass hAnsestor is not safe when is type menu/submenu
        Set oCtl = UserControl.Extender.Container
        Do
            If TypeOf oCtl Is Form Then
                Set Frm = oCtl
                Exit Do
            Else
                Set oCtl = oCtl.Parent
                If Err.Number Then Exit Do
            End If
        Loop
    On Error GoTo 0 '----------------

    If m_ShowFilter Then
        Set Text1.Font = UserControl.Font
        Text1.Visible = True
        PutFocus Text1.hWnd
        Text1.Height = UserControl.Font.Size * 1.5 * DpiFactor
        lHeadHeight = Text1.Height + lMargin * 4
        Text1.BackColor = UserControl.BackColor
        If IsDarkColor(Text1.BackColor) Then
            Text1.ForeColor = vbWhite
        Else
            Text1.ForeColor = vbBlack
        End If
    Else
        Text1.Visible = False
        lHeadHeight = lMargin
    End If

    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    
    Call RegisterHotsKeys
    
    WinHeight = lHeadHeight + lMargin '< footer
      
    For i = 0 To mListCount - 1 'Calculate height
        With m_Items(i)
            If .Separator Then
                WinHeight = WinHeight + IH / 2
            Else
                WinHeight = WinHeight + IH
            End If
            If WinHeight > ScreenHeight Then
                WinHeight = ScreenHeight
                Exit For
            End If
        End With
    Next
 
    If m_MaxVisibleItems > 0 Then
        m_MaxListHeight = (lHeadHeight + IH * m_MaxVisibleItems + lMargin)
    End If
    
    '---
    If m_MaxListHeight > 0 Then If WinHeight > m_MaxListHeight Then WinHeight = m_MaxListHeight
    If Top + WinHeight > ScreenHeight Then
        If BtnTop > 0 Then
            m_BtnTop = BtnTop
            
            Top = (m_BtnTop - WinHeight)
        Else
            Top = (ScreenHeight - WinHeight)
        End If
    End If
    If Top < 0 Then Top = 0
    
    '----
      
    If Left + (m_ListWidth * DpiFactor) > ScreenWidth Then
        If BtnLeft > 0 Then
            m_BtnLeft = BtnLeft
            Left = BtnLeft - (m_ListWidth * DpiFactor)
        Else
            Left = ScreenWidth - (m_ListWidth * DpiFactor)
        End If
    End If
    If Left < 0 Then Left = 0
    '----

    m_Left = Left
    SetParent UserControl.hWnd, 0
    SetWindowPos UserControl.hWnd, HWND_TOPMOST, Left, Top, m_ListWidth * DpiFactor, WinHeight, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    PutFocus UserControl.hWnd


    If m_ShowFilter Then
        Text1.Move lMargin * 2, lMargin * 2, UserControl.ScaleWidth - lMargin * 4 - 16 * DpiFactor
    End If

    ListTop = lHeadHeight
    ListHeight = UserControl.ScaleHeight - lHeadHeight - lMargin
    
    
    With ucScrollbar1
        .Max = IH * (mListCount - SeparatorsCount) + (IH / 2) * SeparatorsCount - ListHeight
       
        If .Max > 0 Then
        
            Select Case m_ListDesign
                Case GoogleDesign
                    .Width = 8 * DpiFactor
                    .ShowButtons = False
                    .Style = sCustomDraw
                Case MicrosoftDesign
                    .Width = 16 * DpiFactor
                    .ShowButtons = True
                    .Style = sThemed
            End Select
           
            .SmallChange = IH * 3
            ' Sugerencia de JB
            '.LargeChange = (ListHeight / (.Max + ListHeight)) * .Max
            .LargeChange = ListHeight - IH
            'If .LargeChange * 100 / .Max < 10 Then .LargeChange = .Max * 10 / 100
            ' Fin Sugerencia de JB
            .WheelChange = IH * 3
            .TrackMouseWheelOnHwnd UserControl.hWnd
            .Visible = True
            
        End If
    End With
    
    UserControl_Resize
    cShadow.BackColor = UserControl.BackColor
    If m_BorderVisible Then
        cShadow.ShowBorders UserControl.hWnd, isWinChild
    End If

    If m_Index > -1 Then FocusItem m_Index
    
    Call Draw
     
    'if ShowList is call on MouseDown, whait to MouseUp
    If (GetAsyncKeyState(vbLeftButton) < 0) Or (GetAsyncKeyState(vbRightButton) < 0) Then
        DoEvents
        Do While (GetAsyncKeyState(vbLeftButton) < 0) Or (GetAsyncKeyState(vbRightButton) < 0)
        Loop
    Else
        'Timer1.Enabled = True
       ' bTimerSkipRoutines = False
    End If

    Timer1.Interval = 1
    If bShow = False Then
        With cSubClass
            If .ssc_Subclass(UserControl.hWnd, , , Me) Then
                .ssc_AddMsg UserControl.hWnd, WM_HOTKEY, MSG_AFTER
                .ssc_AddMsg UserControl.hWnd, WM_MOUSELEAVE, MSG_BEFORE
                .ssc_AddMsg UserControl.hWnd, WM_IME_SETCONTEXT, MSG_BEFORE
            End If
        End With
    End If
    bShow = True
End Sub

'*1
Public Sub HideList()
    If bShow = False Then Exit Sub
    bShow = False
    
    Set Frm = Nothing

    If Not ChildMenu Is Nothing Then
        ChildMenu.HideList
        Set ChildMenu.ParentMenu = Nothing
        
        Set ChildMenu = Nothing
    End If
    
    Timer1.Interval = 0

    Call UnRegisterHotsKeys
    
'    If ParentMenu Is Nothing Then cSubClass.ssc_UnSubclass hAnsestor
    cSubClass.ssc_UnSubclass UserControl.hWnd
    
    hAnsestor = 0

    ShowWindow UserControl.hWnd, 0
    SetParent UserControl.hWnd, m_PhWnd
    
    cShadow.HideBorders
    
    'EnableWindow hAnsestor, 1
    'PutFocus PrevFocus

    If Not ParentMenu Is Nothing Then
        ParentMenu.HideList
        Set ParentMenu = Nothing
    End If
    RaiseEvent HideList
End Sub

Friend Sub UnRegisterHotsKeys()
    Call UnregisterHotKey(UserControl.hWnd, vbKeyDown)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyUp)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyReturn)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyEscape)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyHome)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyEnd)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyPageUp)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyPageDown)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyTab)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyLeft)
    Call UnregisterHotKey(UserControl.hWnd, vbKeyRight)
End Sub


Friend Sub RegisterHotsKeys()
    Call RegisterHotKey(UserControl.hWnd, vbKeyDown, 0, vbKeyDown)
    Call RegisterHotKey(UserControl.hWnd, vbKeyUp, 0, vbKeyUp)
    Call RegisterHotKey(UserControl.hWnd, vbKeyReturn, 0, vbKeyReturn)
    Call RegisterHotKey(UserControl.hWnd, vbKeyEscape, 0, vbKeyEscape)
    Call RegisterHotKey(UserControl.hWnd, vbKeyHome, 0, vbKeyHome)
    Call RegisterHotKey(UserControl.hWnd, vbKeyEnd, 0, vbKeyEnd)
    Call RegisterHotKey(UserControl.hWnd, vbKeyPageUp, 0, vbKeyPageUp)
    Call RegisterHotKey(UserControl.hWnd, vbKeyPageDown, 0, vbKeyPageDown)
    Call RegisterHotKey(UserControl.hWnd, vbKeyTab, 0, vbKeyTab)
    Call RegisterHotKey(UserControl.hWnd, vbKeyLeft, 0, vbKeyLeft)
    Call RegisterHotKey(UserControl.hWnd, vbKeyRight, 0, vbKeyRight)

End Sub


Private Sub DrawLine(lpDC As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color As Long)
    Dim PT      As POINTAPI
    Dim hPen    As Long
    Dim hPenOld As Long
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, X, Y, PT)
    Call LineTo(lpDC, X2, Y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub Frm_Resize()
    HideList
End Sub

Private Sub Frm_Unload(Cancel As Integer)
    HideList
End Sub

'*3
Private Sub Text1_Change()
    Dim i As Long
    Dim lHeight As Long
    Dim ParentIndex As Long
    Dim lWidth As Long
    
    bFiltered = Len(Text1) > 0
    m_Index = -1
    
    VisibleItemsCount = 0
    For i = 0 To mListCount - 1
        With m_Items(i)
            If .IsGroup Then
                .Visible = False
                ParentIndex = i
            Else
                If InStr(UCase(.Text), UCase(Text1.Text)) Then

                    lHeight = lHeight + IH
                    
                    If .IsChild And m_Items(ParentIndex).Visible = False Then
                        m_Items(ParentIndex).Visible = True
                        lHeight = lHeight + IH
                    End If
                    .Visible = True
                    VisibleItemsCount = VisibleItemsCount + 1
                Else
                    .Visible = False
                End If
            End If
        End With
    Next
    
    If isWinChild Then
        If lHeight < ucScrollbar1.Height Then
            ucScrollbar1.Visible = False
            ucScrollbar1.TrackMouseWheelOnHwndStop
        Else
            With ucScrollbar1
                .Value = 0
                .Max = lHeight - ucScrollbar1.Height
                .SmallChange = IH
                .LargeChange = ListHeight - IH
                .Visible = True
                .TrackMouseWheelOnHwnd UserControl.hWnd
            End With
        End If
        Draw
        Exit Sub
    End If
    
    lWidth = m_ListWidth * Screen.TwipsPerPixelX * DpiFactor
   
    If lHeight <= ucScrollbar1.Height Then
        UserControl.Size lWidth, (lHeadHeight + lHeight + lMargin) * Screen.TwipsPerPixelY
        ucScrollbar1.Visible = False
        ucScrollbar1.TrackMouseWheelOnHwndStop

    Else
        If (ListTop + lHeight + lMargin) > m_MaxListHeight And m_MaxListHeight > 0 Then
            UserControl.Size lWidth, m_MaxListHeight * Screen.TwipsPerPixelY
        Else
            UserControl.Size lWidth, ListHeight * Screen.TwipsPerPixelY
          
        End If

        With ucScrollbar1
            .Value = 0
            .Max = lHeight - ucScrollbar1.Height
            .SmallChange = IH
            .LargeChange = ListHeight - IH
            'If .Max > 0 Then If .LargeChange * 100 / .Max < 10 Then .LargeChange = .Max * 10 / 100
            .Visible = True
            .TrackMouseWheelOnHwnd UserControl.hWnd
        End With
    End If
  
    If m_BtnTop > 0 Then
        SetWindowPos UserControl.hWnd, HWND_TOPMOST, m_Left, (m_BtnTop * DpiFactor) - UserControl.ScaleHeight, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
    End If
    bKillFocus = False
    If m_BorderVisible Then cShadow.Refresh
 
    If ucScrollbar1.Value > 0 Then
        ucScrollbar1.Value = 0
    Else
        Draw
    End If
End Sub

Private Sub FocusItem(Index As Long)
    Dim i As Long
    Dim mHeight As Long
    
    For i = 0 To mListCount - 1
        If m_Items(i).Visible Then
            If Not m_Items(i).Separator Then
                mHeight = mHeight + IH
            Else
                mHeight = mHeight + IH / 2
            End If
            If Index = i Then
                If mHeight + IH > ucScrollbar1.Value + ListHeight Then
                    ucScrollbar1.Value = mHeight - ListHeight  '+ IH * 2
                ElseIf mHeight - IH < ucScrollbar1.Value Then
                   ucScrollbar1.Value = mHeight - IH
                End If
                Draw
                Exit For
            End If
        End If
    Next
End Sub

Private Function IsItemSelectable(Index As Long) As Boolean
    If m_Items(Index).Visible And Not m_Items(Index).Disabled And _
        Not m_Items(Index).IsGroup And Not m_Items(Index).Separator Then
        IsItemSelectable = True
    End If
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim ItemPerPage As Long

    Select Case KeyCode
        Case vbKeyUp
            KeyCode = 0
            For i = m_Index - 1 To 0 Step -1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            
        Case vbKeyDown
            KeyCode = 0
            For i = m_Index + 1 To mListCount - 1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            
        Case vbKeyHome
            KeyCode = 0
            For i = 0 To mListCount - 1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            
        Case vbKeyEnd
            KeyCode = 0
            For i = mListCount - 1 To 0 Step -1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            
        Case vbKeyEscape
            If Not isWinChild Then
                KeyCode = 0
                If Not ParentMenu Is Nothing Then
                    Text1_KeyDown vbKeyLeft, 0
                Else
                    HideList
                End If
            End If
            
        Case vbKeyReturn
         
            KeyCode = 0
            If m_MultiSelect Then
                If m_Index <> -1 Then
                    m_Items(m_Index).Checked = Not m_Items(m_Index).Checked
                    Draw
                    RaiseEvent ItemClick
                End If
            Else 'If Not isWinChild Then
                If m_Index <> -1 Then
                    If Not isWinChild Then HideList
                    RaiseEvent ItemClick
                    
                    If SrcTextBox <> 0 Then
                        bTimerSkipRoutines = True
                        Dim sText As String
                        sText = m_Items(m_Index).Text
                        TextBox_SetText SrcTextBox, sText
                        TextBox_SelStart SrcTextBox, Len(sText)
                        bTimerSkipRoutines = False
                        RaiseEvent PostAutoComplete
                    End If
                Else
                    If SrcTextBox <> 0 Then
                        HideList
                        SendMessage SrcTextBox, WM_CHAR, vbKeyReturn, 0&
                    Else
                        Beep
                    End If
                End If
            End If
            
        Case vbKeyPageDown
            KeyCode = 0
            ItemPerPage = ListHeight \ (ItemHeight * DpiFactor)
            For i = m_Index + ItemPerPage To mListCount - 1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            Text1_KeyDown vbKeyEnd, 0
            
        Case vbKeyPageUp
            KeyCode = 0
            ItemPerPage = ListHeight \ (ItemHeight * DpiFactor)
            
            For i = m_Index - ItemPerPage To 0 Step -1
                If IsItemSelectable(i) Then
                    m_Index = i
                    FocusItem m_Index
                    RaiseEvent IndexChange
                    Exit For
                End If
            Next
            Text1_KeyDown vbKeyHome, 0
            
        Case vbKeyTab
            If Not isWinChild Then HideList

       Case vbKeyLeft
            KeyCode = 0
           If Not ParentMenu Is Nothing Then
               GiveCommand
               ParentMenu.TakeCommand
               Set ParentMenu = Nothing
               HideList
           End If
 
       Case vbKeyRight
           KeyCode = 0

           If m_Index <> -1 Then
               
               If Not m_Items(m_Index).SubMenu Is Nothing Then
                    Dim Rect As Rect
                    GetWindowRect UserControl.hWnd, Rect
                    Set m_Items(m_Index).SubMenu.ParentMenu = UserControl.Extender
                    Set ChildMenu = m_Items(m_Index).SubMenu
                    GiveCommand
                    m_Items(m_Index).SubMenu.ListIndex = -1
                    m_Items(m_Index).SubMenu.ShowList Rect.Right, Rect.Top + GetItemTop(m_Index) - lHeadHeight, Rect.Left
               End If
           End If
        'Case vbKeyF4
            'If m_ListStyle = StyleDropDown Then HideList
            
    End Select
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
    Dim hwin As Long
    Dim PT As POINTAPI

    '------------------
    If bShowSubMenu Then
        TickCount = TickCount + 1
        If TickCount > 30 Then
            bShowSubMenu = False
            TimerShowMenu
        End If
    End If
    
    '------------------
    If bHideSubMenu Then
        TickCount = TickCount + 1
        If TickCount > 30 Then
            bHideSubMenu = False
            TimerHideMenu
        End If
    End If
    '------------------

    If bTimerSkipRoutines Then
        If ChildMenu Is Nothing Then
            bTimerSkipRoutines = False
        Else
            Exit Sub
        End If
    End If
    
    If isWinChild Then
        GetCursorPos PT
        hwin = WindowFromPoint(PT.X, PT.Y)
        If hwin <> UserControl.hWnd Then
            m_HotItem = -1
            Timer1.Interval = 0
            Draw
        End If
    Else

        If (GetAsyncKeyState(vbLeftButton) < 0) Or (GetAsyncKeyState(vbRightButton) < 0) Then
            GetCursorPos PT
            hwin = WindowFromPoint(PT.X, PT.Y)
           
            If hwin <> UserControl.hWnd And hwin <> Text1.hWnd And hwin <> ucScrollbar1.hWnd Then
                   
                If Not ParentMenu Is Nothing Then
                    If hwin = ParentMenu.hWnd Then Exit Sub
                End If
           
                'If GetCapture <> UserControl.hWnd Then HideList
                HideList
            End If
        End If
        
        If bKillFocus Then HideList
        ' esto es para que se pueda mostrar u ocultar la lista presionando un boton
        ' con este tiempo de puede consultar si bShow = true
    End If
End Sub

Private Sub TimerHideMenu()
    'TimerHideMenu.Enabled = False
    If Not ChildMenu Is Nothing Then
        Set ChildMenu.ParentMenu = Nothing
        ChildMenu.HideList
        Set ChildMenu = Nothing
        TakeCommand
        m_Index = -1
        Draw
    End If
End Sub

Private Sub TimerShowMenu()
    Dim Rect As Rect
    If Timer1.Tag = m_Index Then Exit Sub
    GetWindowRect UserControl.hWnd, Rect
    m_Index = Timer1.Tag
    Draw
    Set m_Items(m_Index).SubMenu.ParentMenu = UserControl.Extender
    Set ChildMenu = m_Items(m_Index).SubMenu

    GiveCommand
    
    m_Items(m_Index).SubMenu.ListIndex = -1
    m_Items(m_Index).SubMenu.ShowList Rect.Right, Rect.Top + GetItemTop(m_Index) - lHeadHeight, Rect.Left


End Sub

Private Sub ucScrollbar1_Change()
    Draw
    If ucScrollbar1.Value = ucScrollbar1.Max Then
        RaiseEvent ScrollBottom
    End If
End Sub

Private Sub ucScrollbar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
  
End Sub

Private Sub ucScrollbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
  
End Sub

Private Sub ucScrollbar1_Scroll()
    Draw
End Sub


Private Sub ucScrollbar1_OnPaint(ByVal lHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal eState As sbOnPaintPartStateCts)
    Dim Rect As Rect
    Dim Radius As Long, lColor As Long
    Dim mDC As Long, hBmp As Long, OldBmp As Long
    
    Select Case ePart ' = ppNullTrack
        Case ppThumb
            Radius = X2 - X1 - 1

            If IsDarkColor(UserControl.BackColor) Then
                If eState = ppsHot Then
                    lColor = ShiftColor(UserControl.BackColor, vbWhite, 180)
                ElseIf eState = ppsPressed Then
                    lColor = ShiftColor(UserControl.BackColor, vbWhite, 150)
                Else
                    lColor = ShiftColor(UserControl.BackColor, vbWhite, 200)
                End If
            Else
                If eState = ppsHot Then
                    lColor = ShiftColor(UserControl.BackColor, vbBlack, 220)
                ElseIf eState = ppsPressed Then
                    lColor = ShiftColor(UserControl.BackColor, vbBlack, 180)
                Else
                    lColor = ShiftColor(UserControl.BackColor, vbBlack, 200)
                End If
                
            End If

            
            mDC = CreateCompatibleDC(0)
            hBmp = CreateCompatibleBitmap(lHdc, X2 - X1, Y2 - Y1)
            OldBmp = SelectObject(mDC, hBmp)
            
            RoundRectPlus mDC, -DpiFactor, -DpiFactor, X2 - X1 + DpiFactor, Y2 - Y1 + DpiFactor, RGBtoARGB(UserControl.BackColor, 100), 0, 0
            RoundRectPlus mDC, 0, 0, X2 - X1 - DpiFactor, Y2 - Y1 - DpiFactor, RGBtoARGB(lColor, 100), 0, Radius
            
            BitBlt lHdc, X1, Y1, X2 - X1, Y2 - Y1, mDC, 0, 0, vbSrcCopy
            DeleteObject SelectObject(mDC, OldBmp)
            DeleteDC mDC
        Case Else
            SetRect Rect, X1, Y1, X2, Y2
            FillRectangle lHdc, Rect, UserControl.BackColor
    End Select
    
End Sub



Private Sub UserControl_AmbientChanged(PropertyName As String)
    If isWinChild Then
        If PropertyName = "BackColor" Then
            If m_BorderVisible Then cShadow.Refresh
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    bHasFocus = True
    If isWinChild Then
        With cShadow
            .ShadowOpacity = 25
            '.ShadowSize = 10
            .ShadowColor = m_SelectionColor
            .BorderColor = m_SelectionColor
            If m_BorderVisible Then .Refresh
        End With
        Draw
    End If
End Sub

Private Sub UserControl_ExitFocus()
    bHasFocus = False
    If isWinChild Then
        With cShadow
            .ShadowOpacity = 10
            '.ShadowSize = 2
            .ShadowColor = vbBlack
            .BorderColor = m_BorderColor
            If m_BorderVisible Then .Refresh
        End With
        If Extender.Visible Then Draw
        
    End If
End Sub

Private Sub UserControl_Hide()
    cShadow.HideBorders
End Sub

Private Sub UserControl_Initialize()
    Set cImageList = New Collection
    Set cSubClass = New clsSubClass
    Set cShadow = New ClsShadow  '<---GdiplusStartup
    cShadow.ShadowSize = 2
    ucScrollbar1.ShowButtons = False
    ucScrollbar1.Style = sCustomDraw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Text1_KeyDown(KeyCode, Shift)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'*1
Friend Sub TakeCommand()
    Set ChildMenu = Nothing
    bKillFocus = False
    RegisterHotsKeys
    cShadow.DrawBitmap UserControl.Image, UserControl.hWnd, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetParent UserControl.hWnd, m_PhWnd
    PutFocus UserControl.hWnd
    SetParent UserControl.hWnd, 0
    bKillFocus = False
    bTimerSkipRoutines = False
    SubMenuVisible = False
End Sub

Friend Sub GiveCommand()
    bTimerSkipRoutines = True
    SubMenuVisible = True
    UnRegisterHotsKeys
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim lPos As Long, sChar As String
    
    RaiseEvent KeyPress(KeyAscii)
    
    If m_ListStyle = StyleMenu Then
        For i = 0 To mListCount - 1
            With m_Items(i)
                If Not .Disabled And Not .IsGroup And Not .Separator Then
                    lPos = InStr(.Text, "&")
                    If lPos Then
                        sChar = Mid$(.Text, lPos + 1, 1)
                        If UCase$(Chr$(KeyAscii)) = UCase(sChar) Then
                            m_Index = i
                            HideList
                            RaiseEvent ItemClick
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Next

        Beep
        Exit Sub
    End If
    
    For i = m_Index + 1 To mListCount - 1
        With m_Items(i)
            If Not .Disabled And Not .IsGroup Then
                If UCase$(Chr$(KeyAscii)) = UCase(Left$(.Text, 1)) Then
                    m_Index = i
                    FocusItem m_Index
                    Draw
                    RaiseEvent IndexChange
                    Exit Sub
                End If
            End If
        End With
    Next
    

    For i = 0 To m_Index - 1
        With m_Items(i)
            If Not .Disabled And Not .IsGroup Then
                If UCase$(Chr$(KeyAscii)) = UCase(Left$(.Text, 1)) Then
                    m_Index = i
                    FocusItem m_Index
                    Draw
                    RaiseEvent IndexChange
                    Exit Sub
                End If
            End If
        End With
    Next
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Timer1.Enabled = False
    If Button = vbLeftButton Then
        If m_HotItem <> -1 Then
        
            If isWinChild Then
                m_Index = m_HotItem
                bMouseDown = True
                Draw
                RaiseEvent IndexChange
            Else
                If m_Items(m_HotItem).SubMenu Is Nothing Then
                    bMouseDown = True
                    Draw
                Else
                    bShowSubMenu = False
                    TimerShowMenu
                End If
            End If
        End If
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim HotItem As Long
    Dim lTop As Long
    Dim i As Long
    Dim ET As TRACKMOUSEEVENTTYPE

    If ucScrollbar1.Visible And X > ucScrollbar1.Left Then Exit Sub
    
    If isWinChild Then
        Timer1.Interval = 100 'para no subclasificar
    Else
        ET.cbSize = Len(ET)
        ET.hwndTrack = UserControl.hWnd
        ET.dwFlags = TME_LEAVE
        TrackMouseEvent ET
    End If
    
    HotItem = -1
    lTop = ListTop - ucScrollbar1.Value
    
    For i = 0 To mListCount - 1
        With m_Items(i)
            If .Visible Then
                If Y >= lTop And Y < lTop + IH Then
                    If Not .Separator And Not .IsGroup And Not .Disabled Then
                        HotItem = i
                        Exit For
                    Else
                        HotItem = -1
                    End If
                End If
                If Not .Separator Then
                    lTop = lTop + IH
                Else
                    lTop = lTop + IH / 2
                End If
                If lTop - ucScrollbar1.Value > UserControl.ScaleHeight Then Exit For
            End If
        End With
    Next
    
    If HotItem <> m_HotItem Then
        m_HotItem = HotItem

        If HotItem <> -1 Then
            If Not m_Items(HotItem).SubMenu Is Nothing Then
                bHideSubMenu = False
                bShowSubMenu = True
                TickCount = 0
                Timer1.Tag = HotItem
            Else
                bHideSubMenu = True
                bShowSubMenu = False
                TickCount = 0
                Timer1.Tag = -1
            End If
        End If
        
        Draw
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMouseDown = False Then Exit Sub
  
    bMouseDown = False
    If m_HotItem <> -1 Then
        m_Index = m_HotItem

        If m_MultiSelect Then

            m_Items(m_Index).Checked = Not m_Items(m_Index).Checked
            Draw
            RaiseEvent ItemClick
        Else
            If isWinChild Then
                Draw
                'RaiseEvent MouseUp(Button, Shift, X, Y)
                'Exit Sub
            Else
                If m_Items(m_Index).SubMenu Is Nothing Then
                    HideList
                End If
            End If
            
            'RaiseEvent IndexChange
            RaiseEvent ItemClick
            
            If SrcTextBox <> 0 Then
                bTimerSkipRoutines = True
                Dim sText As String
                sText = m_Items(m_Index).Text
                TextBox_SetText SrcTextBox, sText
                TextBox_SelStart SrcTextBox, Len(sText)
                bTimerSkipRoutines = False
                RaiseEvent PostAutoComplete
            End If

        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub LoadFontsIcons()
    Set m_SegoeFont = New StdFont
    m_SegoeFont.Name = "Segoe MDL2 Assets"
    
    If m_SegoeFont.Name = "Segoe MDL2 Assets" Then
        isSegoeFontInstaled = True
    Else
        Set m_SegoeFont = Nothing
        Set m_Wingdings2 = New StdFont
        Set m_Wingdings3 = New StdFont
        m_Wingdings2.Name = "Wingdings 2"
        m_Wingdings3.Name = "Wingdings 3"
        isSegoeFontInstaled = False
    End If
End Sub

Private Sub UserControl_InitProperties()
    DpiFactor = cShadow.GetWindowsDPI
    Me.ItemHeight = 32
    m_ListWidth = 250
    Me.BorderRadius = 10
    Me.ItemHeight = 32
    m_SelectionColor = vbHighlight
    m_LineColor = vbActiveBorder
    m_BorderColor = vbActiveBorder
    lMargin = 8 * DpiFactor
    lCheckSize = 16 * DpiFactor
    m_FindTextColor = ShiftColor(vbHighlight, UserControl.BackColor, 200)
    m_HotItem = -1
    m_BorderVisible = True
    
    Call LoadFontsIcons
    
    Text1.Text = vbNullString
    
    Dim i As Long
    For i = 1 To 7
        If i Mod 3 = 0 Then
            AddItem "-"
        Else
            AddItem "Item " & i, , , "Sub Text", i = 5, i = 3
        End If
    Next
    m_Index = 3


End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    DpiFactor = cShadow.GetWindowsDPI

    lMargin = 8 * DpiFactor
    lCheckSize = 16 * DpiFactor
    m_Index = -1
    
    Call LoadFontsIcons

    UserControl.DrawWidth = 1 * DpiFactor
    
    If Ambient.UserMode Then UserControl.Picture = Nothing
 
    With PropBag
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_SubTitleVisible = .ReadProperty("SubTitleVisible", False)
        m_ShowFilter = .ReadProperty("ShowFilter", False)
        Me.ItemHeight = .ReadProperty("ItemHeight", 32)
        m_MaxListHeight = .ReadProperty("MaxListHeight", 0)
        m_ListWidth = .ReadProperty("ListWidth", 250)
        m_ListDesign = .ReadProperty("ListDesign", GoogleDesign)
        m_MultiSelect = .ReadProperty("MultiSelect", False)
        m_UseMonocromeIcons = .ReadProperty("UseMonocromeIcons", False)
        Me.BorderRadius = .ReadProperty("BorderRadius", 10)
        m_DrawSeparatorLine = .ReadProperty("DrawSeparatorLine", False)
        m_ImagesRadius = .ReadProperty("ImagesRadius", 0)
        m_SelectionRadius = .ReadProperty("SelectionRadius", 0)
        Me.BackColor = .ReadProperty("BackColor", UserControl.BackColor)
        m_FindTextColor = .ReadProperty("FindTextColor", vbHighlight)
        m_SelectionColor = .ReadProperty("SelectionColor", vbHighlight)
        m_MaxVisibleItems = .ReadProperty("MaxVisibleItems", 0)
        'isWinChild = .ReadProperty("ListBoxStyle", False)
        m_LineColor = .ReadProperty("LineColor", vbActiveBorder)
        m_BorderColor = .ReadProperty("BorderColor", vbActiveBorder)
        m_BorderVisible = .ReadProperty("BorderVisible", True)
        UserControl.Enabled = .ReadProperty("Enabled", UserControl.Enabled)
        m_ListStyle = .ReadProperty("ListStyle", StyleDropDown)
        Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
    End With
    
    isWinChild = CBool(m_ListStyle = StyleListBox)
        
    If isWinChild Then
        m_HotItem = -1
        Text1.Text = vbNullString
        UserControl.Picture = Nothing
        If Ambient.UserMode = False Then
            Dim i As Long
            For i = 1 To 7
                If i Mod 3 = 0 Then
                    AddItem "-"
                Else
                    AddItem "Item " & i, , , "Sub Text", i = 5, i = 3
                End If
            Next
            m_Index = 3
            Update
        End If
    Else
        Extender.Visible = False
    End If
End Sub
'*2

Public Sub Update(Optional bResize As Boolean = True)
    Dim i As Long, WinHeight As Long
    
    If m_ShowFilter Then
        Set Text1.Font = UserControl.Font
        Text1.Visible = True
        Text1.Height = UserControl.Font.Size * 1.5 * DpiFactor
        lHeadHeight = Text1.Height + lMargin * 4
        Text1.BackColor = UserControl.BackColor
        If IsDarkColor(Text1.BackColor) Then
            Text1.ForeColor = vbWhite
        Else
            Text1.ForeColor = vbBlack
        End If
    Else
        Text1.Visible = False
        lHeadHeight = lMargin
    End If
    
    For i = 0 To mListCount - 1
        With m_Items(i)
            If .Separator Then
                WinHeight = WinHeight + IH / 2
            Else
                WinHeight = WinHeight + IH
            End If
        End With
    Next

    ListTop = lHeadHeight
    ListHeight = UserControl.ScaleHeight - lHeadHeight - lMargin
    
    With ucScrollbar1
        .Max = IH * (mListCount - SeparatorsCount) + (IH / 2) * SeparatorsCount - ListHeight

        If .Max > 0 Then
        
            Select Case m_ListDesign
                Case GoogleDesign
                    .Width = 8 * DpiFactor
                    .ShowButtons = False
                    .Style = sCustomDraw
                Case MicrosoftDesign
                    .Width = 16 * DpiFactor
                    .ShowButtons = True
                    .Style = sThemed
            End Select
           
            .SmallChange = IH * 3
            ' Sugerencia de JB
            '.LargeChange = (ListHeight / (.Max + ListHeight)) * .Max
            .LargeChange = ListHeight - IH
            'If .LargeChange * 100 / .Max < 10 Then .LargeChange = .Max * 10 / 100
            ' Fin Sugerencia de JB
            .WheelChange = IH * 3
            .TrackMouseWheelOnHwnd UserControl.hWnd
            .Visible = True
        Else
            .TrackMouseWheelOnHwndStop
            .Visible = False
        End If
    End With
    
    If bResize Then UserControl_Resize
    If m_BorderVisible Then cShadow.Refresh
    Draw

End Sub

Private Sub UserControl_Show()
    If isWinChild Then
        Draw
        If Ambient.UserMode Then
            If m_BorderVisible Then
                cShadow.ShowBorders UserControl.hWnd, isWinChild
            End If
        End If
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "SubTitleVisible", m_SubTitleVisible, False
        .WriteProperty "ShowFilter", m_ShowFilter, False
        .WriteProperty "ItemHeight", m_ItemHeight, 32
        .WriteProperty "MaxListHeight", m_MaxListHeight, 0
        .WriteProperty "ListWidth", m_ListWidth, 250
        .WriteProperty "ListDesign", m_ListDesign, GoogleDesign
        .WriteProperty "MultiSelect", m_MultiSelect, False
        .WriteProperty "UseMonocromeIcons", m_UseMonocromeIcons, False
        .WriteProperty "BorderRadius", m_BorderRadius, 10
        .WriteProperty "DrawSeparatorLine", m_DrawSeparatorLine, False
        .WriteProperty "ImagesRadius", m_ImagesRadius, 0
        .WriteProperty "SelectionRadius", m_SelectionRadius, 0
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "FindTextColor", m_FindTextColor, ShiftColor(vbHighlight, UserControl.BackColor, 200)
        .WriteProperty "SelectionColor", m_SelectionColor
        .WriteProperty "MaxVisibleItems", m_MaxVisibleItems, 0
        .WriteProperty "LineColor", m_LineColor, vbActiveBorder
        .WriteProperty "BorderColor", m_BorderColor, vbActiveBorder
        .WriteProperty "BorderVisible", m_BorderVisible, True
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "ListStyle", m_ListStyle, StyleDropDown
        .WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Dim BW As Long
    If Ambient.UserMode Or isWinChild Then
        With ucScrollbar1
            If m_ListDesign = GoogleDesign Then
                .Move UserControl.ScaleWidth - .Width - 2 * DpiFactor, ListTop, .Width, UserControl.ScaleHeight - ListTop - lMargin
            Else
                If isWinChild Then BW = DpiFactor
                    
                If m_ShowFilter Then
                    .Move UserControl.ScaleWidth - .Width - BW * 2, ListTop, .Width, UserControl.ScaleHeight - ListTop - BR / 2 - BW * 2
                Else
                    .Move UserControl.ScaleWidth - .Width - BW * 2, BR / 2 + BW, .Width, UserControl.ScaleHeight - BR - BW * 3
                End If
            End If
        End With
        
        If m_ShowFilter Then
  
            Text1.Move lMargin * 2, lMargin * 2, UserControl.ScaleWidth - lMargin * 4 - 16 * DpiFactor
        End If
        
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BR, BR)
        SetWindowRgn UserControl.hWnd, hRgn, True
        
        If isWinChild Then Call Update(False)
    Else
        UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
    End If
    
    
End Sub

Private Sub UserControl_Terminate()
    Dim i As Long

    Set ParentMenu = Nothing
    Set ChildMenu = Nothing
    Set cSubClass = Nothing

    If Not cImageList Is Nothing Then
        For i = 1 To cImageList.Count
            GdipDisposeImage CLng(cImageList(i))
        Next
    End If
    
    Set cImageList = Nothing
    Set cShadow = Nothing 'GdiplusShutdown
End Sub
'*-

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(new_Font As StdFont)
    With UserControl.Font
        .Name = new_Font.Name
        .Size = new_Font.Size
        .Bold = new_Font.Bold
        .Italic = new_Font.Italic
        .Underline = new_Font.Underline
        .Strikethrough = new_Font.Strikethrough
        .Weight = new_Font.Weight
        .Charset = new_Font.Charset
    End With
    PropertyChanged "Font"
    If isWinChild Then Draw
End Property

Public Property Get ListStyle() As eListStyle
    ListStyle = m_ListStyle
End Property

Public Property Let ListStyle(ByVal New_Value As eListStyle)

    If New_Value = StyleListBox Then
        isWinChild = True
        UserControl.Picture = Nothing
        If New_Value Then
            UserControl.Size 150 * Screen.TwipsPerPixelX, 200 * Screen.TwipsPerPixelY
        End If
    Else
        isWinChild = False
        UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
    End If
    m_ListStyle = New_Value
    
    PropertyChanged "ListStyle"
End Property

'Public Property Get ListBoxStyle() As Boolean
'    ListBoxStyle = isWinChild
'End Property
'
'Public Property Let ListBoxStyle(ByVal new_value As Boolean)
'
'    If Not isWinChild Then
'        isWinChild = new_value
'        UserControl.Picture = Nothing
'        If new_value Then
'            UserControl.Size 150 * Screen.TwipsPerPixelX, 200 * Screen.TwipsPerPixelY
'        End If
'    End If
'
'    PropertyChanged "ListBoxStyle"
'End Property
Public Function ItemHitest(Y As Long) As Long
    Dim Item As Long
    Dim lTop As Long
    Dim i As Long
    Item = -1
    For i = 0 To mListCount - 1
        With m_Items(i)
            If .Visible Then
                If Y >= lTop And Y < lTop + IH Then
                    If Not .Separator And Not .IsGroup And Not .Disabled Then
                        Item = i
                        Exit For
                    Else
                        Item = -1
                    End If
                End If
                If Not .Separator Then
                    lTop = lTop + IH
                Else
                    lTop = lTop + IH / 2
                End If
                If lTop - ucScrollbar1.Value > UserControl.ScaleHeight Then Exit For
            End If
        End With
    Next
    
    ItemHitest = Item
End Function

Public Property Get Text() As String
    If m_Index <> -1 Then
        Text = m_Items(m_Index).Text
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    UserControl.Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal New_Value As OLE_COLOR)
    m_LineColor = New_Value
    PropertyChanged "LineColor"
    If isWinChild Then Draw
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = New_Value
    PropertyChanged "BorderColor"
    cShadow.BorderColor = m_BorderColor
    If isWinChild Then Draw
End Property

Public Property Get BorderVisible() As Boolean
    BorderVisible = m_BorderVisible
End Property

Public Property Let BorderVisible(ByVal New_Value As Boolean)
    m_BorderVisible = New_Value
    PropertyChanged "BorderVisible"
    
    If Ambient.UserMode Then
        If m_BorderVisible Then
            If Extender.Visible Then cShadow.ShowBorders Me.hWnd, True
        Else
            cShadow.HideBorders
        End If
    End If
End Property


Public Property Get SelectionColor() As OLE_COLOR
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_Value As OLE_COLOR)
    m_SelectionColor = New_Value
    PropertyChanged "SelectionColor"
    If isWinChild Then Draw
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let FindTextColor(ByVal New_Value As OLE_COLOR)
    m_FindTextColor = New_Value
    PropertyChanged "FindTextColor"
    If isWinChild Then Draw
End Property

Public Property Get FindTextColor() As OLE_COLOR
    FindTextColor = m_FindTextColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    UserControl.BackColor = New_Value
    If Ambient.UserMode Then cShadow.BackColor = New_Value
    PropertyChanged "BackColor"
    If isWinChild Then Draw
End Property



Public Property Get SubTitleVisible() As Boolean
    SubTitleVisible = m_SubTitleVisible
End Property

Public Property Let SubTitleVisible(ByVal New_Value As Boolean)
    m_SubTitleVisible = New_Value
    PropertyChanged "SubTitleVisible"
    If isWinChild Then Draw
End Property

Public Property Get ShowFilter() As Boolean
    ShowFilter = m_ShowFilter
End Property

Public Property Let ShowFilter(ByVal New_Value As Boolean)
    m_ShowFilter = New_Value
    PropertyChanged "ShowFilter"
    If isWinChild Then Update True
End Property
'*1

Public Property Get ItemHeight() As Long
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_Value As Long)
    m_ItemHeight = New_Value
    IH = m_ItemHeight * DpiFactor
    PropertyChanged "ItemHeight"
    If isWinChild Then Draw
End Property

Public Property Get MaxListHeight() As Long
    MaxListHeight = m_MaxListHeight
End Property

Public Property Let MaxListHeight(ByVal New_Value As Long)
    m_MaxListHeight = New_Value
    PropertyChanged "MaxListHeight"
End Property

Public Property Get MaxVisibleItems() As Long
    MaxVisibleItems = m_MaxVisibleItems
End Property

Public Property Let MaxVisibleItems(ByVal New_Value As Long)
    m_MaxVisibleItems = New_Value
    PropertyChanged "MaxVisibleItems"
End Property

Public Property Get ListWidth() As Long
    ListWidth = m_ListWidth
End Property

Public Property Let ListWidth(ByVal New_Value As Long)
    m_ListWidth = New_Value
    PropertyChanged "ListWidth"
End Property

Public Property Get ListDesign() As eListDesign
    ListDesign = m_ListDesign
End Property

Public Property Let ListDesign(ByVal New_Value As eListDesign)
    m_ListDesign = New_Value
    PropertyChanged "ListDesign"
End Property

Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_Value As Boolean)
    m_MultiSelect = New_Value
    PropertyChanged "MultiSelect"
    If isWinChild Then Draw
End Property

Public Property Get UseMonocromeIcons() As Boolean
    UseMonocromeIcons = m_UseMonocromeIcons
End Property

Public Property Let UseMonocromeIcons(ByVal New_Value As Boolean)
    m_UseMonocromeIcons = New_Value
    PropertyChanged "UseMonocromeIcons"
    If isWinChild Then Draw
End Property

Public Property Get BorderRadius() As Integer
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Integer)
    m_BorderRadius = New_Value
    BR = m_BorderRadius * DpiFactor
    cShadow.BorderRadius = New_Value
    PropertyChanged "BorderRadius"
    If isWinChild Then Draw
End Property

Public Property Get DrawSeparatorLine() As Boolean
    DrawSeparatorLine = m_DrawSeparatorLine
End Property

Public Property Let DrawSeparatorLine(ByVal New_Value As Boolean)
    m_DrawSeparatorLine = New_Value
    PropertyChanged "DrawSeparatorLine"
    If isWinChild Then Draw
End Property

Public Property Get ImagesRadius() As Long
    ImagesRadius = m_ImagesRadius
End Property

Public Property Let ImagesRadius(ByVal New_Value As Long)
    m_ImagesRadius = New_Value
    PropertyChanged "ImagesRadius"
    If isWinChild Then Draw
End Property

Public Property Get SelectionRadius() As Long
    SelectionRadius = m_SelectionRadius
End Property

Public Property Let SelectionRadius(ByVal New_Value As Long)
    m_SelectionRadius = New_Value
    PropertyChanged "ImagesRadius"
    If isWinChild Then Draw
End Property

Public Property Get DragIcon() As IPictureDisp
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Property Get OLEDropMode() As ListPlusOLEDropModeConstants
    OLEDropMode = UserControl.OLEDropMode
End Property

'*1
Public Property Let OLEDropMode(ByVal Value As ListPlusOLEDropModeConstants)
    Select Case Value
        Case OLEDropModeNone, OLEDropModeManual
            UserControl.OLEDropMode = Value
        Case Else
            Err.Raise 380
    End Select
    UserControl.PropertyChanged "OLEDropMode"
End Property


Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Property Let MousePointer(ByVal newValue As MousePointerConstants)
    UserControl.MousePointer = newValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(ByVal newValue As IPictureDisp)
    UserControl.MouseIcon = newValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property
'*-

Public Property Get ItemIsChild(ByVal Index As Long) As Boolean
    ItemIsChild = m_Items(Index).IsChild
End Property

Public Property Get ItemIsGroup(ByVal Index As Long) As Boolean
    ItemIsGroup = m_Items(Index).IsGroup
End Property

Public Function GetItemParentGroup(ByVal Index As Long) As Long
    Dim i As Long
    If m_Items(Index).IsGroup Then
        GetItemParentGroup = Index: Exit Function
    End If
    
    If m_Items(Index).IsChild Then
        For i = Index To 0 Step -1
            If m_Items(i).IsGroup Then
                GetItemParentGroup = i: Exit Function
            End If
        Next
    Else
        GetItemParentGroup = Index: Exit Function
    End If
    GetItemParentGroup = -1
End Function

Public Function GetGroupChildCounts(ByVal Index As Long) As Long
    Dim i As Long, lCount As Long
    If m_Items(Index).IsGroup And Index < mListCount - 1 Then
        For i = Index + 1 To mListCount - 1
            If m_Items(i).IsChild Then
                lCount = lCount + 1
            Else
                Exit For
            End If
        Next
        GetGroupChildCounts = lCount
    Else
        GetGroupChildCounts = -1
    End If
    
End Function

Public Property Let ItemText(ByVal Index As Long, ByVal New_Value As String)
    m_Items(Index).Text = New_Value
    Draw
End Property

Public Property Get ItemText(ByVal Index As Long) As String
    ItemText = m_Items(Index).Text
End Property

Public Property Let ItemTextColor(ByVal Index As Long, ByVal New_Value As Long)
    m_Items(Index).TextColor = New_Value
    Draw
End Property

Public Property Get ItemTextColor(ByVal Index As Long) As Long
    ItemTextColor = m_Items(Index).TextColor
End Property

Public Property Let ItemBackColor(ByVal Index As Long, ByVal New_Value As Long)
    m_Items(Index).BackColor = New_Value
    Draw
End Property

Public Property Get ItemBackColor(ByVal Index As Long) As Long
    ItemBackColor = m_Items(Index).BackColor
End Property

Public Property Let ItemSubText(ByVal Index As Long, ByVal New_Value As String)
    m_Items(Index).SubText = New_Value
    Draw
End Property

Public Property Get ItemSubText(ByVal Index As Long) As String
    ItemSubText = m_Items(Index).SubText
End Property

Public Property Let ItemSubTextColor(ByVal Index As Long, ByVal New_Value As Long)
    m_Items(Index).SubTextColor = New_Value
    Draw
End Property

Public Property Get ItemSubTextColor(ByVal Index As Long) As Long
    ItemSubTextColor = m_Items(Index).SubTextColor
End Property

Public Property Let ItemBold(ByVal Index As Long, ByVal New_Value As Boolean)
    m_Items(Index).ItemBold = New_Value
    Draw
End Property

Public Property Get ItemBold(ByVal Index As Long) As Boolean
    ItemBold = m_Items(Index).ItemBold
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal New_Value As Variant)
    m_Items(Index).Data = New_Value
End Property

Public Property Get ItemData(ByVal Index As Long) As Variant
    ItemData = m_Items(Index).Data
End Property

Public Property Let ItemIcon(ByVal Index As Long, ByVal New_Value As Long)
    m_Items(Index).Icon = New_Value
    Draw
End Property

Public Property Get ItemIcon(ByVal Index As Long) As Long
    ItemIcon = m_Items(Index).Icon
End Property

Public Property Let ItemDisabled(ByVal Index As Long, ByVal New_Value As Boolean)
    m_Items(Index).Disabled = New_Value
    Draw
End Property

Public Property Get ItemDisabled(ByVal Index As Long) As Boolean
    ItemDisabled = m_Items(Index).Disabled
End Property

Public Property Let ItemSelected(ByVal Index As Long, ByVal New_Value As Boolean)
    m_Items(Index).Checked = New_Value
    Draw
End Property

Public Property Get ItemSelected(ByVal Index As Long) As Boolean
    ItemSelected = m_Items(Index).Checked
    Draw
End Property

Public Property Let ListIndex(New_Value As Long)
    m_Index = New_Value
    Draw
End Property

Public Property Get ListIndex() As Long
    ListIndex = m_Index
End Property

Public Property Get ListCount() As Long
    ListCount = mListCount
End Property

Public Property Get Shadow() As ClsShadow
    Set Shadow = cShadow
End Property

Public Property Get IsListVisible() As Boolean
    IsListVisible = bShow
End Property

Public Sub Refresh()
    Draw
End Sub

Public Sub Clear()
    GroupIsOpen = False
    m_Index = -1
    m_HotItem = -1
    mListCount = 0
    ReDim m_Items(0)
End Sub

Public Sub BSort(SortOrder As eSortOrder)
    Dim i As Long
    Dim bResult As Boolean
    Dim Item As tItem
    Dim Sorted As Boolean
    
    Do While Not Sorted
        Sorted = True
        For i = 0 To UBound(m_Items) - 1
            If SortOrder = SortAscendant Then
                bResult = StrComp(m_Items(i).Text, m_Items(i + 1).Text, vbTextCompare) = -1
            Else
                bResult = StrComp(m_Items(i).Text, m_Items(i + 1).Text, vbTextCompare) = -1
            End If
            
            If bResult Then
                Item = m_Items(i + 1)
                m_Items(i + 1) = m_Items(i)
                m_Items(i) = Item
                Sorted = False
            End If
        Next
    Loop
    Draw
End Sub


Public Sub QSort(SortOrder As eSortOrder)
    If SortOrder = SortAscendant Then
        QuickSortAsc LBound(m_Items), UBound(m_Items)
    Else
        QuickSortDes LBound(m_Items), UBound(m_Items)
    End If
    Draw
End Sub

Private Sub QuickSortAsc(ByVal First As Long, ByVal Last As Long)
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=231925[/url]
    Dim Low As Long, High As Long
    Dim MidValue As String
    Dim Item As tItem
    
    Low = First
    High = Last
    MidValue = m_Items((First + Last) \ 2).Text
    
    Do
        While StrComp(m_Items(Low).Text, MidValue, vbTextCompare) = -1
            Low = Low + 1
        Wend
        
        While StrComp(MidValue, m_Items(High).Text, vbTextCompare) = -1
            High = High - 1
        Wend
        
        If Low <= High Then
            Item = m_Items(Low)
            m_Items(Low) = m_Items(High)
            m_Items(High) = Item
           
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QuickSortAsc First, High
    If Low < Last Then QuickSortAsc Low, Last
End Sub
 
Private Sub QuickSortDes(ByVal First As Long, ByVal Last As Long)
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=231925[/url]
    Dim Low As Long, High As Long
    Dim MidValue As String
    Dim Item As tItem
    
    Low = First
    High = Last
    MidValue = m_Items((First + Last) \ 2).Text
    
    Do
        While StrComp(MidValue, m_Items(Low).Text, vbTextCompare) = -1
            Low = Low + 1
        Wend
        
        While StrComp(m_Items(High).Text, MidValue, vbTextCompare) = -1
            High = High - 1
        Wend
        
        If Low <= High Then
            Item = m_Items(Low)
            m_Items(Low) = m_Items(High)
            m_Items(High) = Item
           
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QuickSortDes First, High
    If Low < Last Then QuickSortDes Low, Last
End Sub





Public Function ImageListInitialize(Optional ByVal ImageWidth As Integer = 16, Optional ByVal ImageHeight As Integer = 16) As Boolean
    ImgLstWidth = ImageWidth * DpiFactor
    ImgLstHeight = ImageHeight * DpiFactor
    ImageListClear
End Function

Public Property Get ImgListWidth() As Long
   ImgListWidth = ImgLstWidth
End Property

Public Property Get ImgListHeight() As Long
   ImgListHeight = ImgLstHeight
End Property

Public Function ImageListAddImage(SrcImg As Variant, Optional Key As Variant) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, ImgLstWidth, ImgLstHeight)
    If hImage <> 0& Then
        cImageList.Add hImage, Key
        ImageListAddImage = cImageList.Count
    End If
End Function

Public Function ImageListRemoveImage(Index As Variant) As Long
    GdipDisposeImage cImageList(Index)
    cImageList.Remove Index
    ImageListRemoveImage = cImageList.Count
End Function

Public Function ImageListInsertImage(SrcImg As Variant, Optional Key As Variant, Optional Before, Optional After) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, ImgLstWidth, ImgLstHeight)
    If hImage <> 0& Then
        cImageList.Add hImage, Key, Before, After
        ImageListInsertImage = cImageList.Count
    End If
End Function

Public Sub ImageListClear()
    Dim i As Long
    For i = 1 To cImageList.Count
        GdipDisposeImage cImageList(i)
    Next
    Set cImageList = New Collection
End Sub

Public Property Get ImageListCount() As Long
    ImageListCount = cImageList.Count
End Property

Private Sub RoundRectPlus(ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, _
                       ByVal Width As Long, ByVal Height As Long, ByVal BackColor As Long, ByVal BorderColor As Long, Radius As Long)
    Dim hPen As Long, hBrush As Long
    Dim mpath As Long, hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    If BackColor <> 0 Then GdipCreateSolidFill BackColor, hBrush
    If BorderColor <> 0 Then GdipCreatePen1 BorderColor, 1 * DpiFactor, &H2, hPen
    
    If Radius = 0 Then
      
        If hBrush Then GdipFillRectangleI hGraphics, hBrush, Left, Top, Width, Height
        If hPen Then GdipDrawRectangleI hGraphics, hPen, Left, Top, Width, Height
        
    Else
        If GdipCreatePath(&H0, mpath) = 0 Then
            
            GdipAddPathArcI mpath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI mpath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI mpath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI mpath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure mpath
            
            If hBrush Then GdipFillPath hGraphics, hBrush, mpath
            If hPen Then GdipDrawPath hGraphics, hPen, mpath
            Call GdipDeletePath(mpath)
        End If
    End If
        
    If hBrush Then Call GdipDeleteBrush(hBrush)
    If hPen Then Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics
End Sub

Private Function LoadPictureEx(SrcImg As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
    Dim hImage1 As Long, hImage2 As Long, hGraphics As Long
    Dim DataArr() As Byte
    Dim lPictureRealWidth As Long, lPictureRealHeight As Long
    Dim X As Long, Y As Long, cx As Long, cy As Long
    Dim sngRatio1 As Single, sngRatio2 As Single
    Dim mpath As Long, Radius As Long, hPen As Long
    
    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    Call LoadImageFromArray(DataArr, hImage1)
                Else
                    Dim oXMLHTTP As Object
                    Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
                    
                    oXMLHTTP.Open "GET", SrcImg, True
                    oXMLHTTP.send
                    While oXMLHTTP.readyState <> 4
                        DoEvents
                    Wend
                    If oXMLHTTP.Status = 200 Then
                        DataArr() = oXMLHTTP.responseBody
                        Call LoadImageFromArray(DataArr, hImage1)
                    End If
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), hImage1)
            End If
        Case vbLong
            Const OBJ_BITMAP As Long = 7
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, hImage1)
            Else
                Call GdipCreateBitmapFromHICON(SrcImg, hImage1)
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, hImage1)
            
        Case (vbArray Or vbByte)
            DataArr() = SrcImg
            Call LoadImageFromArray(DataArr, hImage1)
    End Select

    If hImage1 <> 0 Then
        GdipGetImageWidth hImage1, lPictureRealWidth
        GdipGetImageHeight hImage1, lPictureRealHeight
        If Width = 0 Then Width = lPictureRealWidth
        If Height = 0 Then Height = lPictureRealHeight
        
        sngRatio1 = Width / lPictureRealWidth
        sngRatio2 = Height / lPictureRealHeight
        If sngRatio1 > sngRatio2 Then sngRatio1 = sngRatio2
        cx = lPictureRealWidth * sngRatio1: cy = lPictureRealHeight * sngRatio1
        X = (Width - cx) \ 2: Y = (Height - cy) \ 2

        GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        
        If m_ImagesRadius > 0 Then
            Radius = m_ImagesRadius * DpiFactor
            If GdipCreatePath(&H0, mpath) = 0 Then
                X = X + DpiFactor: Y = Y + DpiFactor
                cx = cx - DpiFactor * 2: cy = cy - DpiFactor * 2
                GdipAddPathArcI mpath, X, Y, Radius, Radius, 180, 90
                GdipAddPathArcI mpath, X + cx - Radius, Y, Radius, Radius, 270, 90
                GdipAddPathArcI mpath, X + cx - Radius, Y + cy - Radius, Radius, Radius, 0, 90
                GdipAddPathArcI mpath, X, Y + cy - Radius, Radius, Radius, 90, 90
                GdipClosePathFigure mpath
                GdipSetClipPath hGraphics, mpath, CombineModeIntersect
            End If
        End If
        
        GdipDrawImageRectRectI hGraphics, hImage1, X, Y, cx, cy, 0, 0, lPictureRealWidth, lPictureRealHeight, UnitPixel, 0&, 0&, 0&
        
        If mpath Then
            GdipCreatePen1 RGBtoARGB(vbButtonFace, 100), 1 * DpiFactor, &H2, hPen
            GdipResetClip hGraphics
            GdipDrawPath hGraphics, hPen, mpath
            GdipDeletePen hPen
            GdipDeletePath mpath
        End If
        
        GdipDeleteGraphics hGraphics
        GdipDisposeImage hImage1
        LoadPictureEx = hImage2
    End If
End Function

Public Function ImageListDraw(ByVal hdc As Long, ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional PictureColor As Long = CLR_NONE, Optional Disabled As Boolean) As Boolean
    Dim tMatrixColor As COLORMATRIX, tMatrixGray As COLORMATRIX
    Dim hAttributes As Long, hGraphics As Long
    Dim hImage As Long
    
    hImage = cImageList(Index)
    
    Call GdipCreateImageAttributes(hAttributes)
    
    With tMatrixColor
        If PictureColor <> CLR_NONE Then
            If (PictureColor And &H80000000) Then PictureColor = GetSysColor(PictureColor And &HFF&)
            Dim R As Byte, G As Byte, B As Byte

            B = ((PictureColor \ &H10000) And &HFF)
            G = ((PictureColor \ &H100) And &HFF)
            R = (PictureColor And &HFF)

            .M(0, 0) = R / 255
            .M(1, 0) = G / 255
            .M(2, 0) = B / 255
            .M(0, 4) = R / 255
            .M(1, 4) = G / 255
            .M(2, 4) = B / 255
        Else
            .M(0, 0) = 1
            .M(1, 1) = 1
            .M(2, 2) = 1
        End If
        .M(3, 3) = 1
        .M(4, 4) = 1

        If Disabled Then
            .M(0, 0) = 0.299
            .M(1, 0) = 0.299
            .M(2, 0) = 0.299
            .M(0, 1) = 0.587
            .M(1, 1) = 0.587
            .M(2, 1) = 0.587
            .M(0, 2) = 0.114
            .M(1, 2) = 0.114
            .M(2, 2) = 0.114
        End If
    End With
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
    ImageListDraw = GdipDrawImageRectRectI(hGraphics, hImage, Left, Top, Width, Height, 0, 0, Width, Height, UnitPixel, hAttributes) = 0&
    Call GdipDeleteGraphics(hGraphics)
    Call GdipDisposeImageAttributes(hAttributes)
End Function

Private Function LoadImageFromArray(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromArray_Error
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromArray = True
        End If
    End If
    Set IStream = Nothing
    
LoadImageFromArray_Error:
End Function

Private Function Base64Decode(ByVal sIn As String, ByRef bvOut() As Byte) As Boolean 'By Cocus
    Dim lLenOut As Long
    Const CRYPT_STRING_BASE64 = &H1
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, 0, VarPtr(lLenOut), 0, 0)
    If lLenOut = 0 Then Exit Function
    ReDim bvOut(lLenOut - 1)
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, VarPtr(bvOut(0)), VarPtr(lLenOut), 0, 0)
    Base64Decode = True
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long 'By LaVople
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function
'*-

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

Private Sub TextBox_SelStart(hWnd As Long, ByVal New_Value As Long)
    Call SendMessage(hWnd, EM_SETSEL, New_Value, ByVal New_Value)
End Sub

Private Function TextBox_SetText(hWnd As Long, sText As String) As Long
    TextBox_SetText = SendMessageW(hWnd, WM_SETTEXT, 0&, ByVal StrPtr(sText))
End Function

Private Function TextBox_GetText(hWnd As Long) As String
    Dim sText As String
    Dim lLength As Long

    lLength = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
    sText = Space$(lLength + 1)

    Call SendMessageW(hWnd, WM_GETTEXT, lLength + 1, ByVal StrPtr(sText))
    TextBox_GetText = Left$(sText, lLength)
End Function

' Ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
   ByRef bHandled As Boolean, _
   ByRef lReturn As Long, _
   ByVal hWnd As Long, _
   ByVal uMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long, _
   ByRef lParamUser As Long)

    On Error Resume Next

    If hWnd = UserControl.hWnd Then
        Select Case uMsg
            Case WM_HOTKEY
                Text1_KeyDown CInt(wParam), 0
                bHandled = True
    
            Case WM_MOUSELEAVE
                m_HotItem = -1
                bHideSubMenu = False
                bShowSubMenu = False
                Draw
                
            Case WM_IME_SETCONTEXT
                If Me.ListStyle = StyleDropDown Then HideList
        End Select
    Else
        'WM_CHAR
        If wParam <> vbKeyReturn Then
            SrcTextBox_Change
        End If
    End If

End Sub


