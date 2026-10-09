VERSION 5.00
Begin VB.UserControl ucSelect 
   BackColor       =   &H80000005&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   Begin Proyecto1.ucList ucList1 
      Height          =   480
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      FindTextColor   =   -2147483635
      SelectionColor  =   -2147483635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "ucSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Event ItemClick()

Private Const CLR_NONE          As Long = &HFFFFFFFF
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_VCENTER        As Long = &H4
Private Const DT_WORD_ELLIPSIS  As Long = &H40000
Private Const DT_CENTER         As Long = &H1
Private Const DT_RIGHT          As Long = &H2
Private Const DT_CALCRECT       As Long = &H400
Private Const WM_MOUSEWHEEL     As Long = &H20A
Dim sChrArrowDown As String
Dim sChrArrowUp As String
Dim cShadow As ClsShadow
Dim m_Text As String
Dim DpiFactor As Single
Dim lMargin As Long
Dim cSubClass As clsSubClass

Public Property Get List() As ucList
    Set List = ucList1
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    UserControl.BackColor = New_Value
    DrawArrow
    If Ambient.UserMode Then cShadow.BackColor = New_Value
    PropertyChanged "BackColor"
End Property


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

End Property


Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Value As String)
    m_Text = New_Value
    UserControl.Refresh
    PropertyChanged "Text"
End Property

Private Sub ucList1_IndexChange()
    ucList1_ItemClick
End Sub

Private Sub ucList1_ItemClick()
    m_Text = ucList1.ItemText(ucList1.ListIndex)
    UserControl.Refresh
    RaiseEvent ItemClick
End Sub


Private Sub UserControl_EnterFocus()
    With cShadow
        
        .ShadowOpacity = 30
        .BorderColor = &HFF9033
        .Refresh
    End With

'    If cSubClass.ssc_Subclass(UserControl.hWnd, , , Me) Then
'        cSubClass.ssc_AddMsg UserControl.hWnd, WM_MOUSEWHEEL, MSG_AFTER
'    End If
End Sub

Private Sub UserControl_ExitFocus()
'    cSubClass.ssc_UnSubclass UserControl.hWnd

    With cShadow
        .ShadowOpacity = 0
        .BorderColor = vbActiveBorder
        .Refresh
    End With
End Sub

Private Sub UserControl_Initialize()
       
    Set cShadow = New ClsShadow
    DpiFactor = cShadow.GetWindowsDPI
    lMargin = 8 * DpiFactor
    
    Picture1.Font.Name = "Segoe MDL2 Assets"
    
    If Picture1.Font.Name = "Segoe MDL2 Assets" Then
        sChrArrowDown = ChrW2(&HE011)
        sChrArrowUp = ChrW2(&HE010)
    Else
        Picture1.Font.Name = "Wingdings 3"
        sChrArrowDown = Chr(113)
        sChrArrowUp = Chr(112)
    End If
    Picture1.ForeColor = vbBlack
    Picture1.FillColor = Picture1.ForeColor
    Picture1.Font.Size = 11
    Picture1.AutoRedraw = True
    
    DrawArrow
    

End Sub
'*1

Private Sub DrawArrow()
    Dim Rect As Rect
    
    Picture1.BackColor = UserControl.BackColor
    Rect.Right = Picture1.ScaleWidth
    Rect.Bottom = Picture1.ScaleHeight
    DrawText Picture1.hdc, StrPtr(sChrArrowDown), 1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    
   ' Picture1.Line (0, 0)-(0, Picture1.ScaleHeight)
    
    Picture1.Refresh
End Sub

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then UserControl_MouseDown vbLeftButton, 0, 0, 0
    ucList1.SetKeyDown KeyCode
   
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    ucList1.SetKeyPress KeyAscii
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Rect As Rect
    Dim M As Long, B As Long

    If ucList1.IsListVisible Then
        ucList1.HideList
        Exit Sub
    End If
    GetWindowRect UserControl.hWnd, Rect
    M = cShadow.Margin * DpiFactor
    B = 2 * DpiFactor
    ucList1.ShowList Rect.Left - M, Rect.Bottom + M + B, 0, Rect.Top - M - B

End Sub

Private Sub UserControl_Paint()
    Dim Rect As Rect

    If Len(m_Text) Then
        With Rect
            .Right = UserControl.ScaleWidth - Picture1.Width
            .Bottom = UserControl.ScaleHeight
            .Left = lMargin * 2
        End With
        If ucList1.ImageListCount > 0 Then
        
            ucList1.ImageListDraw UserControl.hdc, ucList1.ItemIcon(ucList1.ListIndex) + 1, _
                Rect.Left, UserControl.ScaleHeight / 2 - ucList1.ImgListHeight / 2, _
                ucList1.ImgListWidth, ucList1.ImgListHeight
            
            Rect.Left = Rect.Left + ucList1.ImgListWidth + lMargin
        End If

        DrawText UserControl.hdc, StrPtr(m_Text), -1, Rect, DT_SINGLELINE Or DT_VCENTER
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Me.BackColor = .ReadProperty("BackColor", UserControl.BackColor)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    If Ambient.UserMode Then
        Set cSubClass = New clsSubClass
        If cSubClass.ssc_Subclass(UserControl.hWnd, , , Me) Then
            cSubClass.ssc_AddMsg UserControl.hWnd, WM_MOUSEWHEEL, MSG_AFTER

        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor)
        Call .WriteProperty("Font", UserControl.Font)
    End With
End Sub

Private Sub UserControl_Resize()
    Dim Rect As Rect
    Picture1.Move UserControl.ScaleWidth - UserControl.ScaleHeight, 0, UserControl.ScaleHeight, UserControl.ScaleHeight
    Picture1.Cls
    Rect.Right = Picture1.ScaleWidth
    Rect.Bottom = Picture1.ScaleHeight
    DrawText Picture1.hdc, StrPtr(sChrArrowDown), 1, Rect, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    Picture1.Refresh

End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        
        With cShadow
            .ShadowColor = &HFF9033
            .ShadowSize = 5
            .Margin = 3
            .ShadowOpacity = 0
            .ShowBorders UserControl.hWnd, True
            
        End With

    End If
End Sub

Private Sub UserControl_Terminate()
    Set cShadow = Nothing
    If Not cSubClass Is Nothing Then
        'cSubClass.ssc_UnSubclass UserControl.hWnd
        Set cSubClass = Nothing
    End If
End Sub


Private Sub WndProc(ByVal bBefore As Boolean, _
    ByRef bHandled As Boolean, _
    ByRef lReturn As Long, _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef lParamUser As Long)

    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            ucList1.SetKeyDown vbKeyDown
        Else
            ucList1.SetKeyDown vbKeyUp
        End If
    End If
End Sub


