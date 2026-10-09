VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmShadow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "&H8000000F&&H8000000F&"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   8
      Left            =   5040
      Max             =   5000
      Min             =   20
      SmallChange     =   100
      TabIndex        =   29
      Top             =   4320
      Value           =   100
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   7
      Left            =   5040
      Max             =   100
      Min             =   20
      TabIndex        =   27
      Top             =   3720
      Value           =   100
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5040
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   24
      Top             =   3120
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   6
      Left            =   5040
      Max             =   100
      TabIndex        =   23
      Top             =   2280
      Value           =   5
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   5
      Left            =   5040
      Max             =   100
      TabIndex        =   22
      Top             =   1560
      Value           =   100
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5040
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   21
      Top             =   960
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   4
      Left            =   5040
      Max             =   100
      TabIndex        =   17
      Top             =   360
      Value           =   1
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   3
      LargeChange     =   2
      Left            =   1920
      Max             =   100
      TabIndex        =   16
      Top             =   4320
      Value           =   100
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      LargeChange     =   2
      Left            =   1920
      Max             =   100
      TabIndex        =   14
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   12
      Top             =   3720
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      Left            =   1920
      Max             =   100
      TabIndex        =   10
      Top             =   1560
      Value           =   8
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   735
      Left            =   2040
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   2040
      Width           =   255
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   1800
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      Left            =   1920
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Value           =   5
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Width:"
      Height          =   255
      Index           =   13
      Left            =   3360
      TabIndex        =   30
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Opacity:"
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   28
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000F&
      Index           =   2
      X1              =   224
      X2              =   376
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000F&
      Index           =   0
      X1              =   16
      X2              =   168
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Color:"
      Height          =   255
      Index           =   11
      Left            =   3360
      TabIndex        =   26
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Radius"
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   25
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Opacity:"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   20
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Color:"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   19
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Width:"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   18
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor Opacity:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margin:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow Opacity:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow Color:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00424242&
      Height          =   435
      Left            =   6150
      TabIndex        =   6
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ShadowOffset:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow size:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "FrmShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_SYSCOMMAND As Long = &H112
Private Const MOUSE_MOVE As Long = &HF012&
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Dim C As ClsShadow

Private Sub axLabelPlus1_Click()

End Sub

Private Sub Form_Activate()
    C.ShowBorders Me.hwnd, False
End Sub

Private Sub Form_Load()
    Dim hRgn As Long
    
    Set C = New ClsShadow
    hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 10 * C.GetWindowsDPI, 10 * C.GetWindowsDPI)
    SetWindowRgn Me.hwnd, hRgn, True
    
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0&
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set C = Nothing
End Sub

Private Sub HScroll1_Change(Index As Integer)
    Dim Value As Long, hRgn As Long
    Value = HScroll1(Index).Value
    Select Case Index
        Case 0: C.ShadowSize = Value
        Case 1: C.ShadowOpacity = Value
        Case 2: C.Margin = Value
        Case 3: C.BackColorOpacity = Value
        Case 4: C.BorderWidth = Value
        Case 5: C.BorderOpacity = Value
        Case 6
            C.BorderRadius = Value
            hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, Value * C.GetWindowsDPI, Value * C.GetWindowsDPI)
            SetWindowRgn Me.hwnd, hRgn, True
        Case 7
            SetLayeredWindowAttributes Me.hwnd, 0, (Value / 100) * 255, LWA_ALPHA
        Case 8:
            Me.Width = 6500 + Value
            hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, Value * C.GetWindowsDPI, C.BorderRadius * C.GetWindowsDPI)
            SetWindowRgn Me.hwnd, hRgn, True
    End Select
    C.Refresh
End Sub

Private Sub HScroll3_Change()
    C.ShadowOffsetX = HScroll3.Value
    C.Refresh
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Picture1_Click(Index As Integer)
    On Error GoTo OnCancelExit
    CommonDialog1.ShowColor
    Picture1(Index).BackColor = CommonDialog1.Color
    Select Case Index
        Case 0: C.ShadowColor = CommonDialog1.Color
        Case 1: C.BackColor = CommonDialog1.Color
        Case 2: C.BorderColor = CommonDialog1.Color
        Case 3: Me.BackColor = CommonDialog1.Color
    End Select
    C.Refresh
OnCancelExit:
End Sub

Private Sub VScroll1_Change()
    C.ShadowOffsetY = VScroll1.Value
    C.Refresh
End Sub

Private Sub Command7_Click()
    HScroll3.Value = 0
    VScroll1.Value = 0
End Sub
