VERSION 5.00
Begin VB.Form FrmDropDownList 
   Caption         =   "Form5"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9075
   LinkTopic       =   "Form5"
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList5 
      Height          =   480
      Left            =   5880
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin Proyecto1.ucList ucList4 
      Height          =   480
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin Proyecto1.ucList ucList3 
      Height          =   480
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin Proyecto1.ucList ucList2 
      Height          =   480
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin Proyecto1.ucList ucList1 
      Height          =   480
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "FrmDropDownList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'--------------EJEMPLO 1

Private Sub Example1()
    Dim MyLine As String
    Dim i As Long, FF As Integer
    
    With ucList1
        .ImageListInitialize 40, 40
        .ImagesRadius = 37
        .SubTitleVisible = True
        .ItemHeight = 48
        .MaxVisibleItems = 5
        .ListStyle = GoogleDesign
        .DrawSeparatorLine = True
        
        For i = 1 To 9
            .ImageListAddImage App.Path & "\Caras\" & i & ".jpg"
        Next
        i = 0
        FF = FreeFile
        Open App.Path & "\Personal.txt" For Input As #FF
            Do While Not EOF(1)
                Line Input #FF, MyLine
                .AddItem Split(MyLine, ", ")(0), i, , Split(MyLine, ", ")(1)
                i = i + 1
            Loop
        Close FF
    End With
End Sub

Private Sub Command1_Click()
    Dim Rect As Rect
    GetWindowRect Command1.hWnd, Rect
    ucList1.ShowList Rect.Left, Rect.Bottom
End Sub



Private Sub ucList1_ItemClick()
    Command1.Caption = ucList1.ItemText(ucList1.ListIndex)
End Sub
'============== FIN EJEMPLO1

'-------------- EJEMPLO 2
Private Sub Example2()
    Dim MyLine As String
    Dim i As Long, FF As Integer
    
    With ucList2
        .ImageListInitialize 16, 16
        .UseMonocromeIcons = True
        .ItemHeight = 32
        .ShowFilter = True
        .MaxVisibleItems = 8
        .ImageListAddImage App.Path & "\Vectores PNG\8.png"
        
        i = 0
        FF = FreeFile
        Open App.Path & "\Nombres.txt" For Input As #FF
            Do While Not EOF(FF)
                Line Input #FF, MyLine
                If InStr(MyLine, ": ") = 0 Then
                    .OpenGroup MyLine
                Else
                    .AddItem Split(MyLine, ": ")(0), 0, , Split(MyLine, ": ")(1)
                End If
            Loop
        Close FF


    End With
End Sub

Private Sub Command2_Click()
    Dim Rect As Rect
    GetWindowRect Command2.hWnd, Rect
    ucList2.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
End Sub

Private Sub ucList2_ItemClick()
    Command2.Caption = ucList2.ItemText(ucList2.ListIndex)
    Label1.Caption = ucList2.ItemSubText(ucList2.ListIndex)

End Sub
'============== FIN EJEMPLO2

'-------------- EJEMPLO 3
Private Sub Example3()
    Dim MyLine As String
    Dim i As Long, FF As Integer
    
    With ucList3

        .UseMonocromeIcons = True
        .ItemHeight = 32
        .ShowFilter = True
        .MaxVisibleItems = 7
        .MultiSelect = True

        i = 0
        FF = FreeFile
        Open App.Path & "\Groups.txt" For Input As #FF
            Do While Not EOF(FF)
                Line Input #FF, MyLine
                i = i + 1
                If Left$(MyLine, 1) = "-" Then
                    .OpenGroup Mid$(MyLine, 2), 0
                Else
                    .AddItem MyLine
                End If
            Loop
        Close FF
    End With

End Sub

Private Sub Command3_Click()
    Dim Rect As Rect
    GetWindowRect Command3.hWnd, Rect
    ucList3.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
    
End Sub

Private Sub ucList3_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList3.ListCount - 1
        If ucList3.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    'If lCount > 3 Then ucList3.ItemSelected(ucList3.ListIndex) = False
    Command3.Caption = "selecionados " & lCount
End Sub


'============== FIN EJEMPLO3

'-------------- EJEMPLO 4
Private Sub Example4()
    Dim i  As Long
    With ucList4
        .BorderRadius = 0
        .ListDesign = MicrosoftDesign
        .ItemHeight = Command4.Height / ucList4.GetWindowDpi
        .ListWidth = Command4.Width / ucList4.GetWindowDpi
        .SelectionColor = vbBlack
        .Shadow.BorderColor = vbBlack
        .MaxVisibleItems = 5
        
        For i = 1 To 10
            .AddItem "Item " & i
        Next

    End With
End Sub

'Mouse Down-----------\/
Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Rect As Rect
    If ucList4.IsListVisible Then Exit Sub
    GetWindowRect Command4.hWnd, Rect
    If ucList4.ListIndex > -1 Then
        Rect.Bottom = Rect.Bottom - ucList4.GetItemTop(ucList4.ListIndex) - (ucList4.ItemHeight * ucList4.GetWindowDpi)
    End If
    ucList4.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
End Sub


Private Sub ucList4_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList4.ListCount - 1
        If ucList4.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    Command4.Caption = ucList4.ItemText(ucList4.ListIndex)
End Sub

'============== FIN EJEMPLO4


'-------------- EJEMPLO 5
Private Sub Example5()
    Dim i  As Long
    With ucList5
        .ImageListInitialize 16, 16
        .UseMonocromeIcons = True
        .ItemHeight = 32
        .SelectionColor = &H8CAD42
        .BackColor = &H424242
        
        .Shadow.BorderColor = vbGrayText
        For i = 1 To 8
            .ImageListAddImage App.Path & "\Vectores PNG\" & i & ".png"
        Next
        i = 0
  
        .ItemHeight = Command5.Height / ucList5.GetWindowDpi
        .ListWidth = Command5.Width / ucList5.GetWindowDpi
        .MaxVisibleItems = 5
        For i = 1 To 8
            If i Mod 3 = 0 Then
                .AddItem "-"
            Else
                .AddItem "Item " & i, i - 1, , , , i Mod 4 = 0
            End If
        Next
        
    End With
End Sub

'Mouse Down-----------\/
Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Rect As Rect
    If ucList5.IsListVisible Then Exit Sub
    GetWindowRect Command5.hWnd, Rect
 
    ucList5.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
End Sub


Private Sub ucList5_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList5.ListCount - 1
        If ucList5.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    If lCount > 5 Then ucList5.ItemSelected(ucList5.ListIndex) = False
    Command5.Caption = ucList5.ItemText(ucList5.ListIndex)
End Sub

'============== FIN EJEMPLO5


Private Sub Form_Load()
    Example1
    Example2
    Example3
    Example4
    Example5
End Sub



