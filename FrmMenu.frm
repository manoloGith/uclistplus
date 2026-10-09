VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Menu2"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu1"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin Proyecto1.ucList ucList4 
      Height          =   480
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmMenu.frx":0000
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   2
   End
   Begin Proyecto1.ucList ucList3 
      Height          =   480
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmMenu.frx":0028
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   2
   End
   Begin Proyecto1.ucList ucList2 
      Height          =   480
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmMenu.frx":0050
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   2
   End
   Begin Proyecto1.ucList ucList1 
      Height          =   480
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmMenu.frx":0078
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   2
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMenu"
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

Private Sub Command1_Click()
    Dim Rect As Rect
    GetWindowRect Command1.hWnd, Rect
    ucList1.ListIndex = -1
    ucList1.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Rect As Rect
    If ucList1.IsListVisible Then Exit Sub
    GetWindowRect Command2.hWnd, Rect
    ucList2.ListIndex = -1
    ucList2.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucList2.HideList
    If ucList1.IsListVisible = False Then
    
    Dim Rect As Rect
    GetWindowRect Command3.hWnd, Rect
    ucList1.ListIndex = -1
    ucList1.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
    End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucList1.HideList
    If ucList2.IsListVisible = False Then
    
    Dim Rect As Rect
    GetWindowRect Command4.hWnd, Rect
    ucList2.ListIndex = -1
    ucList2.ShowList Rect.Left, Rect.Bottom, 0, Rect.Top
    End If
End Sub


Private Sub Form_Load()

    Dim i  As Long
    With ucList1
    
        .ImageListInitialize 16, 16
        .UseMonocromeIcons = True
        .ListWidth = 200
        .ItemHeight = 32
        '.SubTitleVisible = True
        .Shadow.ShadowOffsetX = 2
        .Shadow.ShadowOffsetY = 2
  
        For i = 1 To 8
            .ImageListAddImage App.Path & "\Vectores PNG\" & i & ".png"
        Next
        i = 0
'        For i = 1 To 8
'            .AddItem "Item " & i, i - 1, , , i Mod 3 = 0, , i Mod 4 = 0
'        Next
        .AddItem "Ci&ntia" & vbTab & "Ctrl+C", 1, , "Mama"
        .AddItem "&Mateo" & vbTab & "Ctrl+M", 2, , "Hijo"
        .AddItem "Leandr&o" & vbTab & "Ctrl+L", 3, , "Papa"
        .AddSubMenu "SubMenu", ucList2, 4
        .AddItem "-"
        .AddItem "Poto&", , , "Gato"
        '.Update
        .ItemTextColor(2) = vbRed
        .ItemSubTextColor(2) = vbGreen
        .ItemBackColor(2) = vbYellow
    End With

    With ucList2
        .ListWidth = 100
        .ItemHeight = 24

        .Shadow.ShadowOffsetX = 2
        .Shadow.ShadowOffsetY = 2
        For i = 1 To 3
            .AddItem "Item " & i
        Next
        .AddSubMenu "SubMenu", ucList3
        '.Update
    End With
    
     With ucList3
        .ListWidth = 200
        .ItemHeight = 24

        .Shadow.ShadowOffsetX = 2
        .Shadow.ShadowOffsetY = 2
        For i = 1 To 3
            .AddItem "Item " & i
        Next
        .AddSubMenu "SubMenu", ucList4
        '.Update
    End With
    
    With ucList4
        .ListWidth = 100
        .ItemHeight = 24

        .Shadow.ShadowOffsetX = 2
        .Shadow.ShadowOffsetY = 2
        For i = 1 To 10
            .AddItem "Item " & i
        Next
        
        '.Update
    End With
End Sub



Private Sub ucList1_ItemClick()
        Me.Caption = ucList1.ItemText(ucList1.ListIndex)
End Sub

Private Sub ucList2_ItemClick()
    Me.Caption = ucList2.ItemText(ucList2.ListIndex)
End Sub

Private Sub ucList3_ItemClick()
    Me.Caption = ucList3.ItemText(ucList3.ListIndex)
End Sub

