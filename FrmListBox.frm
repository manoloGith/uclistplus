VERSION 5.00
Begin VB.Form FrmListBox 
   Caption         =   "Form2"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12480
   LinkTopic       =   "Form2"
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   832
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList5 
      Height          =   3000
      Left            =   8880
      TabIndex        =   4
      Top             =   3600
      Width           =   2730
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmListBox.frx":0000
      backcolor       =   4342338
      findtextcolor   =   -2147483635
      selectioncolor  =   8650884
      linecolor       =   6776679
      liststyle       =   1
   End
   Begin Proyecto1.ucList ucList2 
      Height          =   3000
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   2805
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmListBox.frx":0028
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random BackColor"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin Proyecto1.ucList ucList3 
      Height          =   3000
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   3090
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmListBox.frx":0050
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   1
   End
   Begin Proyecto1.ucList ucList4 
      Height          =   3000
      Left            =   10080
      TabIndex        =   3
      Top             =   240
      Width           =   2250
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmListBox.frx":0078
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   1
   End
   Begin Proyecto1.ucList ucList1 
      Height          =   3000
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3330
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmListBox.frx":00A0
      listwidth       =   500
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
      linecolor       =   13487565
      liststyle       =   1
   End
End
Attribute VB_Name = "FrmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cShadow As ClsShadow
Dim cShadow2 As ClsShadow
'--------------EJEMPLO 1

Private Sub Example1()
    Dim MyLine As String
    Dim i As Long, FF As Integer
    
    With ucList1
        .ImageListInitialize 40, 40
        .ImagesRadius = 37
        .SubTitleVisible = True
        .ItemHeight = 48
        .ListDesign = GoogleDesign
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
        .Update
    End With
End Sub


Private Sub Command1_Click()
    Dim RndColor As Long
    
    RndColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
    Me.BackColor = RndColor
End Sub

Private Sub ucList1_IndexChange()
      Me.Caption = ucList1.ItemText(ucList1.ListIndex)
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

Private Sub ucList2_ItemClick()
    Me.Caption = ucList2.ItemText(ucList2.ListIndex)
   ' Label1.Caption = ucList2.ItemSubText(ucList2.ListIndex)

End Sub
'============== FIN EJEMPLO2

'-------------- EJEMPLO 3
Private Sub Example3()
    Dim MyLine As String
    Dim i As Long, FF As Integer
    
    With ucList3

        .UseMonocromeIcons = True
        .ItemHeight = 32
        '.ShowFilter = True
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
        .Update
    End With

End Sub

Private Sub ucList3_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList3.ListCount - 1
        If ucList3.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    'If lCount > 3 Then ucList3.ItemSelected(ucList3.ListIndex) = False
    Me.Caption = "selecionados " & lCount
    
    
    If ucList3.ItemIsChild(ucList3.ListIndex) Then
        Dim GroupIndex As Long
        GroupIndex = ucList3.GetItemParentGroup(ucList3.ListIndex)
        Debug.Print "Item Parent Group is: " & GroupIndex,
        Debug.Print "Parent Group have "; ucList3.GetGroupChildCounts(GroupIndex) & " Items"
    End If
    
End Sub


'============== FIN EJEMPLO3

'-------------- EJEMPLO 4
Private Sub Example4()
    Dim i  As Long
    With ucList4
        .BorderRadius = 0
        .ListDesign = MicrosoftDesign
        .ItemHeight = 32
        .SelectionColor = vbBlack
        .Shadow.BorderColor = vbBlack
        .MaxVisibleItems = 5
        
        For i = 1 To 10
            .AddItem "Item " & i
        Next
        .Update
    End With
End Sub


Private Sub ucList4_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList4.ListCount - 1
        If ucList4.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    Me.Caption = ucList4.ItemText(ucList4.ListIndex)
End Sub

'============== FIN EJEMPLO4


'-------------- EJEMPLO 5
Private Sub Example5()
    Dim i  As Long
    With ucList5
        .ImageListInitialize 16, 16
        .UseMonocromeIcons = True
        .ItemHeight = 32
     
        '.Shadow.BorderColor = vbGrayText
        For i = 1 To 8
            .ImageListAddImage App.Path & "\Vectores PNG\" & i & ".png"
        Next
        i = 0
  
        .ItemHeight = 32

        '.MaxVisibleItems = 5
        For i = 1 To 8
            If i Mod 3 = 0 Then
                .AddItem "-"
            Else
                .AddItem "Item " & i, i - 1, , , , i Mod 4 = 0
            End If
        Next
        .Update
    End With
End Sub

Private Sub ucList5_ItemClick()
    Dim i As Long, lCount As Long
    For i = 0 To ucList5.ListCount - 1
        If ucList5.ItemSelected(i) Then
            lCount = lCount + 1
        End If
    Next
    If lCount > 5 Then ucList5.ItemSelected(ucList5.ListIndex) = False
    Me.Caption = ucList5.ItemText(ucList5.ListIndex)
End Sub

'============== FIN EJEMPLO5


Private Sub Form_Load()
    Example1
    Example2
    Example3
    Example4
    Example5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cShadow = Nothing
    Set cShadow2 = Nothing
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    ucList1.Height = Me.ScaleHeight - ucList1.Top * 2
    ucList2.Height = ucList1.Height
End Sub



