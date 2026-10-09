VERSION 5.00
Begin VB.Form FrmCustomDraw 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Custom Draw"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6720
   LinkTopic       =   "Form3"
   ScaleHeight     =   6345
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList1 
      Height          =   5880
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5970
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
      LineColor       =   13487565
      ListStyle       =   1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "FrmCustomDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Dim DpiFactor As Single

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
'        .OpenGroup vbNullString, , -3
'        .CloseGroup
'
'        .AddItem "asdfdsd", , 2
'        .ItemBackColor(1) = &H6BBD94
        i = 0
        FF = FreeFile
        Open App.Path & "\Personal.txt" For Input As #FF
            Do While Not EOF(1)
                Line Input #FF, MyLine
                .AddItem Split(MyLine, ", ")(0), i, 2, Split(MyLine, ", ")(1)
                i = i + 1
            Loop
        Close FF
        .Update
       
    End With

End Sub

Private Sub Form_Load()
    DpiFactor = ucList1.GetWindowDpi
    
    Example1
End Sub

Private Sub ucList1_ItemClick()
    Dim i As Long
    Dim lStart As Long
    If ucList1.ItemData(ucList1.ListIndex) = 1 Then
        
            lStart = ucList1.ListIndex
            For i = lStart + 1 To ucList1.ListCount - 1
                If ucList1.ItemData(lStart + 1) = 0 Then
                    ucList1.RemoveItem lStart + 1
                Else
                    Exit For
                End If
            Next
            
            ucList1.ItemData(ucList1.ListIndex) = 2
       
    ElseIf ucList1.ItemData(ucList1.ListIndex) = 2 Then
        ucList1.ItemData(ucList1.ListIndex) = 1
        
        For i = 1 To 4 * Rnd + 1
            ucList1.AddItem Space(2) & "HOLA MUNDO", , 0, , , , ucList1.ListIndex + i
        Next
    End If
    ucList1.Update
End Sub

Private Sub ucList1_ItemCustomDraw(Index As Long, ByVal hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long)
    Dim Rect As Rect
    Dim hBrush As Long
 

            
    If ucList1.ItemData(Index) > 0 Then
        DrawTheme hdc, Width - 24 * DpiFactor, Top + Height / 2 - 16 * DpiFactor / 2, IIf(ucList1.ItemData(Index) = 1, True, False)
    Else
        If Index < 4 Then
            hBrush = CreateSolidBrush(vbRed)
        Else
            hBrush = CreateSolidBrush(vbGreen)
        End If
        With Rect
            .Top = Top + 2 * DpiFactor
            .Right = 4 * DpiFactor
            .Bottom = Top + Height - 1 * DpiFactor
        End With
        FillRect hdc, Rect, hBrush
        DeleteObject hBrush
    
    End If
End Sub


Private Sub DrawTheme(hdc As Long, X As Long, Y As Long, Colapsed As Boolean)
    Dim hTheme As Long
    Dim Rect As Rect
    
    hTheme = OpenThemeData(Me.hWnd, StrPtr("ItemsView::ListView"))
    If hTheme = 0 Then Exit Sub
    
    With Rect
        .Left = X
        .Top = Y
        .Right = X + 16 * DpiFactor
        .Bottom = Y + 16 * DpiFactor
    End With
    
    Call DrawThemeBackground(hTheme, hdc, IIf(Colapsed, 9, 8), 0&, Rect, ByVal 0&)
End Sub

