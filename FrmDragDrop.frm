VERSION 5.00
Begin VB.Form FrmDragDrop 
   Caption         =   "Drag & Drop"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5655
   LinkTopic       =   "Form8"
   ScaleHeight     =   3720
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList1 
      Height          =   3000
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2250
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
   End
   Begin Proyecto1.ucList ucList2 
      Height          =   3000
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   2250
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
   End
End
Attribute VB_Name = "FrmDragDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim DragPos1 As Long
Dim DragPos2 As Long
Dim CtlID As Long

Private Sub Form_Load()
    ucList1.AddItem "James"
    ucList1.AddItem "Frederick"
    ucList1.AddItem "Ann"
    ucList1.AddItem "Paul"
    ucList1.AddItem "Sarah"
    ucList1.OLEDropMode = 1
    ucList2.OLEDropMode = 1
    DragPos1 = -2
End Sub

Private Sub ucList1_ItemCustomDraw(Index As Long, ByVal hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long)
    If Not CtlID = 1 Then
        If DragPos1 = -1 And Index = ucList1.ListCount - 1 Then
            DrawLine hdc, 0, Top + Height, Width - 2, Top + Height, vbRed, 1
        ElseIf Index = DragPos1 Then
            DrawLine hdc, 0, Top, Width - 2, Top, vbRed, 1
        End If
    End If
End Sub

Private Sub ucList1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ucList1.ListIndex = -1 Then Exit Sub
    ucList1.OLEDrag
End Sub

Private Sub ucList1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        DragPos1 = -2
    Else
        
        DragPos1 = ucList1.ItemHitest(Y / Screen.TwipsPerPixelY)
    End If
    ucList1.Refresh
End Sub


Private Sub ucList1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    CtlID = 1
    DragPos2 = -2
    Data.SetData ucList1.Text
End Sub

Private Sub ucList2_ItemCustomDraw(Index As Long, ByVal hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long)
    If Not CtlID = 2 Then
        If DragPos2 = -1 And Index = ucList2.ListCount - 1 Then
            DrawLine hdc, 0, Top + Height, Width - 2, Top + Height, vbRed, 1
        ElseIf Index = DragPos2 Then
            DrawLine hdc, 0, Top, Width - 2, Top, vbRed, 1
        End If
    End If
End Sub

Private Sub ucList2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strList As String
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strList = Data.GetData(vbCFText)
    If Not CtlID = 2 Then
        
        ucList2.AddItem strList, , , , , , DragPos2
        DragPos2 = -2
        ucList2.Refresh
        ucList1.RemoveItem ucList1.ListIndex
        ucList1.Refresh
    End If
End Sub

Private Sub ucList2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        DragPos2 = -2
    Else
        DragPos2 = ucList2.ItemHitest(Y / Screen.TwipsPerPixelY)
    End If
    ucList2.Refresh
End Sub

Private Sub ucList2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ucList2.ListIndex = -1 Then Exit Sub
    ucList2.OLEDrag    ' Begin dragging
End Sub
Private Sub ucList2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    ' Only allow moves
    AllowedEffects = vbDropEffectMove
    ' Assign the ListBox selection to the DataObject
    CtlID = 2
    DragPos1 = -2
    Data.SetData ucList2.Text
End Sub
Private Sub ucList1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strList As String
    ' Check the format of the DataObject
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = Data.GetData(vbCFText)
    ' If the item was not dropped on itself
    If Not CtlID = 1 Then
        ucList1.AddItem strList, , , , , , DragPos1
        DragPos1 = -2
        ucList1.Refresh
        'Remove the item from the ListBox
        ucList2.RemoveItem ucList2.ListIndex
        ucList2.Refresh
    End If
End Sub

Public Sub DrawLine(ByVal hdc As Long, _
                    ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    Optional ByVal Color As Long = -1, _
                    Optional ByVal BorderWidth As Long = 1)
  
    Dim hPen As Long, hBrush As Long
    Dim TransColor As Long
    Dim OldPen As Long, OldBrush As Long
    Dim PT(2) As POINTAPI
  

    Call OleTranslateColor(Color, 0&, TransColor)
    hPen = CreatePen(0, BorderWidth, TransColor)
    hBrush = CreateSolidBrush(TransColor)
    
    OldPen = SelectObject(hdc, hPen)
    OldBrush = SelectObject(hdc, hBrush)
  
    If X1 >= 0 Then
        MoveToEx hdc, X1, Y1, 0
    End If
  
    LineTo hdc, X2, Y2
  
    PT(0).X = 0: PT(0).Y = Y1 - 3
    PT(1).X = 0: PT(1).Y = Y1 + 3
    PT(2).X = 3: PT(2).Y = Y1

    Polygon hdc, PT(0), 3
    
    PT(0).X = X2: PT(0).Y = Y1 - 3
    PT(1).X = X2: PT(1).Y = Y1 + 3
    PT(2).X = X2 - 3: PT(2).Y = Y1

    Polygon hdc, PT(0), 3
    
    DeleteObject SelectObject(hdc, OldPen)
    DeleteObject SelectObject(hdc, OldBrush)
    

End Sub

