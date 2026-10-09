VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList1 
      Height          =   6000
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6090
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
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    With ucList1
        .ImageListInitialize 32, 32
        .ImageListAddImage App.Path & "\VB6.png"
        .ItemHeight = 48
        .SubTitleVisible = True
        .AddItem "FrmAutocomplete", 0, , "Autocomplete a Textbox with a suggestion List"
        .AddItem "FrmComboBox", 0, , "Litle example of usercontrol Style ComboBox "
        .AddItem "FrmCustomDraw", 0, , "List items are custom draw"
        .AddItem "FrmDragDrop", 0, , "An example of how to drag and drop an item into another list"
        .AddItem "FrmDropDownList", 0, , "Drop-down list, not combobox, to call from a button or custom control"
        .AddItem "FrmGoogleSearch", 0, , "A complete example with everything, autocomplete in real time with google apis."
        .AddItem "FrmListBox", 0, , "ucList with ListStyle = StyleListBox"
        .AddItem "FrmMenu", 0, , "ucList with ListStyle = StyleMenu"
        .AddItem "FrmShadow", 0, , "Example of using the clsshadow class in forms"
        
        .Update
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        ucList1.Move ucList1.Left, ucList1.Top, Me.ScaleWidth - ucList1.Left * 2, Me.ScaleHeight - ucList1.Top * 2
    End If
End Sub


Private Sub ucList1_ItemClick()
    Select Case ucList1.ListIndex
        Case 0: FrmAutocomplete.Show
        Case 1: FrmComboBox.Show
        Case 2: FrmCustomDraw.Show
        Case 3: FrmDragDrop.Show
        Case 4: FrmDropDownList.Show
        Case 5: FrmGoogleSearch.Show
        Case 6: FrmListBox.Show
        Case 7: FrmMenu.Show
        Case 8: FrmShadow.Show
    End Select
End Sub
