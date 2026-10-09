VERSION 5.00
Begin VB.Form FrmComboBox 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucSelect SelectLocalidades 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ucSelect SelectDepartamentos 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ucSelect SelectProvincias 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      BackColor       =   12451773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://georef-ar-api.readthedocs.io/es/latest/
Private Function LoadJson(ByVal Url As String, ByRef oJsonBag As JsonBag) As Boolean
    Dim ServerXmlHttp As Object
    
    Set ServerXmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ServerXmlHttp.Open "GET", Url
    ServerXmlHttp.send

    If ServerXmlHttp.Status = 200 Then
        Set oJsonBag = New JsonBag
        oJsonBag.JSON = ServerXmlHttp.responseText
        LoadJson = True
    End If
    
    Set ServerXmlHttp = Nothing
End Function

Private Sub Form_Load()
    Dim JB As JsonBag
    Dim i As Long
    SelectProvincias.List.MaxVisibleItems = 12
    
    If LoadJson("https://apis.datos.gob.ar/georef/api/provincias", JB) Then
        For i = 1 To JB.Item("provincias").Count
            With JB.Item("provincias").Item(i)
                SelectProvincias.List.AddItem .Item("nombre"), , .Item("id")
            End With
        Next
        SelectProvincias.List.QSort SortAscendant
    End If
End Sub

Private Sub SelectProvincias_ItemClick()
    Dim JB As JsonBag
    Dim ID As String, i As Long
    
    ID = SelectProvincias.List.ItemData(SelectProvincias.List.ListIndex)
    SelectDepartamentos.List.Clear
    SelectDepartamentos.Text = vbNullString
    SelectLocalidades.List.Clear
    SelectLocalidades.Text = vbNullString
    
    If LoadJson("https://apis.datos.gob.ar/georef/api/departamentos?provincia=" & ID, JB) Then
        For i = 1 To JB.Item("departamentos").Count
            With JB.Item("departamentos").Item(i)
                SelectDepartamentos.List.AddItem .Item("nombre"), , .Item("id")
            End With
        Next
        SelectDepartamentos.List.QSort SortAscendant
    End If
End Sub

Private Sub SelectDepartamentos_ItemClick()
    Dim JB As JsonBag
    Dim ID As String, i As Long
    
    ID = SelectDepartamentos.List.ItemData(SelectDepartamentos.List.ListIndex)

    SelectLocalidades.List.Clear
    SelectLocalidades.Text = vbNullString
    
    If LoadJson("https://apis.datos.gob.ar/georef/api/localidades?departamento=" & ID, JB) Then
        For i = 1 To JB.Item("localidades").Count
            With JB.Item("localidades").Item(i)
                SelectLocalidades.List.AddItem .Item("nombre"), , .Item("id")
            End With
        Next
        SelectLocalidades.List.QSort SortAscendant
    End If
End Sub


