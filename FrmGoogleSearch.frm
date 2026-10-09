VERSION 5.00
Begin VB.Form FrmGoogleSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar imagenes en Google"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8385
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicRec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8C29&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   1200
      ScaleHeight     =   1785
      ScaleWidth      =   5745
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label LblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Habla ahora"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.PictureBox PicButton 
      BackColor       =   &H00DEDEDE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   7200
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   6
      Tag             =   "Cancelar"
      Top             =   6210
      Width           =   975
   End
   Begin VB.PictureBox PicButton 
      BackColor       =   &H00FF8439&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   5880
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   5
      Tag             =   "Aceptar"
      Top             =   6210
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   6240
      Width           =   5445
   End
   Begin Proyecto1.ucList ucList2 
      Height          =   480
      Left            =   5880
      TabIndex        =   3
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   210
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   5655
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   0
         Width           =   4815
      End
      Begin VB.Image ImgMicrophone 
         Height          =   360
         Left            =   5280
         Picture         =   "FrmGoogleSearch.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   30
         Picture         =   "FrmGoogleSearch.frx":0702
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
   End
   Begin Proyecto1.ucList ucList1 
      Height          =   3840
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7530
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
      SelectionColor  =   16745529
      LineColor       =   13487565
      ListStyle       =   1
   End
End
Attribute VB_Name = "FrmGoogleSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Const CRYPT_STRING_BASE64           As Long = &H1
Private Const CRYPT_STRING_NOCRLF           As Long = &H40000000
Private Const CRYPT_STRING_NOCR             As Long = &H80000000

Private Declare Function CryptBinaryToStringA Lib "crypt32.dll" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, ByVal pcchString As Long) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long) As Long

Dim cShadowTxt(1) As ClsShadow
Dim cShadowButton(1) As ClsShadow
Dim ImgIndex As Long
Dim PagePos As Long

Private Sub Search()
    Dim JB As JsonBag
    Dim ID As String, i As Long
    Dim sUrl As String
 
    sUrl = "https://www.googleapis.com/customsearch/v1?googlehost=google.com&lr=lang_es&imgType=photo&imgSize=large&num=10&searchType=image&key=AIzaSyCB6jyUuD6pHJOW5KldSA4W2RXjBYzlX8Y&cx=733d3ac8912e284d5&start=" & PagePos & "&q="
    sUrl = sUrl & Text1(0).Text
    Screen.MousePointer = vbArrowHourglass
    If LoadJson(sUrl, JB) Then
        If JB.Exists("items") Then
            For i = 1 To JB.Item("items").Count
                With JB.Item("items").Item(i)
                    If ucList1.ImageListAddImage(.Item("image").Item("thumbnailLink")) > 0 Then
                        ucList1.AddItem .Item("title"), ImgIndex, .Item("link"), .Item("displayLink")
                        ImgIndex = ImgIndex + 1
                    End If
                End With
            Next
            ucList1.Update ' True
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

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
    Dim i As Long
    
    With ucList1
        .ImageListInitialize 80, 80
        .ImagesRadius = 10
        .ItemHeight = 80
        .SubTitleVisible = True
    End With
    
    For i = 0 To 1
        Set cShadowTxt(i) = New ClsShadow
        With cShadowTxt(i)
            .Margin = 5
            .ShadowColor = vbActiveBorder ' &HFF9033
            .ShadowOpacity = 20
            .ShadowSize = 2
            If i = 0 Then
                .ShowBorders Picture1.hWnd, True
            Else
                .ShowBorders Text1(i).hWnd, True
            End If
        End With
    Next
    
    With ucList2
        .MaxVisibleItems = 8
        .ListWidth = Text1(0).Width
        .AutoCompleteTextBox Text1(0).hWnd, -15, 10, , -Text1(0).Top - 4
    End With
    
    For i = 0 To 1
       Set cShadowButton(i) = New ClsShadow
       
        With cShadowButton(i)
            .Margin = 2
            .BackColor = PicButton(i).BackColor
            .BorderColor = PicButton(i).BackColor
            .ShadowColor = &HFF9033
            .BorderRadius = 5
            .BorderWidth = 1
            .ShadowOpacity = 0
            .ShadowSize = 2
            .ShowBorders PicButton(i).hWnd, True
        End With
    Next
End Sub

Private Sub Form_Resize()
    With ucList1
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top - Text1(1).Height * 3
    End With
End Sub


Private Sub Image2_Click()

End Sub

Private Sub ImgMicrophone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRet As Long
    Const MCIERR_WAVE_INPUTSUNSUITABLE = 328
    Beep
    ImgMicrophone.Tag = 0
    lRet = RecAudio()
    Select Case lRet
        Case 0
            LblInfo.Caption = "Manten presionado y habla."
            PicRec.BackColor = &HFF8439
            ImgMicrophone.Tag = 1
        Case MCIERR_WAVE_INPUTSUNSUITABLE
            LblInfo.Caption = "Conecta un microfono."
            PicRec.BackColor = &H8484FF
        Case Else
            LblInfo.Caption = "Error: " & lRet
            PicRec.BackColor = &H8484FF
    End Select
    PicRec.Visible = True
End Sub

Private Sub ImgMicrophone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sRet As String
    Dim sWavBase64 As String
    
    If ImgMicrophone.Tag = 0 Then
        CancelAudio
        PicRec.Visible = False
        Exit Sub
    End If
    
    LblInfo.Caption = "Prosesando..."
    DoEvents
    
    If GetAudio(sWavBase64) Then
        sRet = SpeechToText(sWavBase64)
    End If
    
    If Len(sRet) Then
        PicRec.Visible = False
        Text1(0).Text = sRet
        Text1_KeyPress 0, vbKeyReturn
    Else
        LblInfo.Caption = "Sin resultados."
        DoEvents
        Sleep 500
        PicRec.Visible = False
    End If
    
End Sub

Private Sub PicButton_Click(Index As Integer)
    If Index = 0 Then
        If ucList1.ListIndex > -1 Then
            MsgBox ucList1.ItemData(ucList1.ListIndex)
        Else
            Beep
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub PicButton_GotFocus(Index As Integer)
    cShadowButton(Index).BorderColor = &HFF9033
    cShadowButton(Index).ShadowOpacity = 30
    cShadowButton(Index).Refresh
End Sub

Private Sub PicButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then PicButton_Click Index
End Sub

Private Sub PicButton_LostFocus(Index As Integer)
    cShadowButton(Index).BorderColor = PicButton(Index).BackColor
    cShadowButton(Index).ShadowOpacity = 0
    cShadowButton(Index).Refresh
End Sub

Private Sub PicButton_Paint(Index As Integer)
    With PicButton(Index)
        .CurrentX = .ScaleWidth / 2 - .TextWidth(.Tag) / 2
        .CurrentY = .ScaleHeight / 2 - .TextHeight(.Tag) / 2
        PicButton(Index).Print .Tag
    End With
End Sub

Private Sub Text1_Change(Index As Integer)
    Dim JB As JsonBag
    Dim ID As String, i As Long
    Dim sUrl As String
    
    If Index = 0 Then
        sUrl = "http://suggestqueries.google.com/complete/search?output=firefox&client=firefox&&ds=yt&hl=eses&q=" & Text1(0).Text
        
        ucList2.Clear
        If LoadJson(sUrl, JB) Then
          
            For i = 1 To JB.Item(2).Count
                ucList2.AddItem JB.Item(2).Item(i)
            Next
            'ucList2.Update True
            
        End If
    End If
    
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = vbKeyDown Then
        If ucList1.ListCount > 0 Then
            ucList1.ListIndex = ucList1.ListIndex + 1
        End If
        ucList1.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And KeyAscii = 13 Then
        KeyAscii = 0
        ucList1.Clear
        ucList1.ImageListClear
        ImgIndex = 0
        PagePos = 1
        Search
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With cShadowTxt(Index)
        .BorderColor = &HFF9033
        .ShadowColor = &HFF9033
        .Refresh
    End With
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    With cShadowTxt(Index)
        .BorderColor = vbActiveBorder
        .ShadowColor = vbActiveBorder
        .Refresh
    End With
End Sub

Private Sub ucList1_IndexChange()
    Text1(1).Text = ucList1.ItemData(ucList1.ListIndex)
End Sub

Private Sub ucList1_ItemClick()
'    MsgBox ucList1.ItemData(ucList1.ListIndex)
End Sub

Private Sub ucList1_ScrollBottom()
    PagePos = PagePos + 10
    Call Search
End Sub

Private Sub ucList2_PostAutoComplete()
    Text1_KeyPress 0, vbKeyReturn
End Sub

Private Function SpeechToText(sWavBase64 As String) As String
    Dim sUrl As String
    Dim JB As JsonBag
    Dim response As String
    Dim xmlhttp As Object
    Dim sText As String
    Dim i As Long
    Dim sBody As String

    If Len(sWavBase64) = False Then Exit Function

    sUrl = "https://speech.googleapis.com/v1p1beta1/speech:recognize?key=AIzaSyBqKVSTAwPrDBF6LxXxCchDg6hsvHCaL-M"
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "POST", sUrl, False
    sBody = "{'audio': {'content':'" & sWavBase64 & "'},'config':{'enableAutomaticPunctuation':true,'encoding':'LINEAR16','languageCode':'es-ES','model':'default'}}"
    xmlhttp.send sBody

    Set JB = New JsonBag
    
    JB.JSON = xmlhttp.responseText
    
    If Not JB.Exists("results") Then Exit Function
    
    For i = 1 To JB.Item("results").Count
         sText = sText & JB.Item("results")(i).Item("alternatives")(1).Item("transcript") & vbCrLf
    Next
    Set xmlhttp = Nothing
    SpeechToText = sText

End Function

Private Function RecAudio() As Long
    Dim COMANDO As String
    Const BITS As Integer = 16
    Const CANALES As Integer = 1
    Const MUESTRAS As Long = 11025
    Const PROMEDIO As Long = BITS * CANALES * MUESTRAS / 8
    Const ALINEACION As Long = BITS * CANALES / 8
    
    COMANDO = "set capture bitspersample " & BITS & " channels " & CANALES & " alignment " & ALINEACION & " samplespersec " & _
    MUESTRAS & " bytespersec " & PROMEDIO & " format tag pcm wait"
    Call mciSendString("close capture", "", 0, 0)
    Call mciSendString("open new type waveaudio alias capture", "", 0, 0)
    Call mciSendString(COMANDO, "", 0, 0)
    RecAudio = mciSendString("record capture", "", 0, 0)
End Function

Private Function CancelAudio()
    Call mciSendString("stop capture", 0&, 0, 0)
    Call mciSendString("close capture", 0&, 0, 0)
End Function

Private Function GetAudio(sRetBase64 As String) As Boolean
    On Error Resume Next
    Dim bData() As Byte
    Dim FF As Integer
    Dim sPath As String
    sPath = Environ$("Temp") & "\temp.wav"
    Call mciSendString("stop capture", 0&, 0, 0)
    Call mciSendString("save capture " & sPath, 0&, 0, 0)
    Call mciSendString("close capture", 0&, 0, 0)
    
    If Len(Dir(sPath)) = 0 Then Exit Function
    
    FF = FreeFile
    Open sPath For Binary As FF
        ReDim bData(LOF(FF) - 1)
        Get FF, , bData
    Close FF
    
    Kill sPath
    If Base64Encode(bData, sRetBase64) Then
        GetAudio = Err.Number = 0
    End If
End Function

Private Function Base64Encode(ByRef bvSrc() As Byte, ByRef sOut As String) As Boolean
   
    Dim lLenOut                 As Long
    '// calculate buffer len
    Call CryptBinaryToStringA(VarPtr(bvSrc(0)), _
                              UBound(bvSrc) + 1, _
                              CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, _
                              vbNullString, _
                              VarPtr(lLenOut))
    If lLenOut = 0 Then
        Exit Function
    End If

    sOut = Space$(lLenOut - 1)

    '// now convert to base64
    Call CryptBinaryToStringA(VarPtr(bvSrc(0)), _
                              UBound(bvSrc) + 1, _
                              CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, _
                              sOut, _
                              VarPtr(lLenOut))
    Base64Encode = True
End Function


