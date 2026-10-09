VERSION 5.00
Begin VB.Form FrmAutocomplete 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucList ucList2 
      Height          =   480
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmAutocomplete.frx":0000
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
   End
   Begin Proyecto1.ucList ucList1 
      Height          =   480
      Left            =   360
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      font            =   "FrmAutocomplete.frx":0028
      backcolor       =   -2147483643
      findtextcolor   =   -2147483635
      selectioncolor  =   -2147483635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   1560
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   7575
      Begin Proyecto1.ucSelect ucSelect1 
         Height          =   540
         Left            =   600
         TabIndex        =   4
         Top             =   6000
         Width           =   6255
         _extentx        =   11033
         _extenty        =   953
         backcolor       =   -2147483643
         font            =   "FrmAutocomplete.frx":0050
      End
      Begin VB.PictureBox PicButton 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2760
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   10
         Top             =   7200
         Width           =   2175
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "SEND INFO"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.PictureBox PicBackTxt 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   3
         Left            =   600
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   417
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4680
         Width           =   6255
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.PictureBox PicBackTxt 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   2
         Left            =   600
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   417
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6255
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.PictureBox PicBackTxt 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   1
         Left            =   600
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   417
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6255
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.PictureBox PicBackTxt 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   0
         Left            =   600
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   417
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   6255
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   15
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Business Email"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmAutocomplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cShadowBack As ClsShadow
Dim cShadowButton As ClsShadow
Dim CollShadow As Collection


Private Sub Form_Load()
    Dim cShadow As ClsShadow
    Dim i As Long
    
    With ucList1
        .ImageListInitialize 16, 16
        .UseMonocromeIcons = True
        .ItemHeight = 32
        .MaxVisibleItems = 8
        .ImageListAddImage App.Path & "\Vectores PNG\8.png"
        .Shadow.ShadowOffsetY = 2
        
    End With
    
    With ucList2
        .ItemHeight = 32
        .MaxVisibleItems = 8
        .ListWidth = ucSelect1.Width / .GetWindowDpi
        .Shadow.ShadowOffsetY = 2
    End With
    
    
    Set cShadowBack = New ClsShadow
    With cShadowBack
        .Margin = 4
        .BorderColor = vbWhite
        .BorderRadius = 5
        .ShadowSize = 15
        .ShowBorders Picture1.hWnd, True
    End With
    
    Set CollShadow = New Collection
    For i = 0 To 3
        Set cShadow = New ClsShadow
        With cShadow
            .ShadowSize = 0
            .Margin = 3
            .ShadowColor = &HFF9033
            .ShadowOpacity = 0
            .ShadowSize = 5
            .ShowBorders PicBackTxt(i).hWnd, True
        End With
        CollShadow.Add cShadow
    Next
    
    Set cShadowButton = New ClsShadow
    With cShadowButton
        .Margin = 4
        .BackColor = vbRed
        .BorderRadius = 3
        .BorderColor = vbRed
        .BorderWidth = 0
        .ShadowSize = 10
        .ShowBorders PicButton.hWnd, True
    End With

    LoadCountrys
    LoadNames

End Sub

Private Sub Form_Resize()
    Picture1.Left = Me.ScaleWidth / 2 - Picture1.ScaleWidth / 2
    cShadowBack.Refresh
End Sub

Private Sub PicButton_GotFocus()
    cShadowButton.BackColor = &HFF9033
    cShadowButton.Refresh
End Sub

Private Sub PicButton_LostFocus()
    cShadowButton.BackColor = vbRed
    cShadowButton.Refresh
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With CollShadow(Index + 1)
         .ShadowOpacity = 30
        .BorderColor = &HFF9033
        .Refresh
    End With
    ucList1.AutoCompleteTextBox Text1(Index).hWnd, -Text1(Index).Left - 4, Text1(Index).Top * 2, Text1(Index).Left * 2 + 4, -Text1(Index).Top - 4
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    With CollShadow(Index + 1)
        .ShadowOpacity = 0
        .BorderColor = vbActiveBorder
        .Refresh
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cShadowBack = Nothing
    Set CollShadow = Nothing
End Sub

Private Sub LoadNames()
    Dim MyLine As String
    Dim FF As Integer

    With ucList1
        .Clear
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
            .Update
        Close FF
    End With
End Sub

Private Sub LoadCountrys()
    Dim MyLine As String
    Dim FF As Integer

    With ucSelect1.List
        .Clear
        FF = FreeFile
        Open App.Path & "\Country.txt" For Input As #FF
            Do While Not EOF(FF)
                Line Input #FF, MyLine
                .AddItem MyLine
            Loop
        Close FF
    End With
End Sub



