VERSION 5.00
Begin VB.Form fmrOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browser Selection"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   Icon            =   "fmrOption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHeadLess 
      Caption         =   "Headless mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4890
      TabIndex        =   14
      Top             =   2940
      Width           =   1935
   End
   Begin VB.PictureBox picOptSEWebDriver 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   2505
      TabIndex        =   5
      Top             =   2865
      Width           =   2535
      Begin VB.OptionButton optWebDriver 
         Appearance      =   0  'Flat
         Caption         =   "Edge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optWebDriver 
         Appearance      =   0  'Flat
         Caption         =   "Chrome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "fmrOption.frx":10F2
      Top             =   3480
      Width           =   6975
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "fmrOption.frx":11E5
      Top             =   2040
      Width           =   6015
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "fmrOption.frx":12B2
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "fmrOption.frx":1381
      Top             =   480
      Width           =   7095
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Welcome to our WhatsApp messaging app!"
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdbOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton optWebApp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optWebApp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblTexto 
      Caption         =   "Webdriver:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      Caption         =   "Developed by: Jose Liza"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000A&
      Index           =   1
      X1              =   16
      X2              =   480
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000A&
      Index           =   0
      X1              =   16
      X2              =   480
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Image imgOption 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "fmrOption.frx":13EB
      Stretch         =   -1  'True
      Top             =   2100
      Width           =   480
   End
   Begin VB.Image imgOption 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "fmrOption.frx":189E
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   480
   End
End
Attribute VB_Name = "fmrOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Project   : VB6-WhatsApp
' Form      : fmrOption.frm
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Private Sub chkHeadLess_Click()
    m_IsHeadless = CBool(chkHeadLess.Value)
End Sub

Private Sub cmdbOk_Click()
    Dim bExito As Boolean, bUpdate As Boolean, bError As Boolean
    Dim pathDownloadFile As String, pathUnzipFile As String
    If optWebApp(0).Value Then 'Microsoft Edge WebView2
        webApp = EdgeWebView2
        SEWebDriver = None
        bExito = WebAppConstruct(EdgeWebView2)
    Else 'SeleniumBasic
        webApp = SeleniumBasic
        SEWebDriver = IIf(optWebDriver(0).Value, ChromeDriver, EdgeDriver)
        m_IsHeadless = CBool(chkHeadLess.Value)
        bExito = WebAppConstruct(SeleniumBasic)
        bUpdate = UpdateWebDriver(SEWebDriver, pathDownloadFile, bError)
    End If
    If bExito Then
        If bError Then
            Exit Sub
        End If
        If bUpdate Then
            If Not pathDownloadFile = vbNullString Then
                If FileExists(pathDownloadFile) Then
                    Call UnzipFile(pathDownloadFile, App.Path & tempDownload, True, "exe", pathUnzipFile)
                    Call MoveFileToPath(pathUnzipFile, Left(App.Path & SeleniumBase, Len(App.Path & SeleniumBase) - 1), IIf(SEWebDriver = EdgeDriver, "edgedriver.exe", ""), True)
                Else
                    MsgBox "Downloaded file does not exist.", vbCritical + vbOKOnly
                    Exit Sub
                End If
            Else
                MsgBox "Download not complete!", vbInformation + vbOKOnly
                Exit Sub
            End If
        Else
            Select Case m_ExitCodeUpdate
                Case 1
                    MsgBox "Download canceled by user.", vbInformation + vbOKOnly
                Case 2
                    MsgBox "URL address is not valid.", vbExclamation + vbOKOnly
                Case 3
                    MsgBox "Download URL is inaccessible.", vbCritical + vbOKOnly
                Case 4
                    MsgBox "An error occurred during the download.", vbCritical + vbOKOnly
            End Select
            If m_ExitCodeUpdate > 0 Then Exit Sub
        End If
        frmWhatsapp.Show
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    m_IsCompiled = CBool(App.LogMode)
    '---
    Call LoadDataEmojis
    '---
End Sub

Private Sub optWebApp_Click(Index As Integer)
    picOptSEWebDriver.Enabled = optWebApp(1).Value
End Sub

Private Sub optWebDriver_Click(Index As Integer)
    SEWebDriver = IIf(optWebDriver(0).Value, ChromeDriver, EdgeDriver)
End Sub
