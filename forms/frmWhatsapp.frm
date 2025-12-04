VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmWhatsapp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envío de mensajes - Whatsapp"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17550
   Icon            =   "frmWhatsapp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeControls 
      Height          =   4800
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton btnEmojis 
         Caption         =   "emojis"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdbWhatsap 
         Caption         =   "&View capture (headless mode)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   4215
      End
      Begin VB.TextBox txtWhatsapp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtWhatsapp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton cmdbWhatsap 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdbWhatsap 
         Caption         =   "&Send"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox lstFiles 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdbWhatsap 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   4
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdbWhatsap 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   5
         Top             =   3120
         Width           =   855
      End
      Begin VB.Line lnSeparator 
         BorderColor     =   &H8000000A&
         Index           =   1
         X1              =   120
         X2              =   4320
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label lblWhatsapp 
         Caption         =   "Phone / Name:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblWhatsapp 
         Caption         =   "Message:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblWhatsapp 
         Caption         =   "Attachments:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Line lnSeparator 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   120
         X2              =   4320
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin MSComDlg.CommonDialog cmdlgFiles 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrTime 
      Left            =   600
      Top             =   4920
   End
   Begin VB.Timer tmrStatus 
      Interval        =   500
      Left            =   120
      Top             =   4920
   End
   Begin VB.PictureBox picWebView2 
      Height          =   7740
      Left            =   4800
      ScaleHeight     =   7680
      ScaleWidth      =   12600
      TabIndex        =   7
      Top             =   120
      Width           =   12660
   End
End
Attribute VB_Name = "frmWhatsapp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Project   : VB6-WhatsApp
' Form      : frmWhatsapp.frm
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Public WithEvents wvBrowser As cWebView2
Attribute wvBrowser.VB_VarHelpID = -1
'Private PageReady As Boolean
Private StateColor As OLE_COLOR

Private Function MakeLong(ByVal LowPart As Integer, ByVal HighPart As Integer) As Long
    MakeLong = LowPart Or (CLng(HighPart) * &H10000)
End Function

Private Sub btnEmojis_Click()
    Dim frm As Form
    Set frm = New frmListEmojis
    frm.Show vbModal
    '---
    With frm
        If .selOk Then
            'Debug.Print "emjEngId: " & .emjEngId
            'Debug.Print "emjEspId: " & .emjEspId
            'Debug.Print "emjUnicode: " & .emjUnicode
            txtWhatsapp(1) = txtWhatsapp(1) & IIf(Right(txtWhatsapp(1), 1) = Space(1), "", Space(1)) & .emjEspId
            txtWhatsapp(1).SetFocus
        End If
    End With
End Sub

Private Sub cmdbWhatsap_Click(Index As Integer)
    Select Case Index
        Case 0 'Add file.
            On Error GoTo ErrCancel
            cmdlgFiles.CancelError = True
            cmdlgFiles.ShowOpen
            
            mdlWhatsapp.Files.Add cmdlgFiles.FileName, cmdlgFiles.FileTitle
            lstFiles.AddItem cmdlgFiles.FileTitle
        Case 1 'Remove file
            If lstFiles.ListIndex = -1 Then Exit Sub
            mdlWhatsapp.Files.Remove lstFiles.ListIndex + 1
            lstFiles.RemoveItem lstFiles.ListIndex
            If lstFiles.ListCount = 0 Then Set mdlWhatsapp.Files = Nothing
        Case 2 'Clear file list
            Set mdlWhatsapp.Files = Nothing
            lstFiles.Clear
        Case 3 'Enviar.
            Call SendWhatsapp(txtWhatsapp(0), ReplaceEmojiInText(txtWhatsapp(1))) ' & " " & GetStrEmojiFromUnicode("U+1F600"))
        Case 4 'Hacer captura y mostrar.
            If webApp = SeleniumBasic Then
                If m_IsHeadless Then
                    With frmViewer
                        On Error Resume Next
                        Set .capturePicture = objSEWebDriver.TakeScreenShot.GetPicture
                        On Error GoTo 0
                        .Show
                    End With
                Else
                    MsgBox "This option is valid with your browser in headless mode.", vbInformation + vbOKOnly
                End If
            Else
                MsgBox "This option is valid with Selenium.", vbInformation + vbOKOnly
            End If
    End Select
ErrCancel:
End Sub

Private Sub Form_Activate()
    Select Case webApp
        Case EdgeWebView2
            If wvBrowser Is Nothing Then
                '---
                Set wvBrowser = New_c.WebView2
                '---
                m_UserDataDir = App.Path & WebView2UserDataPath
                '---
                With wvBrowser
                    If .BindTo(picWebView2.hwnd, , , m_UserDataDir) = 0 Then
                        MsgBox "No se pudo inicializar WebView-Binding"
                        Exit Sub
                    End If
                    '---
                    On Error Resume Next
                    '.EnableWindow wvBrowser.HosthWnd, 0
                End With
                'Redimensionar control y formulario
                fmeControls.Height = 280
                'Abrir WhatsApp Web con WebView2
                If m_PageIsReady = False Then Call p_WebView2OpenWhatsappWeb(wvBrowser)
            End If
        Case SeleniumBasic
            If Not m_PageIsReady Then
                'Ocultar picturebox del WebView2 y redimencionar el formulario
                picWebView2.Visible = False
                With Me
                    .Height = IIf(m_IsHeadless, 5355, 4740)
                    .Width = 4800
                End With
                fmeControls.Height = IIf(m_IsHeadless, 320, 280)
                'Abrir WhatsApp Web con Selenium
                p_SEOpenWhatsappWeb objSEWebDriver
            End If
    End Select
End Sub

Private Sub Form_Load()
    'Inicializar los parametros del modulo **mdlWhatsapp**
    Call p_Initialize(LNG_ENGLISH)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case webApp
        Case EdgeWebView2
            '---
        Case SeleniumBasic
            If ValidateSelenium(objSEWebDriver) Then
                On Error Resume Next
                objSEWebDriver.Close
                objSEWebDriver.Quit
                On Error GoTo 0
                End
            End If
    End Select
End Sub

Private Sub picWebView2_Resize()
    If Not wvBrowser Is Nothing Then
        wvBrowser.SyncSizeToHostWindow
    End If
End Sub

Private Sub tmrStatus_Timer()
    Select Case StateWhatsapp
        Case STATE_PAGE_LOAD, STATE_PAGE_READY, STATE_PAGE_VALIDATE, STATE_CHATS_LOADING, STATE_CHATS_READY
            lblStatus = m_ArrayStatus(StateWhatsapp)
            StateColor = &H8000&
        Case STATE_CHATS_READY
            tmrTime.Tag = vbNullString
            tmrStatus.Interval = 0
        Case STATE_WAITQR
            StateColor = &H8000&
            tmrTime.Interval = 1000
            If tmrTime.Tag = vbNullString Then tmrTime.Tag = Now
        Case STATE_WAITSCANQR
            lblStatus = m_ArrayStatus(StateWhatsapp)
            StateColor = &H8000&
            tmrTime.Interval = 0
            If Not tmrTime.Tag = vbNullString Then tmrTime.Tag = vbNullString
        Case STATE_SCANQR_CANCEL
            lblStatus = m_ArrayStatus(StateWhatsapp)
            StateColor = &HC0&
            tmrTime.Tag = vbNullString
    End Select
    lblStatus.ForeColor = StateColor
End Sub

Private Sub tmrTime_Timer()
    If tmrTime.Tag <> vbNullString Then
        lblStatus.Caption = m_ArrayStatus(StateWhatsapp) & " - " & Format(Now - CDate(tmrTime.Tag), "hh:mm:ss")
    Else
        tmrTime.Interval = 0
    End If
End Sub

Private Sub wvBrowser_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    m_PageIsReady = IsSuccess
    If m_PageIsReady And StateWhatsapp = STATE_PAGE_LOAD Then
        StateWhatsapp = STATE_PAGE_READY
        Call p_ValidateAccess
    End If
End Sub

Private Sub wvBrowser_NavigationStarting(ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal URI As String, Cancel As Boolean)
    m_PageIsReady = False
End Sub

