VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update downloader"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgbDownload 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdbCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Downloading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "http://"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Project   : WebDriver-Update
' Form      : frmDownload.frm
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Private m_StartDownload     As Boolean
Private m_DownloadComplete  As Boolean
Private m_TotalBytes        As Long
Private m_DestPath          As String
Private m_BytesDownloaded   As Long
Private m_Progress          As Integer
Private m_fileNumber        As Integer
Private m_StartTime         As Double
Private m_CurrentTime       As Double
Private m_ElapsedTime       As Double
Private m_DownloadSpeed     As Double
'Private m_DataDownloaded()  As Byte

Private WithEvents cWinHttp As clsWinHttp
Attribute cWinHttp.VB_VarHelpID = -1

Private Sub cmdbCancel_Click()
    If m_StartDownload Or Not m_DownloadComplete Then
        If MsgBox("Download is in process." & vbCrLf & vbCrLf & "Do you want to cancel the download?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        m_ExitCode = Cancel
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub DownloadFile(ByVal url As String, ByVal savePath As String)
    m_DestPath = savePath
    cWinHttp.SendRequest url, "", METHOD_GET, True, True
    m_StartTime = Timer
End Sub

Private Sub cWinHttp_RequestStart(ByVal Status As Long, ByVal ContentType As String)
    On Error GoTo SubError
    '---
    If Not m_DestPath = vbNullString Then
        m_TotalBytes = cWinHttp.GetResponseHeader("Content-Length")
        'Abrir el archivo de destino para escritura binaria
        m_fileNumber = FreeFile
        Open m_DestPath For Binary Access Write As #m_fileNumber
        'Cerrar el archivo inmediatamente
        Close #m_fileNumber
    End If
    Exit Sub
    '---
SubError:
    Call MessageError(Err)
End Sub

Private Sub cWinHttp_ResponseDataAvailable(Data() As Byte)
    On Error GoTo SubError
    Dim strSpeed As String
    '---
    If m_TotalBytes = 0 Then Exit Sub
    'Escribir los datos descargados en el archivo de destino
    m_fileNumber = FreeFile
    Open m_DestPath For Binary Access Write As #m_fileNumber
    'Mover el puntero de archivo al final
    Seek #m_fileNumber, LOF(m_fileNumber) + 1
    'Escribir los datos en el archivo
    Put #m_fileNumber, , Data
    'Cerrar el archivo
    Close #m_fileNumber
    
    
    'Actualizar el progreso de la descarga
    m_BytesDownloaded = m_BytesDownloaded + UBound(Data) + 1
    If m_TotalBytes > 0 Then
        m_Progress = (m_BytesDownloaded / m_TotalBytes) * 100
    Else
        m_Progress = 100 'Si el tamaño total no está disponible, consideramos la descarga como completa
    End If
    
    'Calcular la velocidad de descarga
    m_CurrentTime = Timer
    m_ElapsedTime = m_CurrentTime - m_StartTime
    If m_ElapsedTime > 0 Then
        m_DownloadSpeed = (m_BytesDownloaded * 8) / (m_ElapsedTime * 1024) 'Kbps
        If m_DownloadSpeed < 1000 Then
            strSpeed = Format(m_DownloadSpeed, "0.00") & " Kb/s"
        Else
            m_DownloadSpeed = m_DownloadSpeed / 1000 'Mbps
            strSpeed = Format(m_DownloadSpeed, "0.00") & " Mb/s"
        End If
    End If
    
    'Determinar la unidad de medida del progreso
    If m_TotalBytes < KB Then
        lblInfo(1) = "Downloading " & m_BytesDownloaded & " bytes" & " of " & m_TotalBytes & " Bytes (" & strSpeed & ")"
    ElseIf m_TotalBytes < MB Then
        lblInfo(1) = "Downloading " & Format(m_BytesDownloaded / KB, "0.00") & " KB" & " of " & Format(m_TotalBytes / KB, "0.00") & " KB (" & strSpeed & ")"
    ElseIf m_TotalBytes < GB Then
        lblInfo(1) = "Downloading " & Format(m_BytesDownloaded / MB, "0.00") & " MB" & " of " & Format(m_TotalBytes / MB, "0.00") & " MB (" & strSpeed & ")"
    Else
        lblInfo(1) = "Downloading " & Format(m_BytesDownloaded / GB, "0.00") & " GB" & " of " & Format(m_TotalBytes / GB, "0.00") & " GB (" & strSpeed & ")"
    End If
    'Actualizar la barra de progreso
    pgbDownload.value = m_Progress
    lblPercent = m_Progress & "%"
    
    If m_BytesDownloaded >= m_TotalBytes Then
        lblInfo(1).Caption = "Download complete!"
        If Not m_ArgUnZipPath = vbNullString Then
            lblInfo(1).Caption = "Extract file..."
            Call UnZipFile(m_ArgDownloadFile, m_ArgUnZipPath)
            lblInfo(1).Caption = "Extract complete..."
        End If
        cmdbCancel.Caption = "&Exit"
        m_ExitCode = Ok
        m_DownloadComplete = True
        If m_DownloadComplete Then Unload Me
    End If
    Exit Sub
    '---
SubError:
    Call MessageError(Err)
End Sub

Private Sub Form_Activate()
    'Validar disponibilidad del enlace
    lblInfo(1) = "Validating server..."
    If Not cWinHttp.IsURLAccessible(m_ArgDownloadUrl) Then
        m_ExitCode = InaccessibleURL
        lblInfo(1) = "Server Not Found!"
        MsgBox "Download link not available.", vbCritical + vbOKOnly
        MsgBox "The application will be closed.", vbInformation + vbOKOnly
        Unload Me
    Else
        'lblInfo(1) = "Downloading..."
        Call ProgressBackColor(pgbDownload.hwnd)
        m_StartDownload = True
        lblPercent.Visible = m_StartDownload
        If Not FolderExists(m_ArgDownloadFolder) Then
            Call MkDir(m_ArgDownloadFolder)
        End If
        m_ArgDownloadFile = m_ArgDownloadFolder & "\" & GetFileNameFromURL(m_ArgDownloadUrl)
        Call DownloadFile(m_ArgDownloadUrl, m_ArgDownloadFile)
    End If
End Sub

Private Sub Form_Load()
    Set cWinHttp = New clsWinHttp
    pgbDownload.value = 0
    lblInfo(0) = m_ArgDownloadUrl
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cWinHttp = Nothing
End Sub

