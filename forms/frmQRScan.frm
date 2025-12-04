VERSION 5.00
Begin VB.Form frmQRScan 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QR - Scan Me!"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3495
   Icon            =   "frmQRScan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrQR 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblCountQR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image imgQR 
      Height          =   3255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmQRScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Project   : VB6-WhatsApp
' Form      : frmQRScan.frm
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Private m_DataQR As String
Private m_DataQRValidate As String
Private m_CountQR As Long
Private m_CloseForm As Boolean

Private Sub Form_Load()
    m_DataQR = f_GetDataQR
    Call p_QRImage(m_DataQR)
    tmrQR.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then m_CloseForm = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_CloseForm Then StateWhatsapp = STATE_SCANQR_CANCEL
    tmrQR.Enabled = False
End Sub

Private Sub tmrQR_Timer()
    m_DataQRValidate = f_GetDataQR(True)
    If m_DataQR <> m_DataQRValidate Then
        m_CountQR = m_CountQR + 1
        m_DataQR = m_DataQRValidate
        Call p_QRImage(m_DataQR)
    End If
    '---
    If f_ValidateScanQR Then
        StateWhatsapp = STATE_CHATS_LOADING
        Unload Me
        p_ValidateChatsReady
    End If
End Sub

Private Sub p_QRImage(dataQR As String)
    If Len(dataQR) > 0 Then
        imgQR.Picture = QRCodegenBarcode(dataQR, , 100)
    End If
    If m_CountQR > 0 Then
        lblCountQR.Visible = m_CountQR > 0
        lblCountQR.Caption = m_ArrayMessage(2) & "[ " & m_CountQR & " ]."
        If Me.Height < 4200 Then Me.Height = 4200
    End If
End Sub
