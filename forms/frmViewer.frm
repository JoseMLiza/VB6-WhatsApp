VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Headless viewer"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   12240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdbRefresh 
      Caption         =   "&Refresh"
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
      Left            =   10680
      TabIndex        =   0
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Image imgCapture 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   9000
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Project   : VB6-WhatsApp
' Form      : frmViewer.frm
' Author    : Jose Liza (https://github.com/JoseMLiza)
'
'=========================================================================

Option Explicit

Public capturePicture As StdPicture

Private Type RECT
    X As Long
    Y As Long
    Rigth As Long
    Bottom As Long
End Type

Private m_MinHeight As Long
Private m_MinWidth  As Long
Private m_ImgRECT   As RECT
Private m_CmdbRECT  As RECT

Private Sub cmdbRefresh_Click()
    If Not objSEWebDriver Is Nothing Then
        If (webApp = SeleniumBasic) And m_IsHeadless Then
            imgCapture.Picture = objSEWebDriver.TakeScreenShot.GetPicture
        End If
    End If
End Sub

Private Sub Form_Load()
    m_MinHeight = 2040
    m_MinWidth = 3135
    '--
    With m_ImgRECT
        .X = imgCapture.Left
        .Y = imgCapture.Top
        .Rigth = Me.Width - (imgCapture.Left + imgCapture.Width)
        .Bottom = Me.Height - (imgCapture.Top + imgCapture.Height)
    End With
    With m_CmdbRECT
        .X = Me.Width - cmdbRefresh.Left
        .Y = Me.Height - cmdbRefresh.Top
        .Rigth = Me.Width - (cmdbRefresh.Left + cmdbRefresh.Width)
        .Bottom = Me.Height - (cmdbRefresh.Top + cmdbRefresh.Height)
    End With
    '---
    If Not capturePicture Is Nothing Then imgCapture.Picture = capturePicture
End Sub

Private Sub Form_Resize()
    If Me.Width >= m_MinWidth Then
        With imgCapture
            .Left = m_ImgRECT.X
            .Width = Me.Width - m_ImgRECT.Rigth - m_ImgRECT.X
        End With
        With cmdbRefresh
            .Left = Me.Width - m_CmdbRECT.X
            .Width = Me.Width - m_CmdbRECT.Rigth - .Left
        End With
    End If
    '---
    If Me.Height >= m_MinHeight Then
        With imgCapture
            .Top = m_ImgRECT.Y
            .Height = Me.Height - m_ImgRECT.Bottom - m_ImgRECT.Y
        End With
        With cmdbRefresh
            .Top = Me.Height - m_CmdbRECT.Y
            .Height = Me.Height - m_CmdbRECT.Bottom - .Top
        End With
    End If
End Sub
