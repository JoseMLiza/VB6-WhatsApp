VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListEmojis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select emoji"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   Icon            =   "frmEmojiList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwEmojis 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "eng_id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "esp_id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "unicode"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmListEmojis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public emjEngId As String, emjEspId As String, emjUnicode As String
Public selOk As Boolean

Dim lvwItem As ListItem

Private Sub Form_Load()
    Dim i As Integer
    '---
    For i = 0 To UBound(arrDataEmojis)
        Set lvwItem = lvwEmojis.ListItems.Add()
        lvwItem.Text = arrDataEmojis(i, 0)
        lvwItem.SubItems(1) = arrDataEmojis(i, 1)
        lvwItem.SubItems(2) = arrDataEmojis(i, 2)
    Next
End Sub

Private Sub lvwEmojis_DblClick()
    selOk = True
    Unload Me
End Sub

Private Sub lvwEmojis_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvwEmojis
        If .SelectedItem Is Nothing Then Exit Sub
        Set lvwItem = .SelectedItem
        emjEngId = lvwItem.Text
        emjEspId = lvwItem.SubItems(1)
        emjUnicode = lvwItem.SubItems(2)
    End With
End Sub
