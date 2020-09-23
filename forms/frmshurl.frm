VERSION 5.00
Begin VB.Form frmshurl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Internet Shortcut"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcan 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3510
      TabIndex        =   4
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2115
      TabIndex        =   3
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox txturl 
      Height          =   300
      Left            =   1455
      TabIndex        =   2
      Top             =   1080
      Width           =   3225
   End
   Begin VB.TextBox txtlnkname 
      Height          =   300
      Left            =   1455
      TabIndex        =   1
      Top             =   675
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut URL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut Title:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmshurl.frx":0000
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmshurl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcan_Click()
    txturl.Text = ""
    txtlnkname.Text = ""
    Unload frmshurl
End Sub

Private Sub cmdok_Click()
    If Len(Trim(txtlnkname.Text)) <= 0 Then
        MsgBox "You must enter in a title for the shortcut.", vbInformation, frmshurl.Caption
        txtlnkname.Text = ""
        Exit Sub
    ElseIf Len(Trim(txturl.Text)) <= 0 Then
        MsgBox "You must enter in the URL address of the site.", vbInformation, frmshurl.Caption
        txturl.Text = ""
        Exit Sub
    Else
        SaveURLShortCut txturl.Text, txtlnkname.Text
        cmdcan_Click
    End If
    
End Sub

Private Sub Form_Load()
    MakeFlatControls frmshurl
    txtlnkname.Text = TBookMark
    txturl.Text = TBookURL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmshurl = Nothing
End Sub

Private Sub txtlnkname_GotFocus()
    txtlnkname.BackColor = Config.Hightlight
End Sub

Private Sub txtlnkname_LostFocus()
    txtlnkname.BackColor = vbWhite
End Sub

Private Sub txturl_GotFocus()
    txturl.BackColor = Config.Hightlight
End Sub

Private Sub txturl_LostFocus()
    txturl.BackColor = vbWhite
End Sub
