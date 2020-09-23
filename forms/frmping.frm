VERSION 5.00
Begin VB.Form frmping 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ping"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdping 
      Caption         =   "&Ping"
      Height          =   350
      Left            =   150
      TabIndex        =   3
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton cmdclipcopy 
      Caption         =   "&Copy to clipboard"
      Height          =   350
      Left            =   1560
      TabIndex        =   4
      Top             =   3300
      Width           =   1920
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "C&lose"
      Height          =   350
      Left            =   3705
      TabIndex        =   5
      Top             =   3300
      Width           =   1215
   End
   Begin VB.TextBox txtresponse 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2130
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   660
      Width           =   6630
   End
   Begin VB.TextBox txthost 
      Height          =   285
      Left            =   1410
      TabIndex        =   1
      Top             =   225
      Width           =   4365
   End
   Begin VB.Label blstat 
      AutoSize        =   -1  'True
      Caption         =   "Ide"
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
      Left            =   750
      TabIndex        =   7
      Top             =   2925
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status :"
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
      Left            =   150
      TabIndex        =   6
      Top             =   2925
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hostname or IP"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "frmping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclipcopy_Click()
    If Len(Trim(txtresponse.Text)) <= 0 Then
        MsgBox "There is nothing yet to copy", vbInformation, frmping.Caption
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText txtresponse.Text
        MsgBox "The information above has been successfully copied to the clipboard.", vbInformation, frmping.Caption
    End If
End Sub

Private Sub cmdclose_Click()
    txthost.Text = ""
    txtresponse.Text = ""
    Unload frmping
End Sub

Private Sub cmdping_Click()
    txtresponse.Text = ""
    If Trim(Len(txthost.Text)) <= 0 Then
        blstat.Caption = "Error no IP or Hostname found."
        Exit Sub
    Else
        blstat.Caption = "Pinging " & txthost.Text & " please wait....."
        txtresponse.Text = PingHost(txthost.Text)
        blstat.Caption = "Finished."
    End If
End Sub

Private Sub Form_Load()
    frmping.Icon = Nothing
    MakeFlatControls frmping
    frmping.Caption = "Ping " & TBookMark
    txthost.Text = PhaseDomain(TBookURL)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmping = Nothing
End Sub

Private Sub txthost_GotFocus()
    txthost.BackColor = Config.Hightlight
End Sub

Private Sub txthost_LostFocus()
    txthost.BackColor = vbWhite
End Sub

