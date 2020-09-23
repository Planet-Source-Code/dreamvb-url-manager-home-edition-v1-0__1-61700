VERSION 5.00
Begin VB.Form frmref 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Share Bookmark"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtname2 
      Height          =   300
      Left            =   1635
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2130
      Width           =   1650
   End
   Begin VB.TextBox txtname1 
      Height          =   300
      Left            =   1635
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1155
      Width           =   1650
   End
   Begin Project1.smtp smtp1 
      Left            =   225
      Top             =   4485
      _ExtentX        =   1058
      _ExtentY        =   1164
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Send"
      Height          =   350
      Left            =   1305
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&ancel"
      Height          =   350
      Left            =   4050
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear all"
      Height          =   350
      Left            =   2655
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtmailfrom 
      Height          =   300
      Left            =   1635
      TabIndex        =   3
      Top             =   1665
      Width           =   3510
   End
   Begin VB.TextBox txtmailto 
      Height          =   300
      Left            =   1635
      TabIndex        =   1
      Top             =   720
      Width           =   3480
   End
   Begin VB.TextBox txtbody 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   210
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2820
      Width           =   5010
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmref.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NB max length 255 characters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1920
      TabIndex        =   13
      Top             =   2550
      Width           =   2310
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
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
      Left            =   225
      TabIndex        =   12
      Top             =   2130
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Friends Name:"
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
      Left            =   225
      TabIndex        =   11
      Top             =   1170
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personal message:"
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
      Left            =   225
      TabIndex        =   10
      Top             =   2505
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Email:"
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
      Left            =   225
      TabIndex        =   9
      Top             =   1695
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Friends Email:"
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
      Left            =   225
      TabIndex        =   0
      Top             =   765
      Width           =   1215
   End
End
Attribute VB_Name = "frmref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtname1.Text = ""
    txtmailto.Text = ""
    txtmailfrom.Text = ""
    txtbody.Text = ""
End Sub

Private Sub Command1_Click()
    txtname1.Text = ""
    txtname2.Text = ""
    txtmailto.Text = ""
    txtmailfrom.Text = ""
    txtbody.Text = ""
    txtbody.Text = ""
    Unload frmref
End Sub

Private Sub Command2_Click()
Dim StrBuff As String, StrA As String

    If Len(Trim(txtmailto.Text)) <= 0 Then
        MsgBox "You must include the email of the recipient", vbCritical, frmref.Caption
        Exit Sub
    ElseIf IsEmail(txtmailto.Text) = False Then
        MsgBox "You have not inclucded a vaild email address for the recipient please try again", vbCritical, frmref.Caption
        Exit Sub
    ElseIf Len(Trim(txtname1.Text)) <= 0 Then
        MsgBox "You must enter the name of the person the mail is being sent to", vbCritical, frmref.Caption
        Exit Sub
    ElseIf Len(Trim(txtmailfrom.Text)) <= 0 Then
        MsgBox "You must include your email address", vbCritical, frmref.Caption
        Exit Sub
    ElseIf IsEmail(txtmailfrom.Text) = False Then
        MsgBox "Your email address does not seem to be valid please try again", vbCritical, frmref.Caption
        Exit Sub
    ElseIf Len(Trim(txtname2.Text)) <= 0 Then
        MsgBox "You must enter in your name", vbCritical, frmref.Caption
        Exit Sub
    Else
        StrBuff = StrConv(LoadResData(111, "CUSTOM"), vbUnicode)
        StrBuff = Replace(StrBuff, "#name2#", txtname1.Text)
        StrBuff = Replace(StrBuff, "#ProgName#", frmmain.Caption)
        StrBuff = Replace(StrBuff, "#name1#", txtname2.Text)
        StrBuff = Replace(StrBuff, "#URL#", TBookURL)
        StrBuff = Replace(StrBuff, "#message#", txtbody.Text)
        StrBuff = Replace(StrBuff, "#email1#", txtmailfrom.Text)
        StrBuff = Replace(StrBuff, "#IP#", smtp1.GetLocalIP)
        
        smtp1.SmtpServer = Config.mSMTP_serv
        smtp1.SmtpPort = 25
        smtp1.EmailTo = txtmailto.Text
        smtp1.EmailFrom = txtmailfrom.Text
        smtp1.EmailSubject = "Share Bookmark"
        smtp1.EmailMessage = StrBuff
        smtp1.MimeType = TextHTML
        smtp1.send
        
        If smtp1.MailSent Then
            MsgBox "The mail was successfully sent", vbInformation, frmref.Caption
            Unload frmref
        Else
            MsgBox "There was an error while sending the e-mail" _
            & DoubleCRLF & "This error may be due to:" & DoubleCRLF _
            & "The mail server your using may be expiring problems." _
            & vbNewLine & "You may incorrectly typed in mail server address." _
            & vbNewLine & "Your internet connection may be disconnected." _
            & DoubleCRLF & "Please try again latter.", vbCritical, frmref.Caption
        End If
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set frmref = Nothing
End Sub

Private Sub txtmailto_GotFocus()
    txtmailto.BackColor = Config.Hightlight
End Sub

Private Sub txtmailto_LostFocus()
    txtmailto.BackColor = vbWhite
End Sub

Private Sub txtmailfrom_GotFocus()
    txtmailfrom.BackColor = Config.Hightlight
End Sub

Private Sub txtmailfrom_LostFocus()
    txtmailfrom.BackColor = vbWhite
End Sub

Private Sub txtbody_GotFocus()
    txtbody.BackColor = Config.Hightlight
End Sub

Private Sub txtname1_LostFocus()
    txtname1.BackColor = vbWhite
End Sub

Private Sub txtname1_GotFocus()
    txtname1.BackColor = Config.Hightlight
End Sub

Private Sub txtname2_LostFocus()
    txtname2.BackColor = vbWhite
End Sub

Private Sub txtname2_GotFocus()
    txtname2.BackColor = Config.Hightlight
End Sub
Private Sub txtbody_LostFocus()
    txtbody.BackColor = vbWhite
End Sub

Private Sub Form_Load()
Dim mStr As String
    frmref.Icon = Nothing
    MakeFlatControls frmref
    mStr = "Hi I visited this cool web site " & TBookMark & vbNewLine & vbNewLine & "While I was there I found some very interesting" _
    & vbNewLine & "information That you may like." & vbNewLine
    
    txtbody.Text = mStr
    mStr = ""
End Sub

