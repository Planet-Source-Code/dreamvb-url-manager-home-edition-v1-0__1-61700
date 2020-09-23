VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmwhois 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "whois Lookup - Add-on"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   6225
      TabIndex        =   9
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Whois"
      Height          =   350
      Left            =   4830
      TabIndex        =   8
      Top             =   4260
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock DmWhois 
      Left            =   1650
      Top             =   2835
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtresult 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3000
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1095
      Width           =   7350
   End
   Begin VB.TextBox txtfind 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   3
      Top             =   630
      Width           =   6000
   End
   Begin VB.TextBox txtport 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4290
      TabIndex        =   2
      Text            =   "43"
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox txtserver 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   1
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHOIS Look up"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5115
      TabIndex        =   7
      Top             =   120
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domain Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3765
      TabIndex        =   5
      Top             =   165
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Whois Server:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   1230
   End
End
Attribute VB_Name = "frmwhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function PhaseDomain(lzUrl As String) As String
Dim Ipart As Long, lPart As Long, StrA As String
On Error Resume Next

    Ipart = InStr(lzUrl, "http://")
    If Ipart = 1 Then lzUrl = Mid(lzUrl, Ipart + 7, Len(lzUrl) - Ipart)
    Ipart = InStr(lzUrl, ".")
    If Ipart = 0 Then PhaseDomain = lzUrl: Exit Function
    StrA = Mid$(lzUrl, Ipart + 1, Len(lzUrl) - Ipart)
    lPart = InStr(StrA, "/")
    If lPart = 0 Then PhaseDomain = StrA: Exit Function
    StrA = Mid$(StrA, 1, lPart - 1)
    PhaseDomain = StrA
    Ipart = 0: lPart = 0: StrA = ""
    
End Function

Private Sub cmdcancel_Click()
    Unload frmwhois
End Sub

Private Sub cmdfind_Click()
    If Val(txtport.Text) <= 0 Then
        MsgBox "The server port must not be set to zero please try agian", vbInformation, frmwhois.Caption
        txtport.SetFocus
        Exit Sub
   ElseIf Len(Trim(txtserver.Text)) <= 0 Then
        MsgBox "You must include the WHOIS server's name", vbInformation, frmwhois.Caption
        txtserver.SetFocus
        Exit Sub
   ElseIf Len(Trim(txtfind.Text)) <= 0 Then
        MsgBox "There was no serach string found please try again", vbInformation, frmwhois.Caption
        txtfind.SetFocus
        Exit Sub
    Else
        frmwhois.MousePointer = vbHourglass
        DmWhois.Close
        DmWhois.LocalPort = 0
        DmWhois.Connect Trim(txtserver.Text), Val(txtport.Text)
    End If

End Sub

Private Sub DmWhois_Connect()
    If DmWhois.State <> sckOpen Then DmWhois.SendData txtfind.Text & vbCrLf
    frmwhois.MousePointer = vbDefault
    
End Sub

Private Sub DmWhois_DataArrival(ByVal bytesTotal As Long)
    DmWhois.GetData sData, vbString, bytesTotal
    sData = Replace(sData, vbLf, vbCrLf)
    txtresult.Text = Trim(sData)
    
End Sub

Private Sub DmWhois_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If DmWhois.State = sckError Then MsgBox "There was an error while connecting to the address that you Specified" _
    & vbNewLine & "This may be due to:" & DoubleCRLF _
    & "There may be a Problem with the server" & vbNewLine _
    & "You may have entered an incorrect address" _
    & vbNewLine & "The Domain name may not exist" _
    & DoubleCRLF & "Please try again latter.", vbCritical, frmwhois.Caption

    frmwhois.MousePointer = vbDefault
    DmWhois.Close
    
End Sub

Private Sub Form_Load()
    frmwhois.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmwhois = Nothing
    
End Sub

Private Sub txtport_Change()
    If IsNumeric(txtport.Text) = False Then txtport.Text = ""
    
End Sub


