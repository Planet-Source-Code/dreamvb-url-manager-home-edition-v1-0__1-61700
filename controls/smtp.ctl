VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl smtp 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   Picture         =   "smtp.ctx":0000
   ScaleHeight     =   660
   ScaleWidth      =   600
   Begin MSWinsockLib.Winsock wsksmtp 
      Left            =   495
      Top             =   795
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "smtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DM URL Bookmarker SendMail Script
' writen by Ben Jones
' Date: 10-7-2002

Enum T_MimeType
    TextPlain = 0
    TextHTML = 1
End Enum

Private Type sMail
    mMailTo As String
    mMailFrom As String
    mSubject As String
    mMessageBody As String
End Type

Private Mail_Info As sMail          ' Type to keep all our email information
Private Response As String          ' current smtp server response code

Public SmtpServer As String         ' current smtp server
Public SmtpPort As Integer          ' current smtp server port to use
Public ErrorCode As Integer         ' Returns last error code if any
Public MimeType As T_MimeType       ' Type to store email mime format eg text or html
Public MailSent As Boolean          ' returns if mail was sent returns true on Success
' error code constents
Private Const TimedOut = 0          ' connection timmed out
Private Const UnknownCode = 1       ' Unknown Response code
Private Const InvaildEmail = 2      ' Invaild email address found
Private Const ServConnectErr = 3    ' Error connecting to server

Private MailMimeType As String      ' holds our mime type
Public Function GetLocalIP() As String
    GetLocalIP = wsksmtp.LocalIP
End Function
Private Function ConnectToSmtpServ() As Boolean
    wsksmtp.Close                   ' close winsock first
    wsksmtp.RemoteHost = SmtpServer ' smtp server address
    wsksmtp.RemotePort = SmtpPort   ' smtp server port
    wsksmtp.Connect                 ' connect to the smtp server
    ' this will wait for a connection
    Do While wsksmtp.State <> sckConnected
        DoEvents
        If wsksmtp.State = sckClosed Or _
            wsksmtp.State = sckClosing Or _
            wsksmtp.State = sckError Then
            ConnectToSmtpServ = False
            Exit Function
        End If
    Loop
    ConnectToSmtpServ = True
    
End Function
Private Function GetEmailName(sEmail As String) As String
' This extracts the name in the email up the @ sign
' example GetEmailName "demo@mymail.com" returns demo
Dim iPart As Long
    iPart = InStr(sEmail, "@")
    If iPart > 0 Then GetEmailName = Mid(sEmail, 1, iPart - 1) Else GetEmailName = ""
    iPart = 0
End Function
Public Property Let EmailTo(ByVal strEmailTo As String)
    Mail_Info.mMailTo = strEmailTo  ' holds the email to address
End Property
Public Property Let EmailFrom(ByVal strEmailFrom As String)
    Mail_Info.mMailFrom = strEmailFrom  ' holds the email from address
End Property
Public Property Let EmailSubject(ByVal strEmailSubject As String)
    Mail_Info.mSubject = strEmailSubject    ' holds the email subject
End Property
Public Property Let EmailMessage(ByVal strEmailMessage As String)
    Mail_Info.mMessageBody = strEmailMessage    ' holds the email message text
End Property

Private Sub Replay(strdata As String)
    ' used to send smtp commands to smtp server
    If wsksmtp.State = sckConnected Then ' checks if connected
        wsksmtp.SendData strdata & vbCrLf
        DoEvents
    End If
    
End Sub
Private Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            ErrorCode = TimedOut
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            ErrorCode = UnknownCode
            Exit Sub
        End If
    Wend
    Response = "" ' clean out response code
End Sub
Sub send()
Dim MailHeader As String
Dim EmailName As String
On Error Resume Next
    ' checks to see if we are connected to the smtp sever
    If ConnectToSmtpServ = False Then ErrorCode = ServConnectErr: Exit Sub
    ' This part check for a vaild email address and extract the email name
    If Len(GetEmailName(Mail_Info.mMailFrom)) <= 0 Then ErrorCode = InvaildEmail: Exit Sub
    EmailName = GetEmailName(Mail_Info.mMailFrom) ' extracts the name from the email address

    ' Find out what the mime type of the text is to be
    If MimeType = TextHTML Then
        MailMimeType = "text/html"
    Else
        MailMimeType = "text/plain"
    End If
        ' Email header to send with email message
        MailHeader = "From: """ & EmailName & """" & " <" & Mail_Info.mMailFrom & ">" & vbCrLf
        MailHeader = MailHeader & "To: " & "<" & Mail_Info.mMailTo & ">" & vbCrLf
        MailHeader = MailHeader & "Subject: " & Mail_Info.mSubject & vbCrLf
        MailHeader = MailHeader & "Date: " & Format(Now, "ddd,dd mmm yyyy hh:mm:ss +0100") & vbCrLf
        MailHeader = MailHeader & "MIME-Version: 1.0" & vbCrLf
        MailHeader = MailHeader & "Content-Type: " & MailMimeType & ";" & vbCrLf
        MailHeader = MailHeader & "Contect-Transfer-Encoding: 7bit" & vbCrLf
        MailHeader = MailHeader & "X-Priority: 3" & vbCrLf
        MailHeader = MailHeader & "X-MSMail-Priority: Normal" & vbCrLf
        MailHeader = MailHeader & "X-Mailer: SendMail client" & vbCrLf
        WaitFor ("220")
        Replay "HELO " & wsksmtp.LocalIP            ' sends your local ip
        WaitFor ("250")
        Replay "MAIL FROM: " & Mail_Info.mMailFrom  ' sends the mail from
        WaitFor ("250")
        Replay "RCPT TO: " & Mail_Info.mMailTo      ' sends the mail to
        WaitFor ("250")
        Replay "DATA "
        WaitFor ("354")
        Replay MailHeader & Mail_Info.mMessageBody & vbCrLf & "." ' sends email headers and message data
        WaitFor ("250")
        Replay "QUIT "                              ' sends  quit command to smtp server
        WaitFor ("221")
        wsksmtp.Close                               ' close winsock connection
        
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Size 600, 660
    
End Sub

Private Sub wsksmtp_Connect()
    MailSent = True ' is connected to smtp server other wise returns false
    
End Sub

Private Sub wsksmtp_DataArrival(ByVal bytesTotal As Long)
    wsksmtp.GetData Response ' gets the incomming data
    
End Sub

Private Sub wsksmtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsksmtp.Close           ' close the winsock connection
    ErrorCode = UnknownCode ' sends error code
    
End Sub
