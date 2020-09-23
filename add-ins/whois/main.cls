VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Bookmarks name title
Public lzBookMark As String
' Bookmarks URL address
Public lzBookmarkURL As String
' Whois lookup server form config file
Public WhoisServ As String
' Your SMTP server that in the config file
Public lzSmtpServ As String
' Long hangle of the programs window
Public lzHwnd As Long

Public Sub RunPlug()
    frmwhois.txtfind.Text = frmwhois.PhaseDomain(lzBookmarkURL)
    frmwhois.txtserver.Text = WhoisServ
    frmwhois.Show vbModal
    
End Sub

