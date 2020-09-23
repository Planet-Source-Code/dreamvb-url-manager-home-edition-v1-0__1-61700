VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmstatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bookmark Status"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   165
      TabIndex        =   12
      Top             =   2040
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3240
      TabIndex        =   10
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4245
      Top             =   195
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get &Status"
      Height          =   350
      Left            =   1725
      TabIndex        =   0
      Top             =   2205
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4110
      Top             =   645
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   53
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmstatus.frx":0000
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   945
      TabIndex        =   11
      Top             =   180
      Width           =   60
   End
   Begin VB.Label lblhost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1095
      TabIndex        =   9
      Top             =   1620
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host:"
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
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   1620
      Width           =   450
   End
   Begin VB.Label lblsize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1095
      TabIndex        =   7
      Top             =   1305
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Length:"
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
      Left            =   225
      TabIndex        =   6
      Top             =   1305
      Width           =   645
   End
   Begin VB.Label lblstattext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1485
      TabIndex        =   5
      Top             =   1005
      Width           =   45
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1095
      TabIndex        =   4
      Top             =   1005
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   210
      TabIndex        =   3
      Top             =   1005
      Width           =   540
   End
   Begin VB.Label lbllocation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1095
      TabIndex        =   2
      Top             =   690
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   210
      TabIndex        =   1
      Top             =   690
      Width           =   765
   End
End
Attribute VB_Name = "frmstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PhaseHTTP(sData As String)
Dim iPart, lpart As Long, StatusCode As Integer, StatusText As String, WebURL As String, ContentLen As Long
On Error Resume Next

    iPart = InStr(1, sData, " ", vbTextCompare)
    lpart = InStr(iPart + 1, sData, " ", vbTextCompare)
    
    WebURL = Inet1.URL
    StatusCode = Val(Mid(sData, iPart + 1, lpart - iPart - 1)) ' Extract HTTP status code
    iPart = 0: lpart = 0 ' Reset pointer
    
    iPart = InStr(1, sData, "Content-Length: ", vbTextCompare)
    lpart = InStr(iPart + 1, sData, vbCrLf, vbTextCompare)
    
    ContentLen = Val(Mid(sData, iPart + 16, lpart - iPart - 16)) ' extract HTTP content length value
    iPart = 0: lpart = 0 ' Reset pointer
    
    Select Case StatusCode
        Case 100
            StatusText = "Continue"
        Case 101
            StatusText = "Switching Protocols"
        Case 200
            StatusText = "OK"
        Case 201
            StatusText = "Created"
        Case 202
            StatusText = "Accepted"
        Case 203
            StatusText = "Non-Authoritative Information"
        Case 301
            StatusText = "Moved Permanently"
        Case 302
            StatusText = "Found"
        Case 303
            StatusText = "See other"
        Case 305
            StatusText = "Use Proxy"
        Case 400
            StatusText = "Bad Request"
        Case 401
            StatusText = "Unauthorized"
        Case 403
            StatusText = "Forbidden"
        Case 404
            StatusText = "Not Found"
        Case 408
            StatusText = "Request Timeout"
        Case 410
            StatusText = "Gone"
        Case 414, 416
            StatusText = "Request URI Too Long"
        Case 500, 501, 502, 503, 504
            StatusText = "Internal Server Error"
        Case 505
             StatusText = "HTTP Version Not Supported"
        Case Else
            StatusText = ""
    End Select
    
    ' Update Labels with http header info
    lbllocation.Caption = WebURL
    lblstatus.Caption = StatusCode
    lblstattext.Caption = StatusText
    lblsize.Caption = ContentLen & " bytes"
    lblhost.Caption = Inet1.RemoteHost
    Timer1.Enabled = False
    
    ' Clear up vars
    StatusCode = 0
    ContentLen = 0
    WebURL = ""
    StatusText = ""
    Inet1.Cancel
End Sub
Private Sub cmdcancel_Click()
    Unload frmstatus
End Sub

Private Sub cmdget_Click()
On Error Resume Next
    Timer1.Enabled = True   ' Enable the timer control
    Inet1.Cancel
    Inet1.OpenURL TBookURL, icString ' open the bookmark
    
End Sub

Private Sub Form_Load()
    lbltitle.Caption = "Status for: " & "(" & TBookMark & ")"
    
End Sub

Private Sub Form_Resize()
    Line3D2.Width = frmstatus.ScaleWidth - Line3D2.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmstatus = Nothing
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    If State = icConnecting Then Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
Dim StrBuff As String, A As String
'On Error Resume Next

    If Not Inet1.StillExecuting Then
        A = Inet1.GetHeader
  
        If Len(StrBuff) = 0 Then ' Check buffer length
            StrBuff = StrBuff & A ' Get http header
            Timer1.Enabled = False ' Disbale the timer control
            Inet1.Cancel ' Close connection
            DoEvents
        End If
    End If
    PhaseHTTP StrBuff ' Phase the http headers
    StrBuff = "" ' Clear buffer
End Sub
