VERSION 5.00
Begin VB.Form frmchurl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Check bookmark status"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Line3D Line3D1 
      Height          =   105
      Left            =   0
      TabIndex        =   4
      Top             =   810
      Width           =   4500
      _extentx        =   7938
      _extenty        =   185
   End
   Begin VB.CommandButton cmdchk 
      Caption         =   "Check &Status"
      Height          =   350
      Left            =   2145
      TabIndex        =   0
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   350
      Left            =   3510
      TabIndex        =   2
      Top             =   2325
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kbs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2130
      TabIndex        =   12
      Top             =   1815
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kbs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2130
      TabIndex        =   11
      Top             =   1485
      Width           =   255
   End
   Begin VB.Label kbs2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   210
      Left            =   1485
      TabIndex        =   10
      Top             =   1815
      Width           =   90
   End
   Begin VB.Label kbs1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   1485
      TabIndex        =   9
      Top             =   1485
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data sent:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1815
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data received:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1485
      Width           =   1050
   End
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1485
      TabIndex        =   6
      Top             =   1140
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   1110
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmchurl.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblurlname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#101"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   825
      TabIndex        =   3
      Top             =   495
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait Checking URL Status for:"
      Height          =   195
      Left            =   795
      TabIndex        =   1
      Top             =   210
      Width           =   2670
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   810
      Left            =   0
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmchurl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdchk_Click()
    Dim Ret As QOCINFO
    Ret.dwSize = Len(Ret)
    
    If IsDestinationReachable(lblurlname.Caption, Ret) = 0 Then
        lblstat.Caption = "The destination could not be reached!"
    Else
        lblstat.Caption = "The bookmark status was reached ok"
        kbs1.Caption = Format$(Ret.dwInSpeed / 1024, "#.0")
        kbs2.Caption = Format$(Ret.dwOutSpeed / 1024, "#.0")
    End If
    
End Sub

Private Sub cmdclose_Click()
    Unload frmchurl
End Sub

Private Sub Form_Load()
   lblurlname.Caption = Urlname
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmchurl = Nothing
End Sub

