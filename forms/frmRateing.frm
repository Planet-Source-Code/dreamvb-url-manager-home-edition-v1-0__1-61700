VERSION 5.00
Begin VB.Form frmRateing 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rateing System"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   3525
      TabIndex        =   0
      Top             =   2550
      Width           =   1065
   End
   Begin VB.Label lblrate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   870
      TabIndex        =   7
      Top             =   165
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rated"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   165
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   3
      X1              =   135
      X2              =   4485
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   2
      X1              =   120
      X2              =   4485
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   1
      X1              =   135
      X2              =   4485
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   135
      X2              =   4485
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Very basic website."
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
      Index           =   4
      Left            =   1650
      TabIndex        =   5
      Top             =   2115
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic web site some good things."
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
      Index           =   3
      Left            =   1650
      TabIndex        =   4
      Top             =   1740
      Width           =   2850
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Good web site."
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
      Left            =   1650
      TabIndex        =   3
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excellent web site."
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
      Left            =   1650
      TabIndex        =   2
      Top             =   1020
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Must see web site."
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
      Left            =   1650
      TabIndex        =   1
      Top             =   615
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   135
      Picture         =   "frmRateing.frx":0000
      Top             =   2085
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   135
      Picture         =   "frmRateing.frx":017E
      Top             =   1710
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   135
      Picture         =   "frmRateing.frx":0354
      Top             =   1350
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   135
      Picture         =   "frmRateing.frx":0582
      Top             =   990
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   135
      Picture         =   "frmRateing.frx":07F5
      Top             =   585
      Width           =   1260
   End
End
Attribute VB_Name = "frmRateing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
    Unload frmRateing ' Unload the form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRateing = Nothing
End Sub
