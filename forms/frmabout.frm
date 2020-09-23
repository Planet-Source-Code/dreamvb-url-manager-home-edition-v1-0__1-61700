VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   5790
      TabIndex        =   16
      Top             =   0
      Width           =   5790
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You're favorite bookmark storage helper."
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
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   2955
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL Manager - Home Edition v1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2475
         TabIndex        =   17
         Top             =   90
         Width           =   3180
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "frmabout.frx":0000
         Top             =   15
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4485
      TabIndex        =   14
      Top             =   3030
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   165
      ScaleHeight     =   1065
      ScaleWidth      =   5400
      TabIndex        =   3
      Top             =   1530
      Width           =   5460
      Begin VB.Label lblregserial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000-0000-0000-0000"
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
         Left            =   1785
         TabIndex        =   12
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label lblregcompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered Company"
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
         Left            =   1785
         TabIndex        =   11
         Top             =   420
         Width           =   1980
      End
      Begin VB.Label lblregname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered Name"
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
         Left            =   1785
         TabIndex        =   10
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial number:"
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
         Left            =   90
         TabIndex        =   9
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company name:"
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
         Left            =   90
         TabIndex        =   8
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered name:"
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
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Website"
      Height          =   360
      Left            =   6870
      TabIndex        =   0
      Top             =   2415
      Width           =   1230
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   53
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   360
      Left            =   8220
      TabIndex        =   1
      Top             =   2415
      Width           =   870
   End
   Begin Project1.Line3D Line3D2 
      Height          =   105
      Left            =   60
      TabIndex        =   13
      Top             =   2850
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   53
   End
   Begin Project1.Line3D Line3D3 
      Height          =   105
      Left            =   60
      TabIndex        =   15
      Top             =   3540
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   53
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is freeware"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   6
      Top             =   930
      Width           =   2130
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This software is registered to:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1245
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002-2003 eRay Studios "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2460
      TabIndex        =   4
      Top             =   3750
      Width           =   3285
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmabout
End Sub

Private Sub Form_Load()
Dim mCode1 As String, mCode2 As String

    If CheckRegUser Then
        lblregname.Caption = Config.ProgRegister.mRegName
        lblregcompany.Caption = Config.ProgRegister.mRegCompany
        lblregserial.Caption = Config.ProgRegister.mRegKey
    End If
    
End Sub

Private Sub Form_Resize()
    Line3D1.Width = (frmabout.ScaleWidth)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub

