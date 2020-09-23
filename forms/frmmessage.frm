VERSION 5.00
Begin VB.Form frmmessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cannot find server"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Try &Agian"
      Height          =   350
      Left            =   2250
      TabIndex        =   4
      Top             =   1590
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   350
      Left            =   3375
      TabIndex        =   3
      Top             =   1590
      Width           =   1020
   End
   Begin Project1.Line3D Line3D1 
      Height          =   105
      Left            =   -15
      TabIndex        =   2
      Top             =   645
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The page you currently requested can't be found on the server please close this dialog and try agian latter."
      Height          =   570
      Left            =   225
      TabIndex        =   1
      Top             =   840
      Width           =   4020
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unable to find page 404"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   735
      TabIndex        =   0
      Top             =   180
      Width           =   3240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmmessage.frx":0000
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Urlname = ""
    Unload frmmessage
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload frmmessage
    frmmain.webV.Navigate Urlname
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmessage = Nothing
    
End Sub
