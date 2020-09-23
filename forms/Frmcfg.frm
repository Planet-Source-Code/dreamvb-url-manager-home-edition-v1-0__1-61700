VERSION 5.00
Begin VB.Form Frmcfg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton reglater 
      Caption         =   "Register later"
      Height          =   350
      Left            =   3390
      TabIndex        =   3
      Top             =   915
      Width           =   1215
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "Register Now"
      Height          =   350
      Left            =   1965
      TabIndex        =   2
      Top             =   915
      Width           =   1215
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   180
      TabIndex        =   1
      Top             =   645
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   53
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "URL Manager - Home Edtion is currently unregistered would you like to register now it is free."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   4485
   End
End
Attribute VB_Name = "Frmcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreg_Click()
    frmregnow.Show vbmoal, frmmain
    Unload Frmcfg
End Sub

Private Sub Form_Load()
    Frmcfg.Caption = frmmain.Caption
    Set Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Frmcfg = Nothing
End Sub

Private Sub reglater_Click()
    Unload Frmcfg
End Sub
