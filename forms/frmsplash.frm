VERSION 5.00
Begin VB.Form frmsplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1185
      Top             =   1320
   End
   Begin VB.Image Image3 
      Height          =   1290
      Left            =   975
      Picture         =   "frmsplash.frx":0000
      Top             =   2175
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Picture         =   "frmsplash.frx":0C85
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      Picture         =   "frmsplash.frx":1B94
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3465
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Long


Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "The program is already running please shut down the other instance of the program", vbInformation, frmmain.Caption
        Unload frmmain
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmsplash = Nothing
    
End Sub

Private Sub Timer1_Timer()
    I = I + 1
    If I >= 2 Then
        Unload frmsplash
        frmmain.Show
        I = 0
    End If
    
End Sub
