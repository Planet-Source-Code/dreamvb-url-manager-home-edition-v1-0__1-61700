VERSION 5.00
Begin VB.Form frmfindcat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Category"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   2760
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   350
      Left            =   1395
      TabIndex        =   2
      Top             =   825
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   1
      Top             =   255
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find What:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   915
   End
End
Attribute VB_Name = "frmfindcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    txtname.Text = ""
    Unload frmfindcat
End Sub

Private Sub cmdok_Click()
    If Not frmmain.FindCat(frmmain.tv1, txtname.Text) Then
        MsgBox "The category " & Chr(34) & txtname.Text & Chr(34) & " was not be found.", vbInformation, frmfindcat.Caption
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    frmfindcat.Icon = Nothing    ' Remove the forms icon
    FlatBorder txtname.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmrename = Nothing
End Sub

Private Sub txtname_GotFocus()
    txtname.BackColor = Config.Hightlight
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdok_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtname_LostFocus()
    txtname.BackColor = vbWhite
End Sub
