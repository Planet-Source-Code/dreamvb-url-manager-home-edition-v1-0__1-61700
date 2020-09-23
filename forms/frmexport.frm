VERSION 5.00
Begin VB.Form frmexport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Category"
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
      Caption         =   "New name:"
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
      Width           =   975
   End
End
Attribute VB_Name = "frmexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    Unload frmrename
End Sub

Private Sub cmdok_Click()
Dim iCheck As Long

    If Len(Trim(txtname.Text)) = 0 Then
        MsgBox "You must include a name for the category.", vbExclamation, frmrename.Caption
        Exit Sub
    ElseIf Len(Trim(txtname.Text)) > 30 Then
        MsgBox "Your category name may only be between 1 and 30 in length.", vbExclamation, frmrename.Caption
        Exit Sub
    ElseIf HasSpace(txtname.Text) = True Then
        MsgBox "The category name may not contain spaces within the string.", vbExclamation, frmrename.Caption
        Exit Sub
    Else
        iCheck = RenameTable(RecoredName, txtname.Text)
        If Not iCheck = 1 Then
            MsgBox "there was an error while renameing the category name" _
            & DoubleCRLF & " Please try agian."
            Exit Sub
        Else
            MsgBox "The category name has been successfully chnaged.", vbInformation, frmrename.Caption
        End If
    End If
    iCheck = 0
    RecoredName = ""
    TvIndex = 0
    LstIndex = 0
    frmmain.InitTv
    Unload frmrename
End Sub

Private Sub Form_Load()
    frmrename.Icon = Nothing    ' Remove the forms icon
    txtname.Text = RecoredName  ' Assign the recored name to the textbox
    txtname.SelStart = 0
    txtname.SelLength = Len(txtname.Text)
    FlatBorder txtname.hWnd, True
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
