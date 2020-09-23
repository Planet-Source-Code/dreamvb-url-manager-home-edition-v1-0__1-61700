VERSION 5.00
Begin VB.Form frmregdll 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Install Add-ins"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Install"
      Height          =   350
      Left            =   3045
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   3045
      TabIndex        =   4
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdunreg 
      Caption         =   "&Unregister"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3045
      TabIndex        =   3
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "&Register"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3045
      TabIndex        =   2
      Top             =   1005
      Width           =   1215
   End
   Begin VB.ListBox lstaddins 
      Height          =   2010
      Left            =   105
      TabIndex        =   0
      Top             =   465
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Add-ins"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   1545
   End
End
Attribute VB_Name = "frmregdll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lstInx As Integer

Private Sub cmdadd_Click()
    frminstdll.Show vbModal, frmregdll
End Sub

Private Sub cmdexit_Click()
    lstInx = 0
    lstaddins.Clear ' Clear the listbox
    Unload frmregdll    ' Unload the form
End Sub

Private Sub cmdreg_Click()
Dim mPlg As String

    mPlg = FixPath(App.Path) & "add-ins\" & mAddins(lstInx).mPlugDLL
    If Not RegisterActiveX(mPlg, Register) Then
        MsgBox "There was an unexpected error while registering the add-in.", vbExclamation, "Unexpected error"
        Exit Sub
    Else
        MsgBox mAddins(lstInx).mPlugName & " has now been successfully registered.", vbInformation, "Finished"
    End If
    mPlg = ""
    
End Sub

Private Sub cmdunreg_Click()
Dim mPlg As String

    mPlg = FixPath(App.Path) & "add-ins\" & mAddins(lstInx).mPlugDLL
    If Not RegisterActiveX(mPlg, UnRegister) Then
        MsgBox "There was an unexpected error while unregistering the add-in.", vbExclamation, "Unexpected error"
        Exit Sub
    Else
        MsgBox mAddins(lstInx).mPlugName & " has now been successfully uregistered.", vbInformation, "Finished"
    End If
    
    mPlg = ""

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim I As Long
    For I = LBound(mAddins) To UBound(mAddins)
       If I > 0 Then lstaddins.AddItem mAddins(I).mPlugName
    Next
    I = 0
    MakeFlatControls frmregdll
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmregdll = Nothing ' Release the form from memory
End Sub

Private Sub lstaddins_Click()
    lstInx = lstaddins.ListIndex + 1
    cmdreg.Enabled = True
    cmdunreg.Enabled = True
    
End Sub
