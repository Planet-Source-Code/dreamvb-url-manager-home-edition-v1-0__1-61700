VERSION 5.00
Begin VB.Form frmmoveto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Bookmark"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Flat Flat3 
      Height          =   240
      Left            =   4215
      TabIndex        =   6
      Top             =   840
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   423
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin Project1.Flat Flat1 
      Height          =   270
      Left            =   1140
      TabIndex        =   5
      Top             =   825
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   476
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin VB.ComboBox cboMoveto 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Top             =   810
      Width           =   3360
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3300
      TabIndex        =   3
      Top             =   1290
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1935
      TabIndex        =   2
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the category were you like move this bookmark to the list below:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Move to:"
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
      Left            =   150
      TabIndex        =   0
      Top             =   855
      Width           =   750
   End
End
Attribute VB_Name = "frmmoveto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MoveTo As String

Private Sub cboMoveto_Click()
    MoveTo = cboMoveto.Text
    If UCase(RecoredName) = UCase(MoveTo) Then
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
    End If
    
End Sub

Private Sub cboMoveto_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cboMoveto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdcancel_Click()
    MoveTo = ""
    cboMoveto.Clear
    Unload frmmoveto
End Sub

Private Sub cmdok_Click()
Dim lzReturn As Long, ans As Integer

    lzReturn = TMoveToUrl(RecoredName, SiteID, MoveTo)
    
    If lzReturn > 0 Then
        ans = MsgBox("The bookmark has been now successfully moved to " & cboMoveto.Text _
        & DoubleCRLF & "Do you want to delete the old bookmark now.", vbYesNo Or vbQuestion, frmmoveto.Caption)
        
        If ans = vbNo Then
            cmdcancel_Click
        Else
            DeleteURL SiteID, RecoredName
            RemoveLVItem LstIndex, frmmain.lstsites
            SiteID = 0
            LstIndex = 0
            LoadSites RecoredName   ' Reload all the bookmarks back in
            MsgBox "The bookmark has now been successfully deleted.", vbInformation, frmmoveto.Caption
        End If
    Else
        MsgBox "There was an error while moving the selected bookmark.", vbCritical, frmmoveto.Caption
    End If
    cmdcancel_Click
    
End Sub

Private Sub Form_Load()
    frmmoveto.Icon = Nothing    ' Remove the forms icon
    Initcbocat cboMoveto
    cboMoveto.RemoveItem 0
    cboMoveto.ListIndex = 0
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmrename = Nothing
End Sub

