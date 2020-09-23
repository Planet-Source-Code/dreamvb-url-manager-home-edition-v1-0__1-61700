VERSION 5.00
Begin VB.Form frmregnow 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register Now"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1095
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2295
      Width           =   825
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2070
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2295
      Width           =   825
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   135
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2295
      Width           =   825
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   3045
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2295
      Width           =   825
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3135
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "&Register"
      Height          =   350
      Left            =   2070
      TabIndex        =   3
      Top             =   3000
      Width           =   915
   End
   Begin VB.TextBox txtRegName 
      Height          =   300
      Left            =   1725
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1125
      Width           =   2520
   End
   Begin VB.TextBox txtRegCompany 
      Height          =   300
      Left            =   1725
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1590
      Width           =   2535
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
   Begin Project1.Line3D Line3D4 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   615
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Register your copy of URL Manager - Home Edition v1.0 for free."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   225
      TabIndex        =   13
      Top             =   105
      Width           =   3945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter in your Registration details."
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
      Left            =   630
      TabIndex        =   10
      Top             =   750
      Width           =   3300
   End
   Begin VB.Image imgstate 
      Height          =   405
      Left            =   3975
      Top             =   2250
      Width           =   405
   End
   Begin VB.Image imglock 
      Height          =   480
      Left            =   465
      Picture         =   "frmregnow.frx":0000
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgunlock 
      Height          =   480
      Left            =   420
      Picture         =   "frmregnow.frx":0442
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
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
      Left            =   120
      TabIndex        =   6
      Top             =   2010
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
      Left            =   120
      TabIndex        =   5
      Top             =   1650
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
      Left            =   120
      TabIndex        =   0
      Top             =   1155
      Width           =   1530
   End
End
Attribute VB_Name = "frmregnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TLock As Boolean

Private Sub cmdapply_Click()
Dim mCode1 As String, mCode2 As String
Dim I As Long

    txtRegName.Text = Trim(txtRegName.Text)
    txtRegCompany.Text = Trim(txtRegCompany.Text)
    TLock = True
    
    If Trim(Len(txtRegName.Text)) < 4 Then
        MsgBox "Registered name must be at least 4 or more characters long.", vbInformation, frmregnow.Caption
        Exit Sub
    ElseIf Len(Trim(txtRegCompany.Text)) < 4 Then
        MsgBox "Registered company name must be at least 4 or more characters long.", vbInformation, frmregnow.Caption
        Exit Sub
    End If
    
    
    For I = 0 To txtkey.Count - 1
        If Len(Trim(txtkey(I).Text)) < 5 Then
            Exit For
            TLock = True
        End If
    Next
    I = 0
    
    
    mCode1 = txtkey(0).Text & txtkey(1).Text & txtkey(2).Text & txtkey(3).Text
    
    If Not Len(mCode1) = 20 Then
        TLock = True
    End If
    
    If Check(mCode1, 450, 9) Then
        imgstate.Picture = imgunlock.Picture
        TLock = False
    Else
        imgstate.Picture = imglock.Picture
        TLock = True
    End If
    
    If Not TLock Then
        SaveSetting "DMUrlMan", "Register", "RegName", txtRegName.Text
        SaveSetting "DMUrlMan", "Register", "Company", txtRegCompany.Text
        SaveSetting "DMUrlMan", "Register", "Key", mCode1
        MsgBox "Thank you for registering." & vbCrLf & vbCrLf _
        & "You may need to restart the program for the new changes to take effect.", vbInformation, frmregnow.Caption
        Unload frmregnow
        Exit Sub
    Else
        WriteDefault
        MsgBox "Invalid registration key entered. Please check that you entered the key in correctly.", vbInformation, frmregnow.Caption
    End If
    
    mCode = ""
    mCode2 = ""
    
End Sub

Private Sub cmdcancel_Click()
    Unload frmregnow
End Sub

Private Sub cmdcopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtID, vbCFText
    MsgBox "Your Activation key has now copied to the clipboard.", vbInformation, frmregnow.Caption
    
End Sub

Private Sub Form_Load()
    MakeFlatControls frmregnow
    imgstate.Picture = imglock.Picture
    imgstate.Picture = imglock.Picture
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmregnow = Nothing
End Sub



Private Sub txtID_GotFocus()
    txtID.BackColor = RGB(217, 236, 255)
End Sub

Private Sub txtID_LostFocus()
    txtID.BackColor = vbWhite
End Sub

Private Sub txtkey_Change(Index As Integer)
    txtkey(Index).Text = Replace(txtkey(Index).Text, " ", "")
    txtkey(Index).Text = Trim(UCase(txtkey(Index).Text))
    txtkey(Index).SelStart = Len(txtkey(Index).Text)
    If Index = 3 And Len(txtkey(3)) = 5 Then
        cmdapply.SetFocus
        Exit Sub
    End If
    
    If Len(txtkey(Index).Text) >= 5 Then txtkey(Index + 1).SetFocus
    
End Sub

Private Sub txtkey_GotFocus(Index As Integer)
    txtkey(Index).BackColor = RGB(217, 236, 255)
End Sub

Private Sub txtkey_LostFocus(Index As Integer)
    txtkey(Index).BackColor = vbWhite
End Sub

Private Sub txtRegCompany_GotFocus()
    txtRegCompany.BackColor = RGB(217, 236, 255)
End Sub

Private Sub txtRegCompany_LostFocus()
    txtRegCompany.BackColor = vbWhite
End Sub

Private Sub txtRegName_GotFocus()
    txtRegName.BackColor = RGB(217, 236, 255)
End Sub

Private Sub txtRegName_LostFocus()
    txtRegName.BackColor = vbWhite
End Sub
