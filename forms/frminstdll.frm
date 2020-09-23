VERSION 5.00
Begin VB.Form frminstdll 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Install Add-ins"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   1590
      TabIndex        =   9
      Top             =   2715
      Width           =   1215
   End
   Begin VB.CommandButton cmdinstall 
      Caption         =   "&Install"
      Height          =   350
      Left            =   150
      TabIndex        =   8
      Top             =   2715
      Width           =   1215
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   2475
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   330
      Left            =   3165
      TabIndex        =   6
      Top             =   1860
      Width           =   420
   End
   Begin VB.TextBox txtdllname 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1860
      Width           =   2985
   End
   Begin VB.TextBox txtclassname 
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
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1095
      Width           =   2985
   End
   Begin VB.TextBox txtmnucaption 
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
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Top             =   405
      Width           =   2985
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add-in Dll filename"
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
      Left            =   90
      TabIndex        =   2
      Top             =   1515
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add-in class name eg myplug.main"
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
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add-in menu caption"
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
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1770
   End
End
Attribute VB_Name = "frminstdll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    txtmnucaption.Text = ""
    txtclassname.Text = ""
    txtdllname.Text = ""
    Unload frminstdll
End Sub
Private Sub Writetoini(plgIni As String)
Dim aFile As Long
Dim sPlgInfo As String
    aFile = FreeFile
    sPlgInfo = "plgname=" & txtmnucaption.Text & ";" & txtclassname.Text & ";" & txtdllname.Text
    Open plgIni For Append As #aFile
        Print #aFile, sPlgInfo
    Close #aFile
    sPlgInfo = ""
    
End Sub
Private Sub cmdinstall_Click()
Dim mPlgName As String

    If Len(Trim(txtmnucaption.Text)) = 0 Then
        MsgBox "You must include a name for the add-in", vbInformation, frminstdll.Caption
        Exit Sub
    ElseIf Len(Trim(txtclassname.Text)) = 0 Then
        MsgBox "You must include the class name of your add-in eg myplug.main", vbInformation, frminstdll.Caption
        Exit Sub
    ElseIf Len(Trim(txtdllname.Text)) = 0 Then
        MsgBox "You need to select the filename for your add-in", vbInformation, frminstdll.Caption
        Exit Sub
    Else
        mPlgName = FixPath(App.Path) & "add-ins\" & txtdllname.Text
        If Not RegisterActiveX(mPlgName, Register) Then
            MsgBox "There was an unexpected error while registering the add-in.", vbExclamation, "Unexpected error"
            mPlgName = ""
            Exit Sub
        Else
            Writetoini plgIni
            frmmain.LoadAddinIni plgIni
            MsgBox "The add-in has now been successfully registered.", vbInformation, "Finished"
        End If
    End If
    
    mPlgName = ""
    cmdcancel_Click
    
End Sub

Private Sub cmdopen_Click()
On Error GoTo CanErr
    With frmmain.CDialog1
        .CancelError = True
        .FileName = ""
        .DialogTitle = "Select add-in ActiveX Dll"
        .InitDir = FixPath(App.Path) & "add-ins\"
        .Filter = "ActiveX Dll(*.dll)|*.dll|"
        .ShowOpen
        
        If Not GetFileExt(.FileName) = "DLL" Then
            MsgBox "Invalid Filename.", vbCritical, "Invalid Filename"
            txtdllname.Text = ""
            Exit Sub
        End If
        
        If Len(.FileName) = 0 Then
            txtdllname.Text = ""
            Exit Sub
        Else
            txtdllname.Text = .FileTitle
        End If
CanErr:
        If Err = cdlCancel Then
            Err.Clear
        End If
        
    End With
    
End Sub

Private Sub Form_Load()
    MakeFlatControls frminstdll
End Sub
