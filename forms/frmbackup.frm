VERSION 5.00
Begin VB.Form frmrestore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Database"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddb 
      Caption         =   "...."
      Enabled         =   0   'False
      Height          =   315
      Left            =   3675
      TabIndex        =   7
      Top             =   1230
      Width           =   420
   End
   Begin VB.TextBox txtdb 
      Height          =   285
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1230
      Width           =   3465
   End
   Begin VB.CommandButton cmdbk 
      Caption         =   "...."
      Height          =   315
      Left            =   3675
      TabIndex        =   4
      Top             =   480
      Width           =   420
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   2850
      TabIndex        =   3
      Top             =   1905
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1485
      TabIndex        =   2
      Top             =   1905
      Width           =   1215
   End
   Begin VB.TextBox txtbackup 
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   495
      Width           =   3465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restore file To:"
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
      Left            =   165
      TabIndex        =   5
      Top             =   915
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to Restore:"
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
      Left            =   165
      TabIndex        =   0
      Top             =   210
      Width           =   1305
   End
End
Attribute VB_Name = "frmrestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lzFile As String
Private Function GetFileVersion(lzFile As String) As Long
Dim jFile As Long
    ' This function gets the version of the backup set
    jFile = FreeFile ' Pointer to free file
    Open lzFile For Binary As #jFile ' Open file in binary mode
        Get #jFile, , dbkHead ' Get file head info
    Close #jFile ' Close the file
    
    GetFileVersion = Asc(dbkHead.Version) ' Extract the file version
    
End Function
Private Sub cmdbk_Click()
On Error GoTo CanErr
Dim ans As Integer

    With frmmain.CDialog1
        .CancelError = True
        .DialogTitle = "DM Bookmarks Backup Set"
        .Filter = "DM Bookmarks Backup Set(*.dbk)|*.dbk|"
        .InitDir = Abspath
        .FileName = ""
        .ShowOpen
        If Not GetFileExt(.FileName) = "DBK" Then
            MsgBox "This is not a valid M Bookmarks Backup Set filename, or its format is not supported.", vbExclamation, .DialogTitle
            cmddb.Enabled = False
            Exit Sub
        Else
            If GetFileVersion(.FileName) <> 1 Then
                MsgBox "Incorrect File version" & vbCrLf & "The file cannot be restored", vbCritical, "Incorrect File version"
                Exit Sub
            Else
                lzFile = Left(.FileTitle, Len(.FileTitle) - 3) + "mdb"
                txtbackup.Text = .FileName
                cmddb.Enabled = True
            End If
        End If
        Exit Sub
CanErr:
        If Err = cdlCancel Then
            Err.Clear
            cmddb.Enabled = False
        End If
        
    End With
    
End Sub

Private Sub cmdcancel_Click()
    txtbackup.Text = ""
    txtdb.Text = ""
    lzFile = ""
    Unload frmrestore
End Sub

Private Sub cmddb_Click()
Dim fName As String
    fName = GetFolder(frmrestore.hwnd, "Choose Folder:")
    
    If Len(fName) <= 0 Then
        cmdRestore.Enabled = False
    Else
        txtdb.Text = FixPath(fName) & lzFile
        cmdRestore.Enabled = True
    End If
    
    fName = ""
    
End Sub

Private Sub cmdRestore_Click()
Dim ans As Integer, iResult As Long

On Error Resume Next

    If FindFile(txtdb.Text) Then
        ans = MsgBox(txtdb.Text & vbCrLf & "Already exists do you want to replace this file with the new one?", vbYesNo Or vbQuestion, frmrestore.Caption)
        If ans = vbNo Then
            Exit Sub
        Else
            SetAttr txtdb.Text, vbNormal
            Kill txtdb.Text
            If Err Then
                MsgBox Err.Description & DoubleCRLF & "The database your trying to restore may be in use" _
                & vbCrLf & "Please try selecting a different name or close any Instances that may be running." _
                & vbCrLf & "Or try selecting a different location.", vbCritical, "Error_" & Err.Number
            Else
                RestoreDb txtbackup.Text, txtdb.Text
                MsgBox "Your Database has now been successfully restored.", vbInformation, frmrestore.Caption
                ans = 0
                lzFile = ""
                Set d_base = OpenDatabase(Config.mbkDatabase, False, False, ";pwd=idkfa")
                cmdcancel_Click
                Exit Sub
            End If
        End If
    Else
        RestoreDb txtbackup.Text, txtdb.Text
        MsgBox "Your Database has now been successfully restored.", vbInformation, frmrestore.Caption
        ans = 0
        lzFile = ""
        Set d_base = OpenDatabase(Config.mbkDatabase, False, False, ";pwd=idkfa")
        cmdcancel_Click
    End If
    
End Sub

Private Sub Form_Load()
    FlatBorder txtbackup.hwnd, True
    FlatBorder txtdb.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmrestore = Nothing
End Sub
