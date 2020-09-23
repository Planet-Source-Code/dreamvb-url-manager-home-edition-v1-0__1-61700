VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddUrl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Bookmark"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Flat Flat5 
      Height          =   255
      Left            =   6705
      TabIndex        =   28
      Top             =   5130
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   450
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin Project1.Flat Flat4 
      Height          =   285
      Left            =   3930
      TabIndex        =   27
      Top             =   5115
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin Project1.Flat Flat3 
      Height          =   240
      Left            =   3375
      TabIndex        =   26
      Top             =   5130
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
      Left            =   375
      TabIndex        =   25
      Top             =   5115
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   476
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin Project1.Flat2 Flat21 
      Height          =   6615
      Left            =   90
      TabIndex        =   24
      Top             =   90
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11668
      BackStyle       =   0
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   6015
      TabIndex        =   12
      Top             =   6825
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Apply"
      Height          =   350
      Left            =   3030
      TabIndex        =   10
      Top             =   6825
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4560
      TabIndex        =   11
      Top             =   6825
      Width           =   1215
   End
   Begin VB.CheckBox chkmark 
      Caption         =   "Mark bookmark as unviewed."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3870
      TabIndex        =   9
      Top             =   5865
      Width           =   3255
   End
   Begin VB.ComboBox cbocat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   5100
      Width           =   3285
   End
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   75
      Top             =   6870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtsh 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3345
      Width           =   6000
   End
   Begin VB.CheckBox chksh 
      Caption         =   "Use different Snapshot:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   2985
      Width           =   5610
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   315
      Left            =   6435
      TabIndex        =   4
      Top             =   3330
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      Height          =   1800
      Left            =   315
      ScaleHeight     =   1740
      ScaleWidth      =   2505
      TabIndex        =   19
      Top             =   1020
      Width           =   2565
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   30
         ScaleHeight     =   1680
         ScaleWidth      =   2445
         TabIndex        =   29
         Top             =   30
         Width           =   2445
      End
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   3915
      TabIndex        =   8
      Top             =   5100
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox bar 
      Height          =   345
      Left            =   330
      ScaleHeight     =   285
      ScaleWidth      =   3225
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5910
      Width           =   3285
      Begin VB.CommandButton cmdup 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2475
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   330
      End
      Begin VB.CommandButton cmddown 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2865
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   330
      End
      Begin VB.Label txtrate 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   18
         Top             =   60
         Width           =   120
      End
      Begin VB.Image imgrate 
         Height          =   240
         Index           =   5
         Left            =   2160
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgrate 
         Height          =   240
         Index           =   4
         Left            =   1815
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgrate 
         Height          =   240
         Index           =   3
         Left            =   1425
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgrate 
         Height          =   240
         Index           =   2
         Left            =   1050
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgrate 
         Height          =   240
         Index           =   1
         Left            =   645
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.TextBox txtsitedes 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   3000
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAddUrl.frx":0000
      Top             =   1020
      Width           =   3990
   End
   Begin VB.TextBox txtbkurl 
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
      Left            =   3930
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "http://www."
      Top             =   4320
      Width           =   3060
   End
   Begin VB.TextBox txtbktitle 
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
      Left            =   360
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4320
      Width           =   3285
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmAddUrl.frx":0048
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Rating Leval:"
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
      Left            =   390
      TabIndex        =   23
      Top             =   5625
      Width           =   2070
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Category:"
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
      Left            =   390
      TabIndex        =   22
      Top             =   4830
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Favorite icon:"
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
      Left            =   3930
      TabIndex        =   21
      Top             =   4830
      Width           =   2115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Snapshot"
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
      Left            =   315
      TabIndex        =   20
      Top             =   750
      Width           =   1740
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1350
      Picture         =   "frmAddUrl.frx":0912
      Top             =   6960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Description"
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
      Left            =   3015
      TabIndex        =   14
      Top             =   750
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Address:"
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
      Left            =   3930
      TabIndex        =   13
      Top             =   4035
      Width           =   1710
   End
   Begin VB.Label lblsitename 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Name:"
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
      Left            =   390
      TabIndex        =   0
      Top             =   4050
      Width           =   1515
   End
End
Attribute VB_Name = "frmAddUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TBitmap '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type

Private SnapFile As String
Private LstIcon As Integer
Private BitmapInfo As TBitmap

Private Sub GetBmpinfo(sFilename As String)
Dim hFile As Long
    hFile = FreeFile ' Pointer to freefile
    
    Open sFilename For Binary Access Read As #hFile
        ' Open the file in binary mode
        Get #hFile, 15, BitmapInfo
        ' Extract the bitmap information
    Close #hFile ' Close the file
    
End Sub

Private Sub cbocat_Click()
    RecoredName = cbocat.Text
End Sub

Private Sub cbocat_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chksh_Click()
    txtsh.Enabled = chksh.Value
    cmdopen.Enabled = chksh.Value
End Sub

Private Sub cmdcancel_Click()
    ' Clean up
    DMBook.BookName = ""
    DMBook.BookUrl = ""
    FromWeb = False
    pic2.Picture = Nothing
    txtsitedes.Text = ""
    txtsh.Text = ""
    txtbktitle.Text = ""
    txtbkurl.Text = ""
    cbocat.Clear
    ImageCombo1.ComboItems.Clear
    If FindFile(SnapFile) = True Then
        Kill SnapFile
    End If
        
    Unload frmAddUrl
    
End Sub

Private Sub cmddown_Click()
On Error Resume Next
    If Val(txtrate.Caption) <= 0 Then Exit Sub
    txtrate.Caption = txtrate.Caption - 1

    imgrate(Val(txtrate.Caption) + 1).Picture = Nothing
    
End Sub

Private Sub cmdok_Click()
    cmdcancel_Click
    
End Sub

Private Sub cmdopen_Click()
Dim lzFilename As String
On Error GoTo CanErr
    Cdlg1.CancelError = True
    Cdlg1.DialogTitle = "Open picture File"
    Cdlg1.Filter = "Widnows Bitmap Files(*.bmp)|*.bmp|"
    Cdlg1.InitDir = FixPath(App.Path) & "screens\"
    Cdlg1.ShowOpen
    
    If Len(Cdlg1.FileName) = 0 Then Exit Sub
    If Not (GetFileExt(Cdlg1.FileName)) = "BMP" Then
        MsgBox "This program does not currently support this file type.", vbInformation, frmAddUrl.Caption
        Exit Sub
    Else
        GetBmpinfo Cdlg1.FileName
        If BitmapInfo.biWidth <> 163 Or BitmapInfo.biHeight <> 112 Then
            MsgBox "This program only supports the following sizes:" _
            & DoubleCRLF & "Bitmap width = 163 pixels" & vbCrLf & "Bitmap Height = 112 pixels", vbInformation, frmAddUrl.Caption
            Exit Sub
        Else
            txtsh.Text = Cdlg1.FileName
            pic2.Picture = LoadPicture(Cdlg1.FileName)
        End If
    End If
    
CanErr:
    If Err Then
        Err.Clear
    End If
    
End Sub

Private Sub cmdup_Click()
    txtrate.Caption = txtrate.Caption + 1
    If Val(txtrate.Caption) > 5 Then Exit Sub
    imgrate(Val(txtrate.Caption)).Picture = Image3.Picture
    
End Sub

Private Sub cmdupdate_Click()
Dim mCopySnap As String
On Error Resume Next
    If Len(Trim(txtbktitle.Text)) <= 0 Then
        MsgBox "You must include the name of the bookmark.", vbExclamation, frmAddUrl.Caption
        txtbktitle.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtbkurl.Text)) <= 0 Then
        MsgBox "You must include an address for the bookmark ex http://www.mywebsite.com", vbExclamation, frmAddUrl.Caption
        txtsiteurl.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtsitedes.Text)) <= 0 Then
        MsgBox "You need enter a small description about the web site.", vbExclamation, frmAddUrl.Caption
        Exit Sub
    Else
        EdURL.TSiteName = txtbktitle.Text
        EdURL.TSiteURL = txtbkurl.Text
        EdURL.TDateAdded = Format(Date, "Medium Date")
        EdURL.TAddLastVis = Format(Date, "Medium Date")
        EdURL.TSiteDescription = txtsitedes.Text
        EdURL.THitCnt = 0
        EdURL.TRated = Val(txtrate.Caption)
        EdURL.TVieded = chkmark.Value
        EdURL.TIcon = LstIcon
        
        If Not CBool(chksh.Value) And FromWeb = False Then
            EdURL.TWebCap = "none"
        ElseIf CBool(chksh.Value) = True And Len(txtsh.Text) = 0 Then
            MsgBox "You must select a filename."
            Exit Sub
        ElseIf FromWeb = True And CBool(chksh.Value) = False Then
            mCopySnap = "snp_" & Day(Date) & Second(Time) & Year(Date) & Hex(Second(Time) * Rnd) & ".bmp"
            FileCopy SnapFile, FixPath(App.Path) & "snapshots\" & mCopySnap
            EdURL.TWebCap = FixPath(App.Path) & "snapshots\" & mCopySnap
        Else
            EdURL.TWebCap = txtsh.Text
        End If

        AddNewUrl RecoredName
        cmdok.Enabled = True
        MsgBox "The new link has now been successfully updated.", vbInformation, frmAddUrl.Caption
        frmmain.lstsites.ListItems.Clear
        LoadSites RecoredName
    End If
    If Err Then MsgBox Err.Description
    
End Sub

Private Sub Form_Load()
Dim I As Long, catIndex As Long

    frmAddUrl.Icon = Nothing
    MakeFlatControls frmAddUrl
    FlatBorder bar.hwnd, True
    txtbkurl.SelStart = Len(txtbkurl.Text)
    
    If FromWeb Then
        txtbktitle.Text = DMBook.BookName
        txtbkurl.Text = DMBook.BookUrl
    End If
    
    Set ImageCombo1.ImageList = frmmain.ImageList4
    ImageCombo1.ComboItems.Clear
    For I = 1 To frmmain.ImageList4.ListImages.Count
        ImageCombo1.ComboItems.Add , "a" & I, "Display Icon " & I, I, I
    Next
    
    ImageCombo1.ComboItems(1).Selected = True
    ImageCombo1_Click
    chksh_Click
    I = 0
    
    Initcbocat cbocat ' Load in the items in the combo box
    cbocat.RemoveItem 0 ' Remove the first item
    
    For I = 0 To cbocat.ListCount - 1 ' loop to the end of the combo list count
        If UCase(RecoredName) = UCase(cbocat.List(I)) Then ' check if recoredname is in the list
            catIndex = I ' assign the index of found item
        End If
    Next
    
    If catIndex <= 0 Then
        cbocat.ListIndex = 0 ' set listindex to top item
    Else
        cbocat.ListIndex = catIndex ' set listindex to found item index
    End If
    
    I = 0
    catIndex = 0
    
    SnapFile = GetTempFolder & "scrtemp.bmp"
    
    If Not FromWeb Then
        pic2.Picture = LoadResPicture(101, vbResBitmap)
        Exit Sub
    ElseIf FromWeb = True And FindFile(SnapFile) = False Then
        pic2.Picture = LoadResPicture(101, vbResBitmap)
        FromWeb = False
    Else
        pic2.Picture = LoadPicture(SnapFile)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAddUrl = Nothing
    
End Sub

Private Sub ImageCombo1_Click()
On Error Resume Next
    LstIcon = Val(Right(ImageCombo1.SelectedItem.Key, Len(ImageCombo1.SelectedItem.Key) - 1))
    
End Sub

Private Sub ImageCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ImageCombo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtsh_GotFocus()
    txtsh.BackColor = Config.Hightlight
End Sub

Private Sub txtsh_LostFocus()
    txtsh.BackColor = vbWhite
End Sub

Private Sub txtsitedes_GotFocus()
    txtsitedes.BackColor = Config.Hightlight
    
End Sub

Private Sub txtsitedes_LostFocus()
    txtsitedes.BackColor = vbWhite
    
End Sub

Private Sub txtbktitle_GotFocus()
    txtbktitle.BackColor = Config.Hightlight
    
End Sub

Private Sub txtbktitle_LostFocus()
    txtbktitle.BackColor = vbWhite
    
End Sub

Private Sub txtbkurl_GotFocus()
    txtbkurl.BackColor = Config.Hightlight
    
End Sub

Private Sub txtbkurl_LostFocus()
    txtbkurl.BackColor = vbWhite
    
End Sub
