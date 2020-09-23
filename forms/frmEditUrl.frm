VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Bookmark"
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
   Begin VB.TextBox txtvisdate 
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
      Left            =   3900
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5475
      Width           =   2220
   End
   Begin VB.CommandButton cmdmod2 
      Caption         =   "&Modify"
      Height          =   315
      Left            =   6165
      TabIndex        =   31
      Top             =   5475
      Width           =   975
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
      TabIndex        =   30
      Top             =   3840
      Width           =   3060
   End
   Begin VB.CommandButton cmdmod1 
      Caption         =   "&Modify"
      Height          =   315
      Left            =   2625
      TabIndex        =   29
      Top             =   5475
      Width           =   975
   End
   Begin VB.TextBox txtDateAdd 
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
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   5475
      Width           =   2220
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1980
      TabIndex        =   26
      Top             =   6225
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Max             =   32767
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txthits 
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
      TabIndex        =   25
      Top             =   6240
      Width           =   1545
   End
   Begin Project1.Flat Flat5 
      Height          =   255
      Left            =   6705
      TabIndex        =   20
      Top             =   4650
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
      TabIndex        =   19
      Top             =   4635
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin Project1.Flat2 Flat21 
      Height          =   6615
      Left            =   30
      TabIndex        =   18
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
      TabIndex        =   9
      Top             =   6885
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Apply"
      Height          =   350
      Left            =   3075
      TabIndex        =   7
      Top             =   6885
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4560
      TabIndex        =   8
      Top             =   6885
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
      Left            =   3900
      TabIndex        =   6
      Top             =   6240
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   2955
      Top             =   7995
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
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3045
      Width           =   6000
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   315
      Left            =   6435
      TabIndex        =   3
      Top             =   3030
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      Height          =   1800
      Left            =   315
      ScaleHeight     =   1740
      ScaleWidth      =   2505
      TabIndex        =   16
      Top             =   1020
      Width           =   2565
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   30
         ScaleHeight     =   1680
         ScaleWidth      =   2445
         TabIndex        =   21
         Top             =   30
         Width           =   2445
      End
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   3915
      TabIndex        =   5
      Top             =   4620
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4620
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
         TabIndex        =   14
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   45
         Width           =   330
      End
      Begin VB.Label txtrate 
         AutoSize        =   -1  'True
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
         TabIndex        =   15
         Top             =   60
         Width           =   60
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
      Text            =   "frmEditUrl.frx":0000
      Top             =   1020
      Width           =   3990
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
      TabIndex        =   4
      Top             =   3840
      Width           =   3285
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Last Visited Date:"
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
      TabIndex        =   33
      Top             =   5190
      Width           =   2460
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Added Date:"
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
      TabIndex        =   27
      Top             =   5190
      Width           =   2025
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Hits Level:"
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
      TabIndex        =   24
      Top             =   5970
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark display icon:"
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
      TabIndex        =   23
      Top             =   4350
      Width           =   2025
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
      TabIndex        =   22
      Top             =   4350
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmEditUrl.frx":0048
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Snap Shot"
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
      TabIndex        =   17
      Top             =   750
      Width           =   1830
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   330
      Picture         =   "frmEditUrl.frx":0D12
      Top             =   7770
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
      Left            =   3105
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   3555
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
      Top             =   3570
      Width           =   1515
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TBitmap
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



Private Sub cmdcancel_Click()
    ' Clean up
    pic2.Picture = Nothing
    txtsitedes.Text = ""
    txtsh.Text = ""
    txtbktitle.Text = ""
    txtbkurl.Text = ""
    txtvisdate.Text = ""
    txtDateAdd.Text = ""
    ImageCombo1.ComboItems.Clear

        
    Unload frmEdit
    
End Sub

Private Sub cmddown_Click()
On Error Resume Next
    If Val(txtrate.Caption) <= 0 Then Exit Sub
    txtrate.Caption = txtrate.Caption - 1

    imgrate(Val(txtrate.Caption) + 1).Picture = Nothing
    
End Sub

Private Sub cmdmod1_Click()
    ModiyDate = AddedDate
    MoveFrmToPos frmcal, -180, -180
End Sub

Private Sub cmdmod2_Click()
    ModiyDate = LastViewedDate
    MoveFrmToPos frmcal, -180, -180
    
End Sub

Private Sub cmdok_Click()
    cmdcancel_Click
    
End Sub

Private Sub cmdopen_Click()
On Error GoTo CanErr
    Cdlg1.CancelError = True
    Cdlg1.DialogTitle = "Open picture File"
    Cdlg1.Filter = "Widnows Bitmap Files(*.bmp)|*.bmp|"
    Cdlg1.InitDir = FixPath(App.Path) & "snapshots\"
    Cdlg1.ShowOpen
    
    If Len(Cdlg1.FileName) = 0 Then Exit Sub
    If Not (GetFileExt(Cdlg1.FileName)) = "BMP" Then
        MsgBox "This program does not currently support this file type.", vbInformation, frmEdit.Caption
        Exit Sub
    Else
        GetBmpinfo Cdlg1.FileName
        If BitmapInfo.biWidth <> 163 Or BitmapInfo.biHeight <> 112 Then
            MsgBox "This program only supports the following sizes:" _
            & DoubleCRLF & "Bitmap width = 163 pixels" & vbCrLf & "Bitmap Height = 112 pixels", vbInformation, frmEdit.Caption
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
    If Val(txtrate.Caption) > 5 Then
        txtrate.Caption = "5"
        Exit Sub
    Else
        imgrate(Val(txtrate.Caption)).Picture = Image3.Picture
    End If
    
End Sub

Private Sub cmdupdate_Click()
Dim mCopySnap As String

    If Len(Trim(txtbktitle.Text)) <= 0 Then
        MsgBox "You must include the name of the bookmark.", vbExclamation, frmEdit.Caption
        txtbktitle.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtbkurl.Text)) <= 0 Then
        MsgBox "You must include an address for the bookmark ex http://www.mywebsite.com", vbExclamation, frmEdit.Caption
        txtbkurl.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtsitedes.Text)) <= 0 Then
        MsgBox "You need enter a description for the bookmark", vbExclamation, frmEdit.Caption
        Exit Sub
    Else
        SnapFile = txtsh.Text
        EdURL.TSiteName = txtbktitle.Text
        EdURL.TSiteURL = txtbkurl.Text
        EdURL.TDateAdded = Format(txtDateAdd.Text, "Medium Date")
        EdURL.TAddLastVis = Format(txtvisdate.Text, "Medium Date")
        EdURL.TSiteDescription = txtsitedes.Text
        EdURL.THitCnt = Val(txthits.Text)
        EdURL.TRated = Val(txtrate.Caption)
        EdURL.TVieded = chkmark.Value
        EdURL.TIcon = LstIcon
        
        If UCase(TBookSnapShot) = "NONE" And UCase(SnapFile) = "NONE" Then
            EdURL.TWebCap = "none"
        Else
            EdURL.TWebCap = SnapFile
        End If
         
        EditSite SiteID, RecoredName
        cmdok.Enabled = True
        MsgBox "The new link has now been successfully modified.", vbInformation, frmEdit.Caption
        frmmain.lstsites.ListItems.Clear
        LoadSites RecoredName
    End If
    
    
End Sub




Private Sub Form_Load()
Dim catIndex As Long, I As Long
On Error Resume Next

    frmEdit.Icon = Nothing
    frmEdit.Caption = "Modify Bookmark - " & TBookMark
    MakeFlatControls frmEdit
    FlatBorder bar.hwnd, True
    
    Set ImageCombo1.ImageList = frmmain.ImageList4
    ImageCombo1.ComboItems.Clear
    For I = 1 To frmmain.ImageList4.ListImages.Count
        ImageCombo1.ComboItems.Add , "a" & I, "Display Icon " & I, I, I
    Next
    
    ImageCombo1.ComboItems(TvIcon).Selected = True
    ImageCombo1_Click
    
    txtsitedes.Text = TBookMarkDes
    txtsh.Text = TBookSnapShot
    txtbktitle.Text = TBookMark
    txtbkurl.Text = TBookURL
    txtrate.Caption = TRate
    chkmark.Value = TViewed
    UpDown1.Value = TBookHit
    
    txtDateAdd.Text = Format(TBookAddDate, "Medium Date")
    txtvisdate.Text = Format(TBookLastVist, "Medium Date")
    
    For I = 1 To TRate
        If I > 5 Then Exit For
        imgrate(I).Picture = Image3.Picture
    Next
    
    I = 0
    
    If FindFile(TBookSnapShot) Then
        pic2.Picture = LoadPicture(TBookSnapShot)
        Exit Sub
    Else
        pic2.Picture = LoadResPicture(101, vbResBitmap)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEdit = Nothing
    
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

Private Sub txthits_Change()
    If Not IsNumeric(txthits.Text) Then txthits.Text = 0
    UpDown1.Value = Val(txthits.Text)
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

Private Sub UpDown1_Change()
    txthits.Text = UpDown1.Value
    
End Sub

