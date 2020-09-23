VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   7215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3825
      TabIndex        =   22
      Top             =   7140
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2415
      TabIndex        =   21
      Top             =   7140
      Width           =   1215
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "&Apply"
      Height          =   350
      Left            =   990
      TabIndex        =   20
      Top             =   7140
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6900
      Left            =   120
      TabIndex        =   8
      Top             =   45
      Width           =   4965
      Begin VB.CommandButton cmdopen 
         Caption         =   "....."
         Height          =   315
         Left            =   4440
         TabIndex        =   12
         Top             =   1920
         Width           =   345
      End
      Begin VB.TextBox txtdbpath 
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
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1935
         Width           =   4155
      End
      Begin Project1.Flat Flat5 
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   5115
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
         Left            =   195
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   5085
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.OptionButton optdouble 
         Caption         =   "Use double click to open items."
         Height          =   240
         Left            =   225
         TabIndex        =   16
         Top             =   4260
         Width           =   3495
      End
      Begin VB.OptionButton optsingle 
         Caption         =   "Use single click to open items."
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
         Left            =   225
         TabIndex        =   15
         Top             =   3975
         Width           =   4035
      End
      Begin Project1.Flat Flat3 
         Height          =   240
         Left            =   1785
         TabIndex        =   27
         Top             =   5880
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
         Left            =   210
         TabIndex        =   26
         Top             =   5865
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   476
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
      End
      Begin VB.ComboBox cboweb 
         Height          =   315
         Left            =   195
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5850
         Width           =   1860
      End
      Begin Project1.Line3D Line3D1 
         Height          =   30
         Left            =   135
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4620
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   53
      End
      Begin VB.PictureBox pichg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4590
         ScaleHeight     =   240
         ScaleWidth      =   180
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3105
         Width           =   210
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "...."
         Height          =   270
         Left            =   4050
         TabIndex        =   14
         Top             =   3105
         Width           =   480
      End
      Begin VB.PictureBox piccol 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4590
         ScaleHeight     =   240
         ScaleWidth      =   180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2640
         Width           =   210
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "...."
         Height          =   270
         Left            =   4050
         TabIndex        =   13
         Top             =   2640
         Width           =   480
      End
      Begin VB.CheckBox chktips 
         Caption         =   "Show tip of the day at start-up"
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
         Left            =   210
         TabIndex        =   19
         Top             =   6420
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.TextBox txtwhois 
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
         Left            =   195
         TabIndex        =   10
         Text            =   "whois.internic.net"
         Top             =   1245
         Width           =   4545
      End
      Begin VB.TextBox txtsmtp 
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
         Left            =   195
         TabIndex        =   9
         Text            =   "mail.yourISP.com"
         Top             =   525
         Width           =   4545
      End
      Begin Project1.Line3D Line3D2 
         Height          =   30
         Left            =   150
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2475
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   53
      End
      Begin Project1.Line3D Line3D3 
         Height          =   30
         Left            =   135
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3525
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   53
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bookmarks Database"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   4
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following fav icon for the categories"
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
         Index           =   3
         Left            =   180
         TabIndex        =   1
         Top             =   4785
         Width           =   3765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following colour to mark new items."
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
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   2640
         Width           =   3765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following heightlight colour."
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
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   3105
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following as my default web browser."
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
         Index           =   6
         Left            =   225
         TabIndex        =   0
         Top             =   5550
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following options to open items."
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
         Left            =   180
         TabIndex        =   7
         Top             =   3660
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Whois lookup server"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   3
         Top             =   990
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Server"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   255
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CboIdx As Integer
Private TvIconInx As Integer
Private TvIdx As Integer
Private OpenItemsInx As Integer

Sub SetOptOpenItems()
    Select Case Val(Config.mOpenItems)
        Case 1
            optsingle.Value = True
        Case 2
            optdouble.Value = True
    End Select
    
End Sub
Private Sub cboweb_Click()
    CboIdx = cboweb.ListIndex
End Sub

Private Sub cboweb_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cboweb_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd1_Click()
On Error GoTo ColErr
    frmmain.CDialog1.Flags = cdlCCFullOpen
    frmmain.CDialog1.CancelError = True
    frmmain.CDialog1.ShowColor
    piccol.BackColor = frmmain.CDialog1.Color
    Exit Sub
ColErr:
    If Err Then Exit Sub
    
    
End Sub

Private Sub cmd2_Click()
On Error GoTo ColErr
    frmmain.CDialog1.Flags = cdlCCFullOpen
    frmmain.CDialog1.CancelError = True
    frmmain.CDialog1.ShowColor
    pichg.BackColor = frmmain.CDialog1.Color
    Exit Sub
ColErr:
    If Err Then Exit Sub
    
End Sub

Private Sub cmdopen_Click()
On Error GoTo CanErr
    With frmmain.CDialog1
        .CancelError = True
        .DialogTitle = "Open Bookmarks Database"
        .Filter = "Microsoft Access Databases(*.mdb)|*.mdb|"
        .InitDir = FixPath(App.Path)
        .ShowOpen
        If Not GetFileExt(.FileName) = "MDB" Then
            MsgBox "This is not a valid Access Database filename, or its format is not supported.", vbExclamation, .DialogTitle
            Exit Sub
        Else
            txtdbpath.Text = .FileName
        End If
        Exit Sub
CanErr:
    If Err = cdlCancel Then
        Err.Clear
    End If
    
    End With
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub ImageCombo1_Click()
    TvIdx = ImageCombo1.SelectedItem.Index
    TvIconInx = Val(Right(ImageCombo1.SelectedItem.Key, 1))
End Sub

Private Sub ImageCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ImageCombo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub optdouble_Click()
    OpenItemsInx = 2
End Sub

Private Sub optsingle_Click()
    OpenItemsInx = 1
End Sub

Private Sub piccol_Click()
    cmd1_Click
End Sub

Private Sub pichg_Click()
    cmd2_Click
End Sub

Private Sub txtdbpath_GotFocus()
    txtdbpath.BackColor = Config.Hightlight
End Sub

Private Sub txtdbpath_LostFocus()
    txtdbpath.BackColor = vbWhite
End Sub

Private Sub txtsmtp_GotFocus()
    txtsmtp.BackColor = Config.Hightlight
End Sub

Private Sub txtsmtp_LostFocus()
    txtsmtp.BackColor = vbWhite
End Sub

Private Sub txtwhois_GotFocus()
    txtwhois.BackColor = Config.Hightlight
End Sub

Private Sub txtwhois_LostFocus()
    txtwhois.BackColor = vbWhite
End Sub

Private Sub cmdapply_Click()
Dim tFile As Long

    Config.mSMTP_serv = txtsmtp.Text
    Config.mWHOIS_serv = txtwhois.Text
    
    ' New config settings

    SaveSetting "DMUrlMan", "Config", "FirstRun", "1"
    SaveSetting "DMUrlMan", "Config", "Smtpserver", txtsmtp.Text
    SaveSetting "DMUrlMan", "Config", "Whois", txtwhois.Text
    SaveSetting "DMUrlMan", "Config", "ShowTips", CStr(chktips.Value)
    SaveSetting "DMUrlMan", "Config", "Browser", CStr(CboIdx)
    
    SaveSetting "DMUrlMan", "Config", "Highlight", CStr(pichg.BackColor)
    SaveSetting "DMUrlMan", "Config", "NewItems", CStr(piccol.BackColor)
    
    SaveSetting "DMUrlMan", "Config", "TvFavIcon", CStr(TvIconInx)
    SaveSetting "DMUrlMan", "Config", "TvFavIconIdx", CStr(TvIdx)
    SaveSetting "DMUrlMan", "Config", "OpenItems", CStr(OpenItemsInx)
    SaveSetting "DMUrlMan", "Config", "Db", txtdbpath.Text
    SaveSetting "DMUrlMan", "Config", "WebURL", "http://www.eraystudios.com"
    SaveSetting "DMUrlMan", "Config", "AppPath", FixPath(App.Path)
    
    If Not FindFile(txtdbpath.Text) Or Len(txtdbpath.Text) = 0 Then
        MsgBox "Cannot find file" & vbCrLf & txtdbpath.Text & _
        DoubleCRLF & "Please insure that the file has not been moved or deleted.", vbCritical, "File not found"
        txtdbpath.SetFocus
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
        ReadServConfig
        frmmain.InitDB Config.mbkDatabase
    End If
    
End Sub

Private Sub cmdcan_Click()
    CboIdx = 0
    OpenItemsInx = 0
    Unload frmconfig
End Sub

Private Sub cmdok_Click()
    Unload frmconfig
End Sub

Private Sub Form_Load()
    Dim hSysMenu As Long, nCnt As Long, I, K As Long
    
    hSysMenu = GetSystemMenu(frmconfig.hwnd, False)
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE ' Remove the seperator
            DrawMenuBar frmconfig.hwnd
        End If
    End If
    
    cboweb.AddItem "Internet Explorer"
    cboweb.AddItem "Netscape"
    cboweb.AddItem "Opera"
    cboweb.AddItem "Mozilla"
    cboweb.AddItem "Mozilla Firefox"
    
    frmconfig.Icon = Nothing
    MakeFlatControls frmconfig
    
    Set ImageCombo1.ImageList = frmmain.ImageList1
    ImageCombo1.ComboItems.Clear
  
    For I = 1 To frmmain.ImageList1.ListImages.Count Step 3 - 1
        K = K + 1
        ImageCombo1.ComboItems.Add , "a" & I, "Display Icon " & K, I, I
    Next
    
    ImageCombo1.ComboItems.Remove ImageCombo1.ComboItems.Count
    I = 0
    
    chktips.Value = Val(Config.ShowTips)
    txtsmtp.Text = Config.mSMTP_serv
    txtwhois.Text = Config.mWHOIS_serv
    txtdbpath.Text = Config.mbkDatabase
    pichg.BackColor = Val(Config.Hightlight)
    piccol.BackColor = Val(Config.NewItems)
    cboweb.ListIndex = Val(Config.defBrowser)
    ImageCombo1.ComboItems(1).Selected = True
    SetOptOpenItems
    
    
    ImageCombo1.ComboItems(Val(Config.FavIdx)).Selected = True
    ImageCombo1_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmconfig = Nothing
End Sub



