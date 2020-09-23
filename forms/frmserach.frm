VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmserach 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serach Bookmarks"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear results"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4125
      TabIndex        =   10
      Top             =   3870
      Width           =   1290
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2655
      TabIndex        =   9
      Top             =   3870
      Width           =   1290
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7665
      Top             =   3675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmserach.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstfind 
      Height          =   2430
      Left            =   90
      TabIndex        =   3
      Top             =   1245
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4286
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Close"
      Height          =   350
      Left            =   5655
      TabIndex        =   4
      Top             =   3870
      Width           =   1290
   End
   Begin Project1.Flat Flat3 
      Height          =   240
      Left            =   4755
      TabIndex        =   8
      Top             =   225
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
      Left            =   1815
      TabIndex        =   7
      Top             =   210
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   476
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin VB.ComboBox cbocat 
      Height          =   315
      ItemData        =   "frmserach.frx":0352
      Left            =   1800
      List            =   "frmserach.frx":0354
      TabIndex        =   1
      Top             =   195
      Width           =   3225
   End
   Begin VB.TextBox txtfind 
      Height          =   315
      Left            =   1800
      MaxLength       =   128
      TabIndex        =   2
      Top             =   660
      Width           =   3225
   End
   Begin Project1.Flat2 Flat21 
      Height          =   1110
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   1958
      BackStyle       =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmserach.frx":0356
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      Height          =   195
      Left            =   915
      TabIndex        =   6
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serach for:"
      Height          =   195
      Left            =   915
      TabIndex        =   5
      Top             =   690
      Width           =   780
   End
End
Attribute VB_Name = "frmserach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Find()
    lstfind.ListItems.Clear
    If Len(Trim(txtfind.Text)) <= 0 Then
        MsgBox "You must enter in a serach string to find", vbInformation, frmmain.Caption
        txtfind.SetFocus
        Exit Sub
    Else
        FindSite txtfind.Text, cbocat.Text
        If Not SiteFound Then
            MsgBox "The serach string for " & Chr(34) & txtfind.Text & Chr(34) & " could not be found", vbInformation, frmserach.Caption
            cmdClear.Enabled = False
        Else
            cmdClear.Enabled = True
        End If
    End If
End Sub

Private Sub cbocat_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdcan_Click()
    cmdClear_Click
    txtfind.Text = ""
    Urlname = ""
    lstfind.ListItems.Clear
    Unload frmserach
    
End Sub

Private Sub cmdClear_Click()
    lstfind.ListItems.Clear
    cmdClear.Enabled = False
End Sub

Private Sub cmdFind_Click()
Dim I As Long, A As Long
    A = 0
    If cbocat.ListIndex = 0 Then
        If Len(Trim(txtfind.Text)) <= 0 Then
            MsgBox "You must enter in a serach string to find", vbInformation, frmmain.Caption
            txtfind.SetFocus
            Exit Sub
        Else
            For I = 1 To cbocat.ListCount - 1
                FindSite txtfind.Text, cbocat.List(I)
                If SiteFound Then
                    A = 1
                End If
            Next
            If A = 0 Then
                cmdClear.Enabled = False
                MsgBox "The serach string for " & Chr(34) & txtfind.Text & Chr(34) & " could not be found", vbInformation, frmserach.Caption
            Else
                cmdClear.Enabled = True
            End If
            I = 0
            A = 0
        End If
    Else
        Find
    End If
End Sub

Private Sub Form_Load()
Dim Col_Head As ColumnHeader, lstItem As ListItem
Dim I As Long, catIndex As Long

    MakeFlatControls frmserach
    Initcbocat cbocat
    
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
    
    catIndex = 0 ' clean var
    I = 0   ' clean var

    lstfind.ColumnHeaders.Add , , "Site Name", 2410
    Set Col_Head = lstfind.ColumnHeaders.Add(, , "Site Location", 2510)
    Set Col_Head = lstfind.ColumnHeaders.Add(, , "Date Added", 1080)
    Set Col_Head = lstfind.ColumnHeaders.Add(, , "Last Viewed", 1080)
    Set Col_Head = lstfind.ColumnHeaders.Add(, , "Hits", 600)
    Set frmserach.Icon = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmserach = Nothing
    
End Sub

Private Sub lstfind_DblClick()
Dim mkey As String, RetVal As Long

    If lstfind.ListItems.Count = 0 Then Exit Sub
    RetVal = TOpenSite(frmmain.hwnd, TBookURL)
    If RetVal = 2 Then
        MsgBox "Error opening " & Chr(34) & lstfind.SelectedItem.Text & Chr(34), vbCritical, "Error opening location"
    End If
    
End Sub

Private Sub lstfind_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TBookURL = Trim(lstfind.SelectedItem.SubItems(1))
End Sub

Private Sub lstfind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstfind.ListItems.Count = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu frmPopUpmenu.mnuopenW
    End If
End Sub

Private Sub txtfind_Change()
    If Len(Trim(txtfind.Text)) = 0 Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    
End Sub

Private Sub txtfind_GotFocus()
    txtfind.BackColor = Config.Hightlight
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFind_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtfind_LostFocus()
    txtfind.BackColor = vbWhite
End Sub

