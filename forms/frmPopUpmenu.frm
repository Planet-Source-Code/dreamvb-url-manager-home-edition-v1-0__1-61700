VERSION 5.00
Begin VB.Form frmPopUpmenu 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuedit 
      Caption         =   "&edit"
      Begin VB.Menu mnubkmark 
         Caption         =   "&Bookmark"
         Begin VB.Menu mnuaddnew 
            Caption         =   "&New Bookmark"
         End
         Begin VB.Menu mnueditURL 
            Caption         =   "&Modify Bookmark"
         End
         Begin VB.Menu mnudelURL 
            Caption         =   "&Delete Bookmark"
         End
         Begin VB.Menu mnublank4 
            Caption         =   "-"
         End
         Begin VB.Menu mnubkmove 
            Caption         =   "&Move Bookmark"
         End
         Begin VB.Menu mnusavedtop 
            Caption         =   "&Create Shortcut"
         End
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy Bookmark"
         Begin VB.Menu mnubkname 
            Caption         =   "&Name"
         End
         Begin VB.Menu mnubklocation 
            Caption         =   "&Location"
         End
         Begin VB.Menu mnubkadded 
            Caption         =   "&Added Date"
         End
         Begin VB.Menu mnublank1 
            Caption         =   "-"
         End
         Begin VB.Menu mnubklastv 
            Caption         =   "&Last Viewed"
         End
         Begin VB.Menu mnubkhits 
            Caption         =   "&Hits"
         End
         Begin VB.Menu mnubkhtml 
            Caption         =   "&HTML Code"
         End
      End
      Begin VB.Menu mnuopenW 
         Caption         =   "&Open With"
         Begin VB.Menu mnuIE 
            Caption         =   "&Internet Explorer"
         End
         Begin VB.Menu mnuns 
            Caption         =   "&Netscape Navigator"
         End
         Begin VB.Menu mnuma 
            Caption         =   "&Mozilla"
         End
         Begin VB.Menu mnuoa 
            Caption         =   "&Opera"
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnutools 
         Caption         =   "&Tools"
         Begin VB.Menu mnusaveweb 
            Caption         =   "&Save Web Page"
         End
         Begin VB.Menu mnuslnk 
            Caption         =   "&Share Bookmark"
         End
         Begin VB.Menu mnublank3 
            Caption         =   "-"
         End
         Begin VB.Menu mnubkstat 
            Caption         =   "Bookmark &Status"
         End
         Begin VB.Menu mnuping 
            Caption         =   "&Ping"
         End
         Begin VB.Menu mnuwhois 
            Caption         =   "&Whois Lookup"
         End
      End
   End
   Begin VB.Menu mnucat 
      Caption         =   "&cat"
      Begin VB.Menu mnuaddcat 
         Caption         =   "&Add Category"
      End
      Begin VB.Menu mnudelcat 
         Caption         =   "&Delete Category"
      End
      Begin VB.Menu mnupasscat 
         Caption         =   "&Rename Category"
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufindcat 
         Caption         =   "&Find Category"
      End
      Begin VB.Menu mnufindincat 
         Caption         =   "&Find in Category"
      End
      Begin VB.Menu mnugenhtmlpage 
         Caption         =   "&Generate Web Page"
      End
   End
   Begin VB.Menu mnutray 
      Caption         =   "&tray"
      Begin VB.Menu mnures 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuab 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuex 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuser 
      Caption         =   "&Serach"
      Begin VB.Menu MnuName 
         Caption         =   "cap"
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmPopUpmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DelBookmark()
Dim ans As Integer

    ans = MsgBox("Your about to delete the bookmark for " & _
    TBookMark & "." & DoubleCRLF & "Are you sure you want like to continue?", vbYesNo Or vbQuestion, "Delete Bookmark")
    If ans = vbNo Then
        Exit Sub
    ElseIf ans = vbYes Then
        DeleteURL SiteID, RecoredName
        RemoveLVItem LstIndex, frmmain.lstsites
    End If
    
    MsgBox "The bookmark has been successfully deleted.", vbInformation, "Delete bookmark"
    ' Clear up vars
    If FindFile(TBookSnapShot) Then
        Kill TBookSnapShot
    End If
    
    ans = 0
    SiteID = 0
    LstIndex = 0
    LoadSites RecoredName   ' Reload all the bookmarks back in
    
End Sub
Sub DelCat()
Dim ans As Integer

    ans = MsgBox("Warning you are about to delete the category for " & RecoredName & _
    DoubleCRLF & "Deleting this category will also delete all the bookmarks in this category." _
    & DoubleCRLF & "Are you sure you want to delete this category?", vbYesNo Or vbQuestion, "Delete Category")
    If ans = vbNo Then Exit Sub
    
    DeleteTable RecoredName
    RemoveTvItem TvIndex, frmmain.tv1
    MsgBox "The category has now been successfully deleted.", vbInformation, "Delete Category"
    
    TvIndex = 0
    RecoredName = ""
    ans = 0
    
End Sub

Private Sub mnuab_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub mnuaddcat_Click()
    If Len(RecoredName) > 0 Then
        frmaddcat.Show vbModal, frmmain
    End If
    
End Sub

Private Sub mnuaddnew_Click()
    FromWeb = False ' Book mark was not added form the web broswer
    frmAddUrl.Show vbModal, frmmain ' show add new bookmark dialog
    
End Sub

Private Sub mnubkadded_Click()
    CopyCommand m_SiteAddedDate, frmmain.lstsites.SelectedItem.SubItems(2) ' Copy bookmark added date

End Sub

Private Sub mnubkhits_Click()
    CopyCommand m_SiteHits, frmmain.lstsites.SelectedItem.SubItems(4)
End Sub

Private Sub mnubkhtml_Click()
Dim aStr As String

    aStr = "<a href=" & Chr(34) & frmmain.lstsites.SelectedItem.SubItems(1) & Chr(34) & " target=" & Chr(34) & "_blank" _
    & Chr(34) & ">" & frmmain.lstsites.SelectedItem.Text & "</a>"
    CopyCommand m_HtmlCode, aStr
    aStr = ""
    
End Sub

Private Sub mnubklastv_Click()
    CopyCommand m_SiteLastVistDate, frmmain.lstsites.SelectedItem.SubItems(3)
End Sub

Private Sub mnubklocation_Click()
    CopyCommand m_SiteURL, frmmain.lstsites.SelectedItem.SubItems(1) ' Copy bookmark URL

End Sub

Private Sub mnubkmove_Click()
    frmmoveto.Show vbModal, frmmain ' show moveto bookmark dialog
End Sub

Private Sub mnubkname_Click()
    CopyCommand m_SiteName, frmmain.lstsites.SelectedItem.Text  ' Copy bookmark name

End Sub

Private Sub mnubkstat_Click()
    frmstatus.Show vbModal, frmmain
End Sub

Private Sub mnudelcat_Click()
    DelCat
End Sub

Private Sub mnudelURL_Click()
    DelBookmark
End Sub

Private Sub mnueditURL_Click()
    frmEdit.Show vbModal, frmmain
End Sub

Private Sub mnuex_Click()
    Unload frmmain
End Sub

Private Sub mnufindcat_Click()
    frmfindcat.Show vbModal, frmmain
End Sub

Private Sub mnufindincat_Click()
    frmserach.Show vbModal, frmmain
End Sub

Private Sub mnugenhtmlpage_Click()
    frmgenhtm.Show vbModal, frmmain
End Sub

Private Sub mnuIE_Click()
On Error Resume Next
    ChDir WebBor.IE
    tRet = WinExec("IEXPLORE.EXE " & TBookURL, 3)
    ChDir Abspath
    
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
    
End Sub

Private Sub mnuma_Click()
On Error Resume Next
    ChDir WebBor.Mozilla
    tRet = WinExec("Mozilla.exe " & TBookURL, 3)
    ChDir Abspath
    
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
    
End Sub

Private Sub MnuName_Click(Index As Integer)
Static iCheck As Integer, I As Long

    iCheck = Not iCheck
    SerachOption = Index
    MnuName(Index).Checked = Abs(iCheck)
    
    For I = 0 To MnuName.UBound
        If Not I = Index Then
            MnuName(I).Checked = False
        Else
            MnuName(I).Checked = True
        End If
    Next
    I = 0
    iCheck = 0
    
End Sub

Private Sub mnuns_Click()
On Error Resume Next
    ChDir WebBor.Netscape
    tRet = WinExec("Netscape.exe " & TBookURL, 3)
    ChDir Abspath
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
    
End Sub

Private Sub mnuoa_Click()
On Error Resume Next
    ChDir WebBor.Opera
    tRet = WinExec("Opera.exe " & TBookURL, 3)
    ChDir Abspath
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
End Sub

Private Sub mnupasscat_Click()
    frmrename.Show vbModal, frmmain
End Sub

Private Sub mnuping_Click()
    frmping.Show vbModal, frmmain
End Sub

Private Sub mnures_Click()
    frmmain.WindowState = 0
    frmmain.Show
    frmmain.Tray1.Visible = False
    
End Sub

Private Sub mnusavedtop_Click()
    frmshurl.Show vbModal, frmmain
End Sub

Private Sub mnusaveweb_Click()
Dim lzFile As String, FileExt As String
On Error Resume Next
    frmmain.CDialog1.DialogTitle = "Save Web Page"
    frmmain.CDialog1.Filter = "Hyper Text Document(*.html)|*.html|"
    frmmain.CDialog1.ShowSave
    lzFile = Trim(frmmain.CDialog1.FileName)
    If Len(lzFile) = 0 Then Exit Sub

    If Len(GetFileExt(lzFile)) = 0 Then
        lzFile = lzFile & ".html"
        SavewebPage lzFile
        Exit Sub
    Else
        SavewebPage lzFile
    End If
End Sub

Private Sub mnuslnk_Click()
    frmref.Show vbModal, frmmain
End Sub

Private Sub mnuwhois_Click()
On Error Resume Next
    frmmain.RunAddins 1
End Sub
