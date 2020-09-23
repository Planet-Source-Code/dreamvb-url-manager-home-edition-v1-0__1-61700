Attribute VB_Name = "mod1"
Public SiteID As Long ' Id of the link infor in the database
Public RecoredName As String

Public TBookURL As String
Public TBookMark As String
Public TBookHit As Long
Public TBookMarkDes As String
Public TBookAddDate As Date
Public TBookLastVist As Date

Public TBookSnapShot As String
Public TvIcon As Long
Public TViewed As Integer
Public TRate As Integer

Public DbPath As String

Public Function HasSpace(mStr As String) As Boolean
    ' This code will check the string for a space in the string
    If InStr(1, mStr, Chr(32), vbTextCompare) Then HasSpace = True Else HasSpace = False
End Function

Private Sub OutputImage(lzData As String, imgFile As String)
Dim iFile As Long
    iFile = FreeFile
    Open GetTempFolder & imgFile For Binary As #iFile
        Put #iFile, , lzData
    Close #iFile
End Sub

Public Sub OutputHtml(lzData As String, HtmlFile As String)
Dim iFile As Long
    iFile = FreeFile
    Open GetTempFolder & HtmlFile For Binary As #iFile
        Put #iFile, , lzData
    Close #iFile

End Sub

Public Function GetImage(ImgNum As Long) As String
Dim sData As String, StrImgSrc As String

    Select Case ImgNum
        Case 0
            Exit Function
        Case 1
             sData = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
             StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "16" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
             OutputImage sData, "tmpImg.gif"
        Case 2
            sData = StrConv(LoadResData(102, "CUSTOM"), vbUnicode)
            StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "33" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
            OutputImage sData, "tmpImg.gif"
        Case 3
            sData = StrConv(LoadResData(103, "CUSTOM"), vbUnicode)
            StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "50" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
            OutputImage sData, "tmpImg.gif"
        Case 4
            sData = StrConv(LoadResData(104, "CUSTOM"), vbUnicode)
            StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "67" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
            OutputImage sData, "tmpImg.gif"
        Case 5
            sData = StrConv(LoadResData(105, "CUSTOM"), vbUnicode)
            StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "84" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
            OutputImage sData, "tmpImg.gif"
        Case Is >= 5
            sData = StrConv(LoadResData(105, "CUSTOM"), vbUnicode)
            StrImgSrc = "<img src=" & Chr(34) & "tmpImg.gif" & Chr(34) & " width=" & Chr(34) & "84" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & " align=" & Chr(34) & "absmiddle" & Chr(34) & "> "
            OutputImage sData, "tmpImg.gif"
    End Select
    
    sData = ""
    GetImage = StrImgSrc
    StrImgSrc = ""
    
End Function

Public Function DoubleCRLF() As String
    DoubleCRLF = vbNewLine & vbNewLine
End Function
Public Function GetFileExt(lzFile As String) As String
Dim I As Long, iPart As Long, StrA As String
   For I = Len(lzFile) To 1 Step -1
        StrA = Mid(lzFile, I, 1)
        If StrA = "." Then
            iPart = I
            Exit For
        End If
   Next
   
   If iPart = 0 Then
        GetFileExt = ""
    Else
        GetFileExt = UCase$(Mid$(lzFile, iPart + 1, Len(lzFile)))
   End If
   iPart = 0: I = 0
   StrA = ""
   
End Function
Public Function ShortPath(lzPath As String) As String
Dim iRet As Long
Dim StrBuff As String
    StrBuff = String(255, Chr$(0))
    iRet = GetShortPathName(lzPath, StrBuff, 255)
    ShortPath = Left$(StrBuff, iRet)
    StrBuff = ""
End Function
Public Function CopyCommand(mCommand As CPYOP, mStr As String)
    
    Clipboard.Clear ' Clear the content of the clipboard
    ' Used to copy bookmark information
    Select Case mCommand
        Case m_SiteName
            Clipboard.SetText mStr
        Case m_SiteURL
            Clipboard.SetText mStr
        Case m_SiteAddedDate
            Clipboard.SetText mStr
        Case m_SiteLastVistDate
            Clipboard.SetText mStr
        Case m_SiteHits
            Clipboard.SetText mStr
        Case m_HtmlCode
            Clipboard.SetText mStr
    End Select
    
End Function

Public Function FixPath(lzPath As String) As String
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If MakeControlFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Function

Function GetTempFolder() As String
Dim iRet As Long, lzStr As String

    lzStr = String(512, vbNull)
    iRet = GetTempPath(512, lzStr)
    
    If iRet = 0 Then
        GetTempFolder = ""
        Exit Function
    Else
        lzStr = Left(lzStr, iRet)
        GetTempFolder = lzStr
        lzStr = ""
        iRet = 0
    End If
    
End Function


Function MakeFlatControls(frm As Form)
Dim Icnt As Long
    ' Returns long 32bit hangle of each control found for the flatborder function
    For Icnt = 0 To frm.Controls.Count - 1
        Select Case TypeName(frm.Controls(Icnt))
            Case "ListView", "DirListBox", "ListBox", "TextBox", "ListView"
                FlatBorder frm.Controls(Icnt).hwnd, True ' applys flatborder to each control found
        End Select
    Next Icnt
    Icnt = 0
    
End Function
Public Function FindFile(lzFile As String) As Boolean
    If Dir$(lzFile) = "" Then FindFile = False Else FindFile = True
    
End Function


Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer

    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If
    
    
End Function

Function RemoveTvItem(ItemIndex As Long, TreeV As TreeView)
On Error Resume Next
    TreeV.Nodes.Remove ItemIndex
    TreeV.Refresh
    If Err Then Err.Clear
    
End Function
Function RemoveLVItem(ItemIndex As Long, lstView As ListView)
On Error Resume Next
    lstView.ListItems.Remove ItemIndex ' Remove selected index
    lstView.ListItems.Clear ' Clear list items
    lstView.Refresh
    If Err Then Err.Clear
    
End Function
Public Function DMGetSpecialFolderLocation(TFolder As TSpecialFolders) As String
    Dim StrBuff As String
    Dim RetVal As Long
    Dim IDL As ITEMIDLIST
    RetVal = SHGetSpecialFolderLocation(100, TFolder, IDL)
    If RetVal = NOERROR Then
        StrBuff = String$(512, Chr(0))
        RetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal StrBuff)
        DMGetSpecialFolderLocation = Left$(StrBuff, InStr(StrBuff, Chr(0)) - 1)
        Exit Function
    End If
    DMGetSpecialFolderLocation = ""
    StrBuff = ""
    
End Function
Public Function SaveURLShortCut(Urlname As String, UrlLink As String)
Dim ans
    If Not FindFile(FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url") = True Then
        WritePrivateProfileString "InternetShortcut", "URL", Urlname, FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
        WritePrivateProfileString "InternetShortcut", "IconIndex", "3", FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
        WritePrivateProfileString "InternetShortcut", "IconFile", ShortPath(FixPath(App.Path) & App.EXEName & ".exe"), FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
        MsgBox "The shortcut has now been successfully placed on your desktop", vbInformation, frmmain.Caption
        Exit Function
    Else
        ans = MsgBox(FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url" & _
        vbCrLf & "already exists Do you want to replace this item?", vbYesNo Or vbQuestion)
        If ans = vbNo Then
            Exit Function
        Else
            WritePrivateProfileString "InternetShortcut", "URL", Urlname, FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
            WritePrivateProfileString "InternetShortcut", "IconIndex", "3", FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
            WritePrivateProfileString "InternetShortcut", "IconFile", ShortPath(FixPath(App.Path) & App.EXEName & ".exe"), FixPath(DMGetSpecialFolderLocation(DM_DESKTOP)) & UrlLink & ".url"
        End If
    End If
    MsgBox "The shortcut has been successfully saved to your desktop", vbInformation, frmmain.Caption
        
End Function

Function SetLstVFlScroolBar(lstView As ListView)
Dim I_Ret As Long
   SendMessageByLong lstView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FLATSB, True
   
End Function

Function PhaseDomain(lzUrl As String) As String
Dim iPart As Long, lpart As Long, StrA As String
On Error Resume Next

    iPart = InStr(lzUrl, "http://")
    If iPart = 1 Then lzUrl = Mid(lzUrl, iPart + 7, Len(lzUrl) - iPart)
    iPart = InStr(lzUrl, ".")
    If iPart = 0 Then PhaseDomain = lzUrl: Exit Function
    StrA = Mid$(lzUrl, iPart + 1, Len(lzUrl) - iPart)
    lpart = InStr(StrA, "/")
    If lpart = 0 Then PhaseDomain = StrA: Exit Function
    StrA = Mid$(StrA, 1, lpart - 1)
    PhaseDomain = StrA
    iPart = 0: lpart = 0: StrA = ""
    
End Function
Public Function SavewebPage(lzFilename As String)
Dim RetVal As Long
    RetVal = URLDownloadToFile(0, TBookURL, lzFilename, 0, 0)
    
    If RetVal < 0 Then
        MsgBox "There was an error while carrying out your request." & DoubleCRLF _
        & "This may be due to:" _
        & DoubleCRLF & "Their maybe a problem on the server" & DoubleCRLF _
        & "The link does not exists please try again latter.", vbCritical
        Exit Function
    Else
        MsgBox "The Bookmark has now been successfully saved.", vbInformation, frmmain.Caption
    End If
    
    RetVal = 0
    
End Function

Function OpenFile(lzFile As String) As String
Dim iFile As Long, StrBuff As String

    iFile = FreeFile
    Open lzFile For Binary As #iFile
        StrBuff = Space$(LOF(iFile))
        Get #iFile, , StrBuff
    Close #iFile
    OpenFile = StrBuff
    StrBuff = ""
    
End Function
Function IsEmail(sCheckEmail) As Boolean
    Dim sEmail, nAtLoc
    IsEmail = True
    sEmail = Trim(sCheckEmail)
    nAtLoc = InStr(1, sEmail, "@")

    If Not (nAtLoc > 1 And (InStrRev(sEmail, ".") > nAtLoc + 1)) Then
        IsEmail = False
    ElseIf InStr(nAtLoc + 1, sEmail, "@") > nAtLoc Then
        IsEmail = False
    ElseIf Mid(sEmail, nAtLoc + 1, 1) = "." Then
        IsEmail = False
    ElseIf InStr(1, Right(sEmail, 2), ".") > 0 Then
        IsEmail = False
    End If
End Function

Public Function SHWait(ByVal ProgID As Long) As Boolean
Dim mExitID As Long, hdlProg As Long
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgID)
    GetExitCodeProcess hdlProg, mExitID
    Do While mExitID = STILL_ACTIVE
        DoEvents
        GetExitCodeProcess hdlProg, mExitID
    Loop
    CloseHandle hdlProg
    SHWait = mExitID
End Function

Public Function PingHost(sHost As String) As String
Dim RetVal As Long, iFile As Long, ExecuteCmd As String
    On Error Resume Next
    ExecuteCmd = "ping " & sHost & ">" & FixPath(ShortPath(App.Path)) & "ping.dat"
    
    RetVal = Shell("command.com /c" & ExecuteCmd, vbHide)

    If Not SHWait(RetVal) Then
        PingHost = OpenFile(FixPath(ShortPath(App.Path)) & "ping.dat")
    End If
    
    Kill FixPath(ShortPath(App.Path)) & "ping.dat"

    
End Function

'
Public Function ReadServConfig()
'On Error Resume Next
Dim AppName As String, SubItem As String
    AppName = "DMUrlMan"
    SubItem = "Config"
    
    Config.FirstRun = GetSetting(AppName, SubItem, "FirstRun", "")
    Config.mSMTP_serv = GetSetting(AppName, SubItem, "Smtpserver", "mail.yourISP.com")
    Config.mWHOIS_serv = GetSetting(AppName, SubItem, "Whois", "whois.internic.net")
    Config.ShowTips = GetSetting(AppName, SubItem, "ShowTips", "1")
    Config.defBrowser = GetSetting(AppName, SubItem, "Browser", "0")
    Config.Hightlight = GetSetting(AppName, SubItem, "Highlight", "16772313")
    Config.NewItems = GetSetting(AppName, SubItem, "NewItems", "255")
    Config.FavIcon = GetSetting(AppName, SubItem, "TvFavIcon", "1")
    Config.mOpenItems = GetSetting(AppName, SubItem, "OpenItems", "2")
    Config.mbkDatabase = GetSetting(AppName, SubItem, "Db", FixPath(App.Path) & "Bookmarks.mdb")
    Config.dmWebSite = GetSetting(AppName, SubItem, "WebURL", "http://www.eraystudios.co.uk")
    Config.FavIdx = GetSetting(AppName, SubItem, "TvFavIconIdx", "1")
    Config.dmView = GetSetting(AppName, SubItem, "View", "3")
    Config.dmViewCat = GetSetting(AppName, SubItem, "dmViewCats", "True")
    Config.dmViewDesWnd = GetSetting(AppName, SubItem, "dmViewDesWnd", "True")
    Config.dmViewTime = GetSetting(AppName, SubItem, "dmViewTime", "True")
    Config.dmViewWebView = GetSetting(AppName, SubItem, "dmViewWebView", "True")
    Config.dmLastView = GetSetting(AppName, SubItem, "dmLastView", "1")
    Config.ProgRegister.mRegName = GetSetting(AppName, "Register", "RegName", "Unregistered Name")
    Config.ProgRegister.mRegCompany = GetSetting(AppName, "Register", "Company", "Unregistered Company")
    Config.ProgRegister.mRegKey = GetSetting(AppName, "Register", "Key", "0000-0000-0000-0000")
    
    If CBool(Val(Config.FirstRun)) = False Then
         SaveSetting AppName, SubItem, "FirstRun", "1"
         SaveSetting "DMUrlMan", "Config", "AppPath", FixPath(App.Path)
    End If
    
End Function

Private Function PhaseData(lData As String)
Dim ipOS As Long
    ipOS = InStr(1, lData, ":")
    DMBook.BookName = Mid(lData, 1, ipOS - 1)
    DMBook.BookUrl = Mid(lData, ipOS + 1, Len(lData))
End Function
Public Function TOpenSite(lHangle As Long, lzCommand As String) As Long
    OpenSite = ShellExecute(frmmain.hwnd, vbNullString, lzCommand, vbNullString, vbNullString, 0)
End Function
Public Function RemoveSelection(LstV As ListView)
Dim I As Long

    For I = 1 To LstV.ListItems.Count
        If LstV.ListItems(I).Selected Then
            LstV.ListItems(I).Selected = False
            Exit For
        End If
    Next
    
End Function

Public Function MoveFrmToPos(frm As Form, X_Pos As Long, Y_Pos As Long)
Dim T_Point As POINTAPI
    GetCursorPos T_Point
    frm.Move (T_Point.x + X_Pos) * Screen.TwipsPerPixelX, (T_Point.y + Y_Pos) * Screen.TwipsPerPixelY
    frm.Show vbModal
    
End Function

Public Function LoadImgLst(picsrc As PictureBox, picdst As PictureBox, ImgList As ImageList)
    Dim frmCnt As Long, I As Long
    
    picsrc.Picture = LoadResPicture(102, vbResBitmap)
    frmCnt = (picsrc.Width / picsrc.Height) ' get the number of images in the strip

    For I = 1 To frmCnt
        BitBlt picdst.hDC, 0, 0, 16, 16, picsrc.hDC, I * 16 - 16, 0, vbSrcCopy
        ImgList.ListImages.Add I, "a" & I, picdst.Image
        picdst.Refresh
    Next
    
    frmCnt = 0
    I = 0
    Set picsrc.Picture = Nothing
    Set picdst.Picture = Nothing
    
End Function

Public Function LoadTvImg(picsrc As PictureBox, picdst As PictureBox, ImgList As ImageList)
    Dim frmCnt As Long, I As Long
    
    picsrc.Picture = LoadResPicture(103, vbResBitmap)
    frmCnt = 11

    For I = 1 To frmCnt
        BitBlt picdst.hDC, 0, 0, 16, 16, picsrc.hDC, I * 16 - 16, 0, vbSrcCopy
        ImgList.ListImages.Add I, "a" & I, picdst.Image
        picdst.Refresh
    Next
    
    frmCnt = 0
    I = 0
    Set picsrc.Picture = Nothing
    Set picdst.Picture = Nothing

End Function

Public Function ResizeLstHeader(LstV As ListView, ColHeaderInx As Long)
    SendMessage LstV.hwnd, LVM_SETCOLUMNWIDTH, ColHeaderInx, LVSCW_AUTOSIZE_USEHEADER
    
End Function
