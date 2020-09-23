Attribute VB_Name = "modmain"
Enum RegOp
    Register = 1
    UnRegister
End Enum

Private Type Plugins
    mPlugName As String
    mPlugClass As String
    mPlugDLL As String
End Type

Public mAddins() As Plugins

Public Function GetWindowsRoot() As String
Dim iRet As Long, StrB As String
    StrB = String(164, Chr$(0))
    iRet = GetWindowsDirectory(StrB, 164)
    GetWindowsRoot = Left(StrB, iRet)
    StrB = ""
    iRet = 0
    
End Function

Public Function IsAppOpen() As Boolean
    If App.PrevInstance Then
        IsAppOpen = True
    Else
        IsAppOpen = False
    End If
    
End Function


Public Function RegisterActiveX(lzAxDll As String, mRegOption As RegOp) As Boolean
Dim mLib As Long, DllProcAddress As Long
Dim mThread
Dim sWait As Long
Dim mExitCode As Long
Dim lpThreadID As Long

Dim slib As String

    slib = lzAxDll
    mLib = LoadLibrary(slib)
    
    If mLib <= 0 Then
        RegisterActiveX = False
        Exit Function
    End If
    
    If mRegOption = Register Then
        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
    Else
        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
    End If
    
    If DllProcAddress = 0 Then
        RegisterActiveX = True
        Exit Function
    Else
        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
        
        If mThread = 0 Then
            FreeLibrary mLib
            RegisterActiveX = False
            Exit Function
        Else
            sWait = WaitForSingleObject(mThread, 10000)
            If sWait <> 0 Then
                FreeLibrary lLib
                mExitCode = GetExitCodeThread(mThread, mExitCode)
                ExitThread mExitCode
                Exit Function
            Else
                FreeLibrary mLib
                CloseHandle mThread
            End If
        End If
    End If
    slib = ""
    RegisterActiveX = True
    
End Function


Public Sub LoadSerchList()
Dim iRet As Long, mTotal As Integer, Icnt As String * 2, StrB As String
    If FindFile(SerachLst) = False Then Exit Sub
    Icnt = String(4, Chr$(0)) ' Create a buffer
    iRet = GetPrivateProfileString("General", "Total", "", Icnt, 4, SerachLst) ' Read in ini settings
    
    mTotal = Val(Left(Icnt, iRet)) ' Get the number of serach engines
    StrB = String(128, Chr(0)) ' Create a buffer for the serach engine captions
    
    For I = 1 To mTotal
        iRet = GetPrivateProfileString("Serach" & I, "MenuCaption", "", StrB, 128, SerachLst) ' Read in ini settings
        Load frmPopUpmenu.MnuName(I) ' Load in new menus for the serach engine captions
        frmPopUpmenu.MnuName(I).Checked = False
        frmPopUpmenu.MnuName(0).Visible = False ' Hide the first menu item
        frmPopUpmenu.MnuName(I).Caption = Left(StrB, iRet) ' Update menus caption with serach engines name
        frmPopUpmenu.MnuName(I).Visible = True ' Show all the menus
        frmPopUpmenu.MnuName(1).Checked = True
    Next
    ' Clear up vars
    StrB = ""
    Icnt = ""
    mTotal = 0
    I = 0

End Sub


Public Function MakeShPath(sPath As String) As String
Dim iRet As Long, sBuff As String
    sBuff = String(164, Chr$(0))
    iRet = GetShortPathName(sPath, sBuff, 164)
    MakeShPath = Left(sBuff, iRet)
    
End Function

Public Sub ShowHeaderIcon(LstV As ListView, colNo As Long, IconIdx As Long, ShowColIcon As Boolean)
   Dim hHeader As Long
   Dim LstHd As HD_ITEM
   
   hHeader = SendMessage(LstV.hwnd, LVM_GETHEADER, 0&, ByVal 0&)
   
   With LstHd
      .mask = HDI_IMAGE Or HDI_FORMAT
      .pszText = LstV.ColumnHeaders(colNo + 1).Text
      
       If ShowColIcon Then
         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
         .iImage = IconIdx
       Else
         .fmt = HDF_STRING
      End If
      
   End With
   
   SendMessage hHeader, HDM_SETITEM, colNo, LstHd
   
End Sub

Function En(s As String) As String
Dim I As Long, p As String, ch As Long

    For I = 1 To Len(s)
        ch = Asc(Mid(s, I, 1)) Xor 5
        p = p & Chr(ch)
    Next
    
    En = p
    
End Function
