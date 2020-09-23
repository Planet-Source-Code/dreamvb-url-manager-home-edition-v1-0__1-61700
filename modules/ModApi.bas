Attribute VB_Name = "ModApi"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

'Public Const GWL_WNDPROC = -4
'Public Const DM_ADDBOOK = &H206

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const STATUS_PENDING = &H103
Public Const STILL_ACTIVE = STATUS_PENDING
Public Const MF_BYPOSITION = &H400&
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const HDS_BUTTONS = &H2

Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203

' used for the listview control
Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVS_EX_FLATSB = &H100
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEMA = (HDM_FIRST + 4)
Public Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2

Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_IMAGE = &H800
Public Const HDF_STRING = &H4000

Public Const HDM_SETITEM = (HDM_FIRST + 4)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Public Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

' Browse Folder consts
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_NEWDIALOGSTYLE = &H40

' Window consts
Public Const GWL_HINSTANCE = (-6)
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

' Web Broswer Types
Public Const IE = 0
Public Const Netscape = 1
Public Const Opera = 2
Public Const Mozilla = 3
Public Const FireFox = 4

' Modify Date Consts
Public ModiyDate As ModDate

Public M_Colour As Long, LstIndex As Long, TvIndex As Long, plgIni As String

'Public DMBook As IE_URLS
Public Config As DMURL_cfg
Public WebBor As BROSWER

Public Abspath As String
Public WebExe As String
Public SerachLst As String ' Serach list config file
Public mSerachPattern As String ' The text you want to serach for
Public SerachOption As Integer ' The serach engine your using
Public FromWeb As Boolean
Public SiteFound As Boolean
Public DbError As Boolean

Enum ModDate
    LastViewedDate = 1
    AddedDate = 2
End Enum

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Type POINTAPI
   x As Long
   y As Long
End Type

Type HITTESTINFO
   pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem  As Long
End Type

Type SHITEMID
    cb As Long
    abID As Byte
End Type

Type ITEMIDLIST
    mkid As SHITEMID
End Type

Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   iImage As Long
   iOrder As Long
End Type

Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'Type IE_URLS
 '   BookName As String
  '  BookUrl As String
'End Type

Enum CPYOP
    m_SiteName
    m_SiteURL
    m_SiteAddedDate
    m_SiteLastVistDate
    m_SiteHits
    m_HtmlCode
End Enum

Type RegUser
    mRegName As String
    mRegCompany As String
    mRegKey As String
End Type

Type DMURL_cfg
    FirstRun As String
    mSMTP_serv As String
    mWHOIS_serv As String
    mbkDatabase As String
    Hightlight As String
    NewItems As String
    FavIcon As String
    FavIdx As String
    defBrowser As String
    mOpenItems As String
    ShowTips As String
    dmWebSite As String
    dmWebRegister As String
    dmView As String
    dmViewCat As String
    dmViewDesWnd As String
    dmViewTime As String
    dmViewWebView As String
    dmLastView As String
    ProgRegister As RegUser
End Type

Type BROSWER
    IE As String
    Netscape As String
    Opera As String
    Mozilla As String
    FireFox As String
End Type

Public Enum TSpecialFolders
    DM_DESKTOP = &H0
    DM_PROGRAMS = &H2
    DM_Controls = &H3
    DM_PRINTERS = &H4
    DM_PERSONAL = &H5
    DM_FAVORITES = &H6
    DM_STARTUP = &H7
    DM_RECENT = &H8
    DM_SENDTO = &H9
    DM_BITBUCKET = &HA
    DM_STARTMENU = &HB
    DM_DESKTOPDIRECTORY = &H10
    DM_DRIVES = &H11
    DM_NETWORK = &H12
    DM_NETHOOD = &H13
    DM_FONTS = &H14
    DM_TEMPLATES = &H15
End Enum

