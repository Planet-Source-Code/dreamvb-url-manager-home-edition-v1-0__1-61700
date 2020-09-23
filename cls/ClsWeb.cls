VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETTEXT = &HD

Private WBOpen As Boolean

Private Type IEURL
    mTitle As String
    mUrl As String
End Type

Private shIE As IEURL


Public Sub GetWBInfo()
Dim IeWnd As Long, wndC(10) As Long, I As Long, clsName As String, cName As String
Dim WorkerW As Long, iRet As Long, T As Long, TToolbar32 As Long, ComboBoxEx32 As Long, _
ComboB As Long, TEdit As Long, TEditLen As Long, TEditTxt As String, WndName As String

    WBOpen = False
    IeWnd = FindWindow("IEFrame", vbNullString) ' Find the current ie open window
    
    If IeWnd <= 0 Then WBOpen = False: Exit Sub ' exit function if wnd not found
    
    iRet = GetWindowTextLength(IeWnd) ' get the len of wnd caption
    cName = Space(iRet) ' create a buffer
    GetWindowText IeWnd, cName, iRet + 1 ' get the window caption
    
    shIE.mTitle = cName ' Store the windows caption
    cName = "" ' clear the buffer
    
    WorkerW = FindWindowEx(IeWnd, 0, "WorkerW", vbNullString)
    TToolbar32 = FindWindowEx(WorkerW, 0, "ReBarWindow32", vbNullString)
    ComboBoxEx32 = FindWindowEx(TToolbar32, 0, "ComboBoxEx32", vbNullString)
    ComboB = FindWindowEx(ComboBoxEx32, 0, "ComboBox", vbNullString)
    TEdit = FindWindowEx(ComboB, 0, "Edit", vbNullString)
    TEditLen = SendMessage(TEdit, WM_GETTEXTLENGTH, 0, 0) + 1
    TEditTxt = Space(TEditLen) ' create a buffer

    T = SendMessage(TEdit, WM_GETTEXT, ByVal TEditLen, ByVal TEditTxt)
    ' The code above sends a message to the wnd to get the current url text
    TEditTxt = Left(TEditTxt, Len(TEditTxt) - 1) ' trim the buffer
    
    T = 0
    iRet = 0
    
    shIE.mUrl = TEditTxt
    WBOpen = True
    
End Sub

Public Property Get IsWeBOpen() As Boolean
    IsWeBOpen = WBOpen
End Property

Public Property Get Title() As String
    Title = shIE.mTitle
End Property

Public Property Get URL() As String
    URL = shIE.mUrl
End Property


