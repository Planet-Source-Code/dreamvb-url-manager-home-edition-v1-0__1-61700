Attribute VB_Name = "Module1"
Private hHook As Long
Public ButtonIndex As Long
Public ButtonCaption(0 To 10) As String

Public Sub SubClass(frm As Form)
    Dim hInst As Long
    Dim Thread As Long
    
    hInst = GetWindowLong(frm.hwnd, GWL_HINSTANCE) ' Get the windows hangle
    Thread = GetCurrentThreadId()   ' Get the current thread of the window
    hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread) ' Set the hook on the window
    
End Sub

Public Function Manipulate(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iRet As Long, I As Long, cName As String, clsName As String
Dim TButton(0 To 8) As Long

    If lMsg = HCBT_ACTIVATE Then ' Active Window Found
   
        clsName = Space(20) ' Create a buffer
        For I = 1 To 8
            TButton(I) = FindWindowEx(wParam, TButton(I - 1), vbNullString, vbNullString) ' Find the HWND of each window
            iRet = GetClassName(TButton(I), clsName, Len(clsName)) ' Get the windows class name
            If TButton(I) = 0 Then Exit For ' Exit if we have no more windows
            
            cName = UCase(Left(clsName, iRet)) ' Trim down the window name
            
            Select Case cName
                Case "BUTTON" ' Button was found
                    SetWindowText TButton(ButtonIndex), ButtonCaption(ButtonIndex)
                    ' the code above updates the buttons caption based on it's index
            End Select
        Next
        'Destroy the hook and clean up
        UnhookWindowsHookEx hHook
        I = 0
        iRet = 0
        clsName = ""
        cName = ""
        Erase TButton
        Erase ButtonCaption
    End If
    
End Function


