VERSION 5.00
Begin VB.Form frmgenhtm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbocat 
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Top             =   600
      Width           =   3930
   End
   Begin Project1.Flat2 Flat21 
      Height          =   1845
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3254
      BackStyle       =   0
   End
   Begin VB.CommandButton cmdfolder 
      Caption         =   "...."
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      Top             =   1365
      Width           =   375
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3255
      TabIndex        =   5
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Generate"
      Height          =   350
      Left            =   1875
      TabIndex        =   3
      Top             =   2130
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      Height          =   300
      Left            =   285
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generate for:"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   330
      Width           =   930
   End
   Begin VB.Label lblsavepath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save generated web page to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   315
      TabIndex        =   4
      Top             =   1110
      Width           =   2550
   End
End
Attribute VB_Name = "frmgenhtm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TmpFol As String
Private GenOption As String

Private Function WriteFile(strdata As String, sFile As String)
Dim mFile As Long
    mFile = FreeFile
    Open sFile For Binary As #mFile
        Put #mFile, , strdata
    Close #mFile
    strdata = ""
     
End Function

Private Sub cbocat_Click()
    If cbocat.ListIndex = 0 Then
         GenOption = "ALL"
    Else
        GenOption = "SELECTED"
    End If

    RecoredName = cbocat.Text
    End Sub

Private Sub cbocat_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 ' Disable the text from chnageing
End Sub

Private Sub cmdcancel_Click()
    cbocat.Clear        ' Clear the combo box
    txtname.Text = ""   ' Clear the text box
    TmpFol = ""         ' Clear temp folder
    GenOption = ""      ' Clear the option buffer
    Unload frmgenhtm ' unload the form
End Sub

Private Sub cmdfolder_Click()
Dim FolName As String
    FolName = GetFolder(frmgenhtm.hwnd, "Choose location:")
    If Len(FolName) <= 0 Then
        txtname.Text = TmpFol
        txtname.ToolTipText = TmpFol
        Exit Sub
    Else
        txtname.Text = FixPath(FolName)
        txtname.ToolTipText = txtname.Text
    End If
    FolName = ""
End Sub

Private Sub cmdok_Click()
Dim StrBuff As String, StrA As String, I As Long

    StrA = StrConv(LoadResData(107, "CUSTOM"), vbUnicode)
    
    If GenOption = "SELECTED" Then
        StrBuff = StrA & GenHtmlPage(RecoredName) & vbCrLf & "</body>" & vbCrLf & "</html>" _
        & vbCrLf & "<!--bookmarks page created by " & frmmain.Caption & "--!>"
        WriteFile StrBuff, txtname.Text & "bookmarks.html"
        StrA = ""
        StrBuff = ""
        MsgBox "Your bookmarks have now been saved to: " & DoubleCRLF & txtname.Text & "bookmarks.html", vbInformation, frmgenhtm.Caption
        cmdcancel_Click ' call cancel button to unload the form
        Exit Sub
    Else
        For I = 2 To frmmain.tv1.Nodes.Count
            StrBuff = StrBuff & GenHtmlPage(frmmain.tv1.Nodes(I).Text)
        Next
        
        I = 0
        
        StrA = StrA & StrBuff & vbCrLf & "</body>" & vbCrLf & "</html>" _
        & vbCrLf & "<!--bookmarks page created by " & frmmain.Caption & "--!>"
        WriteFile StrA, txtname.Text & "bookmarks.html"
        MsgBox "Your bookmarks have now been saved to: " & DoubleCRLF & txtname.Text & "bookmarks.html", vbInformation, frmgenhtm.Caption
        StrA = ""
        StrBuff = ""
        cmdcancel_Click ' call cancel button to unload the form
    End If

End Sub

Private Sub Form_Load()
Dim catIndex As Long, I As Long

    frmgenhtm.Icon = Nothing    ' Remove the forms icon
    TmpFol = FixPath(App.Path)  ' assign the temp path for bookmark file
    txtname.Text = TmpFol       ' assign the path to the text box
    txtname.ToolTipText = TmpFol    ' Assign the tooltip text the tmp path
    FlatBorder txtname.hwnd, True   ' Add flat border around textbox
    
    
    Initcbocat cbocat ' fill combo box with table names in db
    
    
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
    
    frmgenhtm.Caption = "Generate web page for " & RecoredName ' update the caption
    catIndex = 0 ' clean var
    I = 0   ' clean var

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmgenhtm = Nothing ' unload form from memory
End Sub

Private Sub txtname_GotFocus()
    txtname.BackColor = Config.Hightlight ' update textbox highlight colour
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
    ' This little bit of code stops the beep in a textbox when enter is pressed
    If KeyAscii = 13 Then
        cmdok_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtname_LostFocus()
    txtname.BackColor = vbWhite ' ' update textbox highlight colour
End Sub
