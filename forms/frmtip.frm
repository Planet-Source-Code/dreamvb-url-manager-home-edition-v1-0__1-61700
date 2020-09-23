VERSION 5.00
Begin VB.Form frmtip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the day"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Flat Flat3 
      Height          =   240
      Left            =   2670
      TabIndex        =   7
      Top             =   2145
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
      Left            =   120
      TabIndex        =   6
      Top             =   2130
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   476
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
      BorderColor     =   -2147483643
   End
   Begin VB.ComboBox cboShowTip 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   2
      Top             =   2115
      Width           =   2835
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close Tips"
      Height          =   365
      Left            =   4290
      TabIndex        =   4
      Top             =   2085
      Width           =   1095
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next Tip"
      Height          =   365
      Left            =   3030
      TabIndex        =   3
      Top             =   2085
      Width           =   1095
   End
   Begin VB.PictureBox picback 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   120
      ScaleHeight     =   1860
      ScaleWidth      =   5265
      TabIndex        =   1
      Top             =   105
      Width           =   5265
      Begin VB.Label lbltipdisplay 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1065
         Left            =   150
         TabIndex        =   0
         Top             =   690
         Width           =   4980
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   -45
         X2              =   2160
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4575
         Picture         =   "frmtip.frx":0000
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   45
         TabIndex        =   5
         Top             =   180
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TipCnt As Long
Private nIndex As Long
Private TipFile As String
Private TipText As String

Private Function ShowTip() As String
Dim RndTip As Integer, iRet As Long

    Randomize TipCnt ' Randomize
    RndTip = Int((TipCnt * Rnd) + 1) ' Get a Random number based on the tip count
    TipText = String(120, Chr$(0))
    iRet = GetPrivateProfileString("Tip" & RndTip, "Tip", "No tip found", TipText, 120, TipFile)
    TipText = Left(TipText, iRet)
    iRet = 0
    
End Function

Private Sub cboShowTip_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cboShowTip_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdclose_Click()
    If cboShowTip.ListIndex = 0 Then
        nIndex = 1
    Else
        nIndex = 0
    End If
    SaveSetting "DMUrlMan", "Config", "ShowTips", CStr(nIndex)
    TipCnt = 0 ' reset the tip counter
    TipText = ""
    Unload frmtip   ' unload the form
End Sub

Private Sub cmdnext_Click()
    If TipCnt = 0 Then Exit Sub
   ' Randomize TipCnt
    ShowTip
    lbltipdisplay.Caption = TipText ' update the tip display with random tip
    TipText = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim iRet As Long
Dim mTotal As String
Dim tFile As Long

    TipFile = FixPath(App.Path) & "Tips.ini"
    frmtip.Icon = Nothing
    FlatBorder picback.hwnd, True
    mTotal = String(4, Chr$(0))
    iRet = GetPrivateProfileString("General", "Total", "ERROR", mTotal, 4, TipFile)
    mTotal = Left(mTotal, iRet)
    TipCnt = Val(mTotal)
    mTotal = ""
    iRet = 0
    ShowTip
    lbltipdisplay.Caption = TipText ' display tip
    cboShowTip.AddItem "Show tips at start-up"
    cboShowTip.AddItem "Never show tips at start-up"
    
    If Val(Config.ShowTips) = 0 Then
        cboShowTip.ListIndex = 1
    Else
        cboShowTip.ListIndex = 0
    End If
    
    
End Sub

Private Sub Form_Resize()
    Line1.X2 = picback.Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmtip = Nothing ' Release the form from memory
End Sub
