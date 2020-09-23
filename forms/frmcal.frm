VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      ShowToday       =   0   'False
      StartOfWeek     =   44498945
      TitleBackColor  =   -2147483646
      TitleForeColor  =   -2147483639
      CurrentDate     =   37877
   End
   Begin VB.Shape Shape1 
      Height          =   2340
      Left            =   0
      Top             =   0
      Width           =   2670
   End
End
Attribute VB_Name = "frmcal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Select Case ModiyDate
        Case LastViewedDate
            MonthView1.Value = TBookLastVist
        Case AddedDate
            MonthView1.Value = TBookAddDate
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcal = Nothing
    
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    
    Select Case ModiyDate
        Case AddedDate
            frmEdit.txtDateAdd.Text = Format(MonthView1.Value, "Medium Date")
        Case LastViewedDate
            frmEdit.txtvisdate.Text = Format(MonthView1.Value, "Medium Date")
    End Select
    
    Unload frmcal
    
End Sub
