VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "URL Manager - Home Edition v1.0"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10815
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Line3D Line3D3 
      Height          =   30
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   53
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1905
      Top             =   5175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   7
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0CCA
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0D94
            Key             =   "B"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarserach 
      Height          =   330
      Left            =   9765
      TabIndex        =   18
      Top             =   465
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarfind 
      Height          =   330
      Left            =   8880
      TabIndex        =   17
      Top             =   465
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      ButtonWidth     =   1429
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBarWebBro 
      Height          =   330
      Left            =   150
      TabIndex        =   16
      Top             =   480
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "WEB_IE"
            Object.ToolTipText     =   "Open Bookmark with Internet Explorer"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "WEB_NS"
            Object.ToolTipText     =   "Open Bookmark with Netscape"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "WEB_OP"
            Object.ToolTipText     =   "Open Bookmark with Opera"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "WEB_MZ"
            Object.ToolTipText     =   "Open Bookmark with Mozilla"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "WEB_FOX"
            Object.ToolTipText     =   "Open Bookmark with Mozilla Firefox"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtserach 
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
      Left            =   1980
      TabIndex        =   15
      Tag             =   "SR"
      Text            =   "Enter serach patten here."
      Top             =   480
      Width           =   6765
   End
   Begin VB.PictureBox picdst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1245
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      Top             =   5265
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picsrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1230
      ScaleHeight     =   210
      ScaleWidth      =   120
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4740
      Top             =   5520
   End
   Begin VB.PictureBox picshot 
      BackColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   495
      ScaleHeight     =   1740
      ScaleWidth      =   2505
      TabIndex        =   11
      Top             =   2625
      Width           =   2565
      Begin VB.PictureBox picsh 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   30
         ScaleHeight     =   945
         ScaleWidth      =   1185
         TabIndex        =   12
         Top             =   30
         Width           =   1185
      End
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   53
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   4125
      Top             =   5175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebB 
      Height          =   1530
      Left            =   4470
      TabIndex        =   7
      Top             =   900
      Width           =   885
      ExtentX         =   1561
      ExtentY         =   2699
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox picshbut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   105
      Picture         =   "frmmain.frx":0E5E
      ScaleHeight     =   915
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   990
      Width           =   285
   End
   Begin Project1.Flat2 sidebar 
      Height          =   1665
      Left            =   30
      TabIndex        =   5
      Top             =   915
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   2937
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   495
      TabIndex        =   4
      Top             =   825
      Width           =   2550
      Begin MSComctlLib.Toolbar CatBar 
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   150
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "NEW_CAT"
               Object.ToolTipText     =   "Add new category"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "DEL_CAT"
               Object.ToolTipText     =   "Remove category"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "RENAME_CAT"
               Object.ToolTipText     =   "Rename category"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "GEN_CAT"
               Object.ToolTipText     =   "Generate Web Page"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FIND_CAT"
               Object.ToolTipText     =   "Find category"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "HIDE_CAT"
               Object.ToolTipText     =   "Hide category bar"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   330
         Index           =   0
         Left            =   45
         Picture         =   "frmmain.frx":1CEC
         Top             =   135
         Width           =   90
      End
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   360
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tBar1 
      Height          =   330
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NEW_BOOK"
                  Text            =   "New Bookmark"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NEW_DB"
                  Text            =   "New Database"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN_BOOK"
            Object.ToolTipText     =   "Open Exsiting Database"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BK_DB"
            Object.ToolTipText     =   "Backup Database"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RES_BK"
            Object.ToolTipText     =   "Restore Backup"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EDIT_BOOK"
            Object.ToolTipText     =   "Modify Bookmark"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DEL_BOOK"
            Object.ToolTipText     =   "Delete Bookmark"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOOK_NAME"
                  Text            =   "Bookmark Name"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOOK_URL"
                  Text            =   "Bookmark Location"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOOK_ADD_DATE"
                  Text            =   "Added Date"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnub1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOOK_LAST_VIEW"
                  Text            =   "Last Viewed"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOOK_HIT"
                  Text            =   "Bookmark Hits"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "COPY_HTML"
                  Text            =   "Bookmark HTML Code"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "GET_BOOK"
            Object.ToolTipText     =   "Paste URL From IE"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MOVE_BOOK"
            Object.ToolTipText     =   "Move Bookmark"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND_BOOK"
            Object.ToolTipText     =   "Find Bookmark"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BOOK_STATE"
            Object.ToolTipText     =   "Check Bookmark Status"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAV_WEB"
            Object.ToolTipText     =   "Save Web Page"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SH_LINK"
            Object.ToolTipText     =   "Share Bookmark"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WHO_IS"
            Object.ToolTipText     =   "Whois Lookup"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SH_CUT"
            Object.ToolTipText     =   "Create Shortcut"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bookmarks View"
            ImageIndex      =   27
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MNU_REP"
                  Text            =   "Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MNU_ICO"
                  Text            =   "Icon"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VIEW_STAT"
            Object.ToolTipText     =   "Bookmark Viewed Status"
            ImageIndex      =   30
            Style           =   1
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CFG_NOW"
            Object.ToolTipText     =   "Config"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ABOUT"
            Object.ToolTipText     =   "About this program"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HELP"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   33
         EndProperty
      EndProperty
   End
   Begin Project1.Tray Tray1 
      Left            =   945
      Top             =   5190
      _ExtentX        =   529
      _ExtentY        =   529
      Icon            =   "frmmain.frx":1EE6
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3450
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3264
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":35B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3908
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":42FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4650
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":49A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5046
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5398
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":56EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":60E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6432
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6784
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":717A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":74CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":781E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8214
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8566
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":88B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":89CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9180
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":94D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstsites 
      Height          =   1620
      Left            =   3075
      TabIndex        =   2
      Top             =   900
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   2858
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList4"
      SmallIcons      =   "ImageList4"
      ColHdrIcons     =   "ImageList3"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   1155
      Left            =   510
      TabIndex        =   1
      Top             =   1395
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   2037
      _Version        =   393217
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2700
      Top             =   5295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13900
            MinWidth        =   9
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   53
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   2
      Left            =   15
      Picture         =   "frmmain.frx":9824
      Top             =   480
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   30
      Picture         =   "frmmain.frx":9A1E
      Top             =   60
      Width           =   90
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New Database"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnublank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnubak 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnurestore 
         Caption         =   "&Restore Backup"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucat 
         Caption         =   "&Categorys"
         Begin VB.Menu mnunewcat 
            Caption         =   "&New Category"
         End
         Begin VB.Menu mnudelcat 
            Caption         =   "&Delete Category"
         End
         Begin VB.Menu mnurename 
            Caption         =   "&Rename Category"
         End
         Begin VB.Menu mnufindcat 
            Caption         =   "&Find Category"
         End
         Begin VB.Menu mnugenhem 
            Caption         =   "&Generate Web Page"
         End
      End
      Begin VB.Menu mnubookMark 
         Caption         =   "&Bookmarks"
         Begin VB.Menu mnunewbook 
            Caption         =   "&New Bookmark"
         End
         Begin VB.Menu mnumodbook 
            Caption         =   "&Modify Bookmark"
         End
         Begin VB.Menu mnudelbook 
            Caption         =   "&Delete Bookmark"
         End
         Begin VB.Menu mnublank1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuexp 
            Caption         =   "&Export to IE favorites"
         End
         Begin VB.Menu mnumove 
            Caption         =   "&Move Bookmark to"
         End
         Begin VB.Menu mnufindbook 
            Caption         =   "&Find Bookmark"
         End
         Begin VB.Menu mnucbk 
            Caption         =   "&Create Shortcut"
         End
      End
      Begin VB.Menu mnupasteIE 
         Caption         =   "&Paste URL From IE"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Begin VB.Menu mnucpybook 
            Caption         =   "&Bookmark Name"
         End
         Begin VB.Menu mnubkloc 
            Caption         =   "Bookmark &Location"
         End
         Begin VB.Menu mnudate 
            Caption         =   "&Added Date"
         End
         Begin VB.Menu mnublank3 
            Caption         =   "-"
         End
         Begin VB.Menu mnulastv 
            Caption         =   "&Last Viewed"
         End
         Begin VB.Menu mnubkhits 
            Caption         =   "&Bookmark Hits"
         End
         Begin VB.Menu mnubkhtm 
            Caption         =   "Bookmark &HTML Code"
         End
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuvCats 
         Caption         =   "&Categorys"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuvdescription 
         Caption         =   "&Description Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuvtime 
         Caption         =   "&Current Time"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuwensnap 
         Caption         =   "&Web Snapshot"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnulayout 
         Caption         =   "&Layout"
         Begin VB.Menu mnureport 
            Caption         =   "&Report"
         End
         Begin VB.Menu mnuicon 
            Caption         =   "&Icon"
         End
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnubkstats 
         Caption         =   "Bookmark &Status"
      End
      Begin VB.Menu mnusavepage 
         Caption         =   "&Save Web Page"
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnushLnk 
         Caption         =   "&Share Bookmark"
      End
      Begin VB.Menu mnuping 
         Caption         =   "&Ping"
      End
      Begin VB.Menu mnuconf 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuaddin 
      Caption         =   "Add-ins"
      Begin VB.Menu mnuplg 
         Caption         =   "&Install Add-ins"
         Index           =   0
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucont 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuTip 
         Caption         =   "&Tip of the Day"
      End
      Begin VB.Menu mnureadme 
         Caption         =   "&Read Me"
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnublank7 
         Caption         =   "-"
      End
      Begin VB.Menu manuabout 
         Caption         =   "&About URL Manager - Home Edition"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BETA 1 VER 1.1.1

Dim HideCatlist As Boolean ' Stats of the cat list
Dim ShowDescription As Boolean
Dim WebView As Boolean

Private Const WebPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"

Private Sub DisableItems(mEnable As Boolean)
On Error Resume Next

    If mEnable Then
        tBar1.Buttons(4).Enabled = True
        tBar1.Buttons(6).Enabled = True
        tBar1.Buttons(7).Enabled = True
        tBar1.Buttons(8).Enabled = True
        tBar1.Buttons(9).Enabled = True
        tBar1.Buttons(10).Enabled = True
        tBar1.Buttons(11).Enabled = True
        tBar1.Buttons(12).Enabled = True
        tBar1.Buttons(14).Enabled = True
        tBar1.Buttons(15).Enabled = True
        tBar1.Buttons(16).Enabled = True
        tBar1.Buttons(18).Enabled = True
        tBar1.Buttons(20).Enabled = True
        tBar1.Buttons(21).Enabled = False
        
        CatBar.Buttons(5).Enabled = True
        tBarWebBro.Buttons(1).Enabled = True
        tBarWebBro.Buttons(2).Enabled = True
        tBarWebBro.Buttons(3).Enabled = True
        tBarWebBro.Buttons(4).Enabled = True
        tBarWebBro.Buttons(5).Enabled = True
        
        mnuexp.Enabled = True
        mnuedit.Enabled = True
        mnubkstats.Enabled = True
        mnusavepage.Enabled = True
        mnushLnk.Enabled = True
        mnuview.Enabled = True
        mnubak.Enabled = True
        txtserach.Enabled = True
    Else
        tBar1.Buttons(1).ButtonMenus(1).Enabled = False
        tBar1.Buttons(4).Enabled = False
        tBar1.Buttons(7).Enabled = False
        tBar1.Buttons(8).Enabled = False
        tBar1.Buttons(9).Enabled = False
        tBar1.Buttons(10).Enabled = False
        tBar1.Buttons(11).Enabled = False
        tBar1.Buttons(12).Enabled = False
        tBar1.Buttons(14).Enabled = False
        tBar1.Buttons(15).Enabled = False
        tBar1.Buttons(16).Enabled = False
        tBar1.Buttons(18).Enabled = False
        tBar1.Buttons(20).Enabled = False
        tBar1.Buttons(21).Enabled = False

        tBarWebBro.Buttons(1).Enabled = False
        tBarWebBro.Buttons(2).Enabled = False
        tBarWebBro.Buttons(3).Enabled = False
        tBarWebBro.Buttons(4).Enabled = False
        tBarWebBro.Buttons(5).Enabled = False
        
        CatBar.Buttons(1).Enabled = False
        CatBar.Buttons(2).Enabled = False
        CatBar.Buttons(3).Enabled = False
        CatBar.Buttons(4).Enabled = False
        CatBar.Buttons(5).Enabled = False
        mnuexp.Enabled = False
        mnuedit.Enabled = False
        mnubkstats.Enabled = False
        mnusavepage.Enabled = False
        mnushLnk.Enabled = False
        mnuview.Enabled = False
        mnubak.Enabled = False
        txtserach.Enabled = False
    End If
    
End Sub

Public Function RunAddins(AddinIdx As Integer)
Dim myObj As Object
Dim plgFile As String
Dim ans As Integer
Dim lsResult As Boolean

On Error GoTo PlugErr
    plgFile = FixPath(App.Path) & "add-ins\" & mAddins(AddinIdx).mPlugDLL
    
    If Not FindFile(plgFile) Then
        MsgBox mAddins(AddinIdx).mPlugDLL & " Could not be found.", vbCritical, "File not found"
        Exit Function
    End If
    
    Set myObj = CreateObject(mAddins(AddinIdx).mPlugClass) ' Create the add-in object
    myObj.lzBookMark = TBookMark                     ' Sends the whois server to the plugin
    myObj.lzBookmarkURL = TBookURL
    myObj.WhoisServ = Config.mWHOIS_serv           ' Send the bookmark dommain name
    myObj.lzHwnd = frmmain.hwnd
    myObj.RunPlug                               ' Run the plug-in
PlugErr:

    If Err Then
        ans = MsgBox("The addin you specified is not currently registered." & DoubleCRLF _
        & "would you like to register this add-in now?", vbYesNo Or vbQuestion, "Error loading add-in")
        If ans = vbNo Then
            Exit Function
        Else
           lsResult = RegisterActiveX(plgFile, Register)
           If Not lsResult Then
                MsgBox "There was an unexpected error while registering the add-in.", vbExclamation, "Unexpected error"
                Exit Function
            Else
                RunAddins AddinIdx
            End If
        End If
    End If
    
    ans = 0
    plgFile = ""
    Set myObj = Nothing
    
    
End Function

Public Function InitDB(dbFilename As String)
Dim OpenAsReO As Boolean
On Error Resume Next

    d_base.Close
    Set RecoredSet = Nothing
    
    DbPath = dbFilename ' Path and file name of the database
    
    If FindFile(DbPath) = False Then
        MsgBox "There was an error while loading your Bookmarks Database." & _
        DoubleCRLF & "Please check that the programs settings are correct." & vbCrLf & "Or check that the Database has not been deleted by mistake." _
        & vbCrLf & "You can set the database by going to tool->options", vbCritical, "Bookmarks Database not found"
        DbError = True
        DisableItems False
        Exit Function
    Else
        DbError = False
        DisableItems True
    End If
    
    ' The code below check if the database is readonly.
    If GetAttr(DbPath) = vbArchive + vbReadOnly Then
        OpenAsReO = True ' the database was found to be readonly
    Else
        OpenAsReO = False ' Database was not readonly
    End If
    
    ' the code below will open the database.
    Set d_base = OpenDatabase(DbPath, False, OpenAsReO, ";pwd=idkfa")
    StatusBar1.Panels(3).Text = "DB Version : " & d_base.Version ' show the db version

    InitTv
    
End Function

Public Sub UpdateForm()
    frmmain.WindowState = 0  ' Update window state
    frmmain.Show ' Show the form
    Tray1.ToolTip = "" ' remove the tooltip we finished with it
    Tray1.Visible = False ' Hide tray
End Sub

Sub ShowSnapShot(nSnapFilename As String)
' This sub is used to show the snap shot of the website

    If UCase(Trim$(nSnapFilename)) = "NONE" Then
        picsh.Picture = LoadResPicture(101, vbResBitmap) ' Use default image
        Exit Sub
    ElseIf FindFile(nSnapFilename) = False Then
        ' the snap shot was not found s we must use default one
        picsh.Picture = LoadResPicture(101, vbResBitmap)
        Exit Sub
    Else
        ' The snap shot file was found so we show it
        picsh.Picture = LoadPicture(nSnapFilename)
    End If
    
End Sub

Private Sub OpenSite()
Dim tRet As Long
On Error Resume Next

    If Not Len(TBookURL) > 0 Then Exit Sub ' No URL was found zso exit here
    
    UpdateDB SiteID, RecoredName
    ' New code Starts Here
    lstsites.ListItems.Clear
    LoadSites RecoredName
    lstsites.ListItems(LstIndex).Selected = True
    ' End of New code
    
    Select Case Val(Config.defBrowser) ' Find what's the default web broswer
        Case IE
            'ChDir WebBor.IE   ' Move to the programs folder
            'tRet = WinExec("IEXPLORE.EXE " & TBookURL, 3)
            'WinExec "IEXPLORE.EXE " & TBookURL, 3
            tRet = WinExec(ShortPath(WebBor.IE) & "\IEXPLORE.EXE " & TBookURL, 1)
        Case Netscape
            'ChDir WebBor.Netscape ' Move to the programs folder
            tRet = WinExec(ShortPath(WebBor.Netscape) & "\Netscape.exe " & TBookURL, 1)
        Case Opera
            'ChDir WebBor.Opera ' Move to the programs folder
             tRet = WinExec(ShortPath(WebBor.Opera) & "\Opera.exe " & TBookURL, 1)
        Case Mozilla
           'ChDir WebBor.Mozilla ' Move to the programs folder
            tRet = WinExec(ShortPath(WebBor.Mozilla) & "\Mozilla.exe " & TBookURL, 1)
        Case FireFox
            'ChDir WebBor.FireFox
            tRet = WinExec(ShortPath(WebBor.FireFox) & "\firefox.exe " & TBookURL, 1)
    End Select
   
    ChDir Abspath ' Move back to the programs folder
    
    
    Exit Sub
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
    
End Sub

Sub GetBroswer()
On Error Resume Next
    WebBor.IE = GetString(HKEY_LOCAL_MACHINE, WebPath & "IEXPLORE.EXE", "Path")
    If Right(WebBor.IE, 1) = ";" Then WebBor.IE = Left(WebBor.IE, Len(WebBor.IE) - 1)
    WebBor.Netscape = GetString(HKEY_LOCAL_MACHINE, WebPath & "Netscape.exe", "Path")
    WebBor.Opera = GetString(HKEY_LOCAL_MACHINE, WebPath & "Opera.exe", "Path")
    WebBor.Mozilla = GetString(HKEY_LOCAL_MACHINE, WebPath & "Mozilla.exe", "Path")
    WebBor.FireFox = GetString(HKEY_LOCAL_MACHINE, WebPath & "firefox.exe", "Path")
    If Err Then MsgBox Err.Description, vbInformation, "ERROR_" & Err.Number
End Sub

Public Function FindCat(tvC As TreeView, CatName As String) As Boolean
Dim I As Long
Dim ToFind As String

    FindCat = False
    ToFind = UCase(CatName)
    
    For I = 2 To tv1.Nodes.Count
        If UCase(tvC.Nodes(I).Text) = ToFind Then
            FindCat = True
            tvC.Nodes(I).Selected = True
            tv1_Click
            I = 0
            ToFind = ""
            Exit Function
        End If
    Next
    
    I = 0
    ToFind = ""
    
End Function
Sub KillTmpFile()
On Error Resume Next

    If Not FindFile(GetTempFolder & "tmpxtygtp4.html") Then
        Exit Sub
    Else
        Kill GetTempFolder & "tmpxtygtp4.html"
        Kill GetTempFolder & "tmpImg.gif"
    End If
    
End Sub

Sub hideBar(mHide As Boolean)
    If mHide Then
        mnuvCats.Checked = True
        CatBar.Visible = True
        HideCatlist = mHide
        picshbut.Visible = False
        sidebar.Visible = False
        Frame1.Visible = True
        Frame1.Left = 15
        Frame1.Top = 780
        tv1.Visible = True
        tv1.Left = 15
        tv1.Top = 1335
        lstsites.Top = 885
        lstsites.Left = (tv1.Width + 50)
        WebB.Left = lstsites.Left
        picshot.Visible = WebView
        
        picshot.Left = tv1.Left
    Else
        mnuvCats.Checked = False
        HideCatlist = mHide
        Frame1.Visible = False
        tv1.Visible = False
        CatBar.Visible = False
        lstsites.Left = 480
        WebB.Left = lstsites.Left
        lstsites.Top = 870
        picshbut.Top = 945
        picshbut.Left = 75
        sidebar.Top = 870
        sidebar.Left = 15
        picshbut.Visible = True
        sidebar.Visible = True
        picshot.Visible = False

        
    End If
    
End Sub

Sub ClearAllVars()
On Error Resume Next
    d_base.Close
    Set d_base = Nothing
    Set Recored_Set = Nothing
    
    plgIni = ""
    RecoredName = ""
    TBookMark = ""
    TBookURL = ""
    TBookHit = 0
    BookCount = 0
    EdURL.TAddLastVis = 0
    EdURL.TDateAdded = ""
    EdURL.TSiteDescription = ""
    EdURL.TSiteName = ""
    EdURL.TSiteURL = ""
    EdURL.THitCnt = 0
    EdURL.TIcon = 0
    EdURL.TRated = 0
    EdURL.TVieded = 0
    EdURL.TWebCap = ""
    
    DbPath = ""
    Config.mSMTP_serv = ""
    Config.mWHOIS_serv = ""
    ' unload all the forms
    Set frmconfig = Nothing
    Set frmabout = Nothing
    Set frmAddUrl = Nothing
    Set frmEdit = Nothing
    Set frmnews = Nothing
    Set frmping = Nothing
    Set frmPopUpmenu = Nothing
    Set frmref = Nothing
    Set frmshurl = Nothing
    Set frmsplash = Nothing
    Set frmmain = Nothing
    Tray1.Visible = False
    End
    
End Sub
Sub InitTv()
On Error Resume Next

    If Val(Config.FavIcon) = 0 Then Config.FavIcon = 1
    tv1.Nodes.Clear
    tv1.Indentation = 49
    tv1.Nodes.Add , tvwFirst, "bookmark", "Bookmarks", 11, 11
    For Each T_def In d_base.TableDefs
        If T_def.Attributes = 0 Then
            tv1.Nodes.Add 1, tvwChild, T_def.Name, T_def.Name, Val(Config.FavIcon), Val(Config.FavIcon) + 1
        End If
    Next

    tv1.Nodes(2).Selected = True
    tv1.Nodes(1).Selected = True
    tv1_Click
    
End Sub

Private Sub CatBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "NEW_CAT"
            mnunewcat_Click
        Case "DEL_CAT"
            mnudelcat_Click
        Case "RENAME_CAT"
            mnurename_Click
        Case "GEN_CAT"
            mnugenhem_Click
        Case "FIND_CAT"
            mnufindcat_Click
        Case "HIDE_CAT"
            mnuvCats_Click
            hideBar False ' Hide the cat bar
            Form_Resize ' Resize the form
    End Select

End Sub

Sub LoadAddinIni(plgIni As String)
Dim sFile As Long, I As Long, aStr As String, vStr As Variant

    On Error Resume Next
    Erase mAddins ' Erase the array
    sFile = FreeFile ' Pointer to a free file
    
    Open plgIni For Input As #sFile ' Open the file for input
        Do While Not EOF(sFile) ' Loop while we'r not at the end of the file
            Input #sFile, aStr ' Input each line
            If Left(aStr, 7) = "plgname" Then ' check for vaild plugin line
                I = I + 1 ' up date our counter
                ReDim Preserve mAddins(I) ' Resize array
                aStr = Right(aStr, Len(aStr) - 8) ' Extract the plugin info
                vStr = Split(aStr, ";")
                
                If UBound(vStr) > 0 Then
                    mAddins(I).mPlugName = vStr(0)
                    mAddins(I).mPlugClass = vStr(1)
                    mAddins(I).mPlugDLL = vStr(2)
                    Load mnuplg(I)
                    mnuplg(I).Caption = mAddins(I).mPlugName
                End If
                
            End If
            DoEvents
        Loop
        Close #sFile
        ' Clear up
        Erase vStr
        aStr = ""
        I = 0
    Close #FILE
    
End Sub





Private Sub Form_Initialize()
    Dim x As Long
    x = InitCommonControls
End Sub

Private Sub Form_Load()
Dim Col_Head As ColumnHeader, lstItem As ListItem
Dim ans As Integer
On Error Resume Next

    Me.MousePointer = vbHourglass
    
    If IsAppOpen Then
        MsgBox "An Instance of DM Bookmarks is already open.", vbInformation, App.ProductName
        End
    End If
        
    plgIni = FixPath(App.Path) & "add-ins\addins.ini" ' Path to the plugins list
    SerachLst = FixPath(App.Path) & "Serach.ini"
    LoadTvImg picsrc, picdst, ImageList1
    
    ReadServConfig
    Text1.Text = Config.Hightlight
    If Not FindFile(plgIni) Then
        ans = MsgBox("Unable to locate the programs add-ins list" _
        & DoubleCRLF & "Do you still whish to continue loading the program?", vbYesNo Or vbQuestion)
        If ans = vbNo Then
            End
        End If
    Else
        LoadAddinIni plgIni
    End If

    InitDB Config.mbkDatabase ' load in the database
    
  ' mWnd = frmmain.hwnd ' Hangle of the form
   'Hook ' Hook the form
    
    KillTmpFile ' Kill tmp file
    InitTv ' Load in the table in the treeview control
    ' Below code ajusts and adds colums to the listview control
    Abspath = FixPath(App.Path)
    GetBroswer ' Get the paths of the web browsers
    LoadImgLst picsrc, picdst, ImageList4 ' Load in the image list
    LoadSerchList ' Load in the serach engines list
    
    lstsites.ColumnHeaders.Add , , "Site Name", 2510, , 0
    Set Col_Head = lstsites.ColumnHeaders.Add(, , "Address", 2510)
    Set Col_Head = lstsites.ColumnHeaders.Add(, , "Date Added", 1318)
    Set Col_Head = lstsites.ColumnHeaders.Add(, , "Last Viewed", 1318)
    Set Col_Head = lstsites.ColumnHeaders.Add(, , "Hits", 600)
    
    ShowHeaderIcon lstsites, 0, 0, True
    picsh.MouseIcon = LoadResPicture(101, vbResCursor) ' Load cursor for websnap
    
    SerachOption = 1 ' Default serach engin to use
    WebB.Navigate "about:blank"
    StrSql = ""

    MakeFlatControls frmmain
    SetLstVFlScroolBar lstsites
    WebView = True
    hideBar True
    
    mnuwensnap.Checked = True
    mnuvtime.Checked = True
    mnuvdescription.Checked = True
    ShowDescription = True
   
    If CBool(Config.dmViewCat) = False Then
        mnuvCats_Click
    End If
    
    If CBool(Config.dmViewDesWnd) = False Then
        mnuvdescription_Click
    End If

    If CBool(Config.dmViewTime) = False Then
        mnuvtime_Click
    End If
    
    If CBool(Config.dmViewWebView) = False Then
        mnuwensnap_Click
    End If
    
    If CInt(Config.ShowTips) = 1 Then
        frmtip.Show modal, frmmain
    End If
    
    tv1.Nodes(Val(Config.dmLastView)).Selected = True
    tv1_Click

    Me.MousePointer = vbNormal
    lstsites.View = Val(Config.dmView)
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    DoEvents
  
    tbarserach.Left = (frmmain.ScaleWidth - tbarserach.Width) - 99
    tbarfind.Left = (tbarserach.Left) - 810
    txtserach.Width = (tbarfind.Left) - 1800
    
    sidebar.Height = (frmmain.ScaleHeight - StatusBar1.Height - sidebar.Top - 32)
    WebB.Top = (frmmain.ScaleHeight - StatusBar1.Height - WebB.Height - 40)
    picshot.Top = (frmmain.ScaleHeight - StatusBar1.Height - picshot.Height - 40)
    
    Line3D1.Width = frmmain.ScaleWidth - Line3D1.Left
    Line3D2.Width = frmmain.ScaleWidth - Line3D2.Left
    Line3D3.Width = frmmain.ScaleWidth - Line3D3.Left
    
    ' check if the Cat list is shown
    If HideCatlist Then
        lstsites.Width = (frmmain.ScaleWidth - tv1.Width - 60)
    Else
        lstsites.Width = (frmmain.ScaleWidth - sidebar.Width - 60)
    End If
    
    WebB.Width = (lstsites.Width)
    
    If ShowDescription Then
        lstsites.Height = (WebB.Top - 950)
    Else
        lstsites.Height = (frmmain.ScaleHeight - StatusBar1.Height - lstsites.Top - 48)
    End If
    
    
    If WebView Then
        tv1.Height = (frmmain.ScaleHeight - StatusBar1.Height - Frame1.Top - picshot.Height - 600)
    Else
        tv1.Height = (frmmain.ScaleHeight - StatusBar1.Height - Frame1.Top - 600)
    End If
    
    If WindowState = 1 Then
        ' Store all the old window settings
        frmmain.Hide ' Hide the window
        Tray1.ToolTip = frmmain.Caption ' Update tooltip text
        Tray1.Visible = True ' Make the tar icon visable
    End If
    
    If Err Then
        frmmain.Height = 5370
        frmmain.Width = 8775
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "DMUrlMan", "Config", "dmLastView", CStr(tv1.SelectedItem.Index)
    ClearAllVars
End Sub

Private Sub lstsites_Click()
    If Len(TBookMark) = 0 Then Exit Sub
    If Val(Config.mOpenItems) = 2 Then
        Exit Sub
    Else
        OpenSite
    End If
    
End Sub

Private Sub lstsites_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static sSort As Integer
Dim I As Long

    sSort = Not sSort
    lstsites.SortKey = ColumnHeader.Index - 1
    lstsites.SortOrder = Abs(sSort)
    lstsites.Sorted = True
    
    For I = 0 To lstsites.ColumnHeaders.Count - 1
      If I = lstsites.SortKey Then
            ShowHeaderIcon lstsites, lstsites.SortKey, lstsites.SortOrder, True
        Else
            ShowHeaderIcon lstsites, I, 0, False
      End If
   Next
   
   I = 0
   
End Sub

Private Sub lstsites_DblClick()
    If Not Val(Config.mOpenItems) = 2 Then
        Exit Sub
    Else
        OpenSite
    End If
    
End Sub

Private Sub lstsites_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'On Error Resume Next
    KillTmpFile
    LstIndex = lstsites.SelectedItem.Index
    If LstIndex <= 0 Then Exit Sub
    
    TBookMark = lstsites.SelectedItem.Text          ' Get the bookmark name
    TBookURL = lstsites.SelectedItem.SubItems(1)    ' Get the bookmarks URL
    TBookHit = Val(lstsites.SelectedItem.SubItems(4))
    TBookAddDate = lstsites.SelectedItem.SubItems(2)
    TBookLastVist = lstsites.SelectedItem.SubItems(3)
    
    SiteID = Val(Right(lstsites.SelectedItem.Key, Len(lstsites.SelectedItem.Key) - 1)) ' get the website ID
    ShowInfo RecoredName, SiteID ' Show the bookmarks
    WebB.Navigate GetTempFolder & "tmpxtygtp4.html" ' Update the description webpage
    picsh.MousePointer = vbCustom
    picsh.Tag = TBookURL
    picsh.Enabled = True
    ShowSnapShot TBookSnapShot

    tBar1.Buttons(7).Enabled = True    ' Enable modify bookmark button
    tBar1.Buttons(8).Enabled = True    ' Enable bookmark button
    tBar1.Buttons(9).Enabled = True    ' Enable copy bookmark button
    tBar1.Buttons(11).Enabled = True   ' Enable move bookmark
    tBar1.Buttons(14).Enabled = True   ' Enable bookmark status button
    tBar1.Buttons(15).Enabled = True   ' Enable bookmark save web page button
    tBar1.Buttons(16).Enabled = True   ' Enable bookmark share link
    tBar1.Buttons(18).Enabled = True   ' Enable create short cut button
    tBar1.Buttons(21).Enabled = True
    
    mnucopy.Enabled = True ' Enable the copy button
    mnumodbook.Enabled = True  ' Enable Modify bookmark button
    mnudelbook.Enabled = True  ' Enable delete bookmark button
    mnumove.Enabled = True     ' Enable moveto bookmark button
    mnucbk.Enabled = True      ' Enable create short cut button
    mnubkstats.Enabled = True  ' Enable bookmark status menu
    mnusavepage.Enabled = True ' Enable save webpage button
    mnushLnk.Enabled = True    ' Enable share bookmark menu button
    mnuping.Caption = "&Ping [" & TBookMark & "]"
    ' Enable the web broswer button
    tBarWebBro.Buttons(1).Enabled = True
    tBarWebBro.Buttons(2).Enabled = True
    tBarWebBro.Buttons(3).Enabled = True
    tBarWebBro.Buttons(4).Enabled = True
    tBarWebBro.Buttons(5).Enabled = True
    
    DoEvents
    If TViewed = 1 Then
        tBar1.Buttons(21).Image = ImageList2.ListImages(29).Index
        tBar1.Buttons(21).Value = tbrPressed
    Else
        tBar1.Buttons(21).Image = ImageList2.ListImages(29).Index
        tBar1.Buttons(21).Value = tbrUnpressed
    End If
    
    
End Sub

Private Sub lstsites_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(TBookMark) <= 0 Then
        Exit Sub
    ElseIf KeyCode = 46 Then
        mnudelbook_Click
    End If

End Sub

Private Sub lstsites_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstsites.ListItems.Count <= 0 Then Exit Sub
    If Button = vbLeftButton Then
        
    End If
    
End Sub

Private Sub lstsites_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstsites.ListItems.Count <= 0 Then Exit Sub
    If Not Button = vbRightButton Then Exit Sub
    frmPopUpmenu.mnuping.Caption = "&Ping [" & PhaseDomain(TBookMark) & "]"
    PopupMenu frmPopUpmenu.mnuedit
    
End Sub

Private Sub manuabout_Click()
    frmabout.Show vbModal, frmmain
End Sub


Private Sub mnubak_Click()
Dim ans As Integer
On Error GoTo CanErr
' Backup database sub
    With CDialog1
        .CancelError = True
        .DialogTitle = "Backup Bookmarks"
        .Filter = "DM Bookmarks Backup Set(*.dbk)|*.dbk|"
        .InitDir = Abspath
        .FileName = "Bookmarks"
        .ShowSave
        If Not GetFileExt(.FileName) = "DBK" Then .FileName = Left(.FileName, Len(.FileName) - 3) & "dbk"
        
        If FindFile(.FileName) Then
            ans = MsgBox(.FileName & vbCrLf & "Already exists do you want to replace this file with the new one?", vbYesNo Or vbQuestion, .DialogTitle)
            If ans = vbNo Then
                Exit Sub
            Else
                On Error Resume Next
                SetAttr .FileName, vbNormal
                Kill .FileName
                If Err Then
                    MsgBox Err.Description, vbCritical, "Error_" & Err.Number
                    Exit Sub
                Else
                    BackupDB .FileName
                    MsgBox "Your bookmarks have now been successfully backed up.", vbInformation, .DialogTitle
                    Exit Sub
                End If
            End If
        Else
            BackupDB .FileName
            MsgBox "Your bookmarks have now been successfully backed up.", vbInformation, .DialogTitle
        End If
CanErr:
        If Err = cdlCancel Then Err.Clear
        
    End With

End Sub

Private Sub mnubkhits_Click()
    CopyCommand m_SiteHits, lstsites.SelectedItem.SubItems(4)
End Sub

Private Sub mnubkhtm_Click()
Dim aStr As String

    aStr = "<a href=" & Chr(34) & lstsites.SelectedItem.SubItems(1) & Chr(34) & " target=" & Chr(34) & "_blank" _
    & Chr(34) & ">" & lstsites.SelectedItem.Text & "</a>"
    CopyCommand m_HtmlCode, aStr
    aStr = ""
    
End Sub

Private Sub mnubkloc_Click()
    CopyCommand m_SiteURL, lstsites.SelectedItem.SubItems(1) ' Copy bookmark URL
End Sub

Private Sub mnubkstats_Click()
    frmstatus.Show vbModal, frmmain
End Sub



Private Sub mnucbk_Click()
    frmshurl.Show vbModal, frmmain ' Show create shortcut
End Sub

Private Sub mnuconf_Click()
    frmconfig.Show vbModal, frmmain
End Sub

Private Sub mnucont_Click()
    On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, FixPath(App.Path) & "Help\help.chm", vbNullString, vbNullString, 1
    If Err Then MsgBox Err.Description, vbOKOnly, Err.Number
    
End Sub

Private Sub mnucpybook_Click()
    CopyCommand m_SiteName, lstsites.SelectedItem.Text  ' Copy bookmark name
End Sub

Private Sub mnudate_Click()
    CopyCommand m_SiteAddedDate, lstsites.SelectedItem.SubItems(2) ' Copy bookmark added date
End Sub

Private Sub mnudelbook_Click()
    frmPopUpmenu.DelBookmark
End Sub

Private Sub mnudelcat_Click()
    frmPopUpmenu.DelCat
    
End Sub

Private Sub mnuexit_Click()
    Unload frmmain
End Sub

Private Sub mnuexp_Click()
On Error Resume Next
Dim FolName As String
' This function is used to output all the book marks to a Favorites,
' That may be used for internet explorer
Dim I As Long, J As Long, tFile As Long
Dim FavFol As String, sFol As String, rName As String, StrB As String
Dim vData As Variant, sLn As Variant
    
    ButtonIndex = 8 ' Button index to chnage caption on
    ButtonCaption(ButtonIndex) = "Export" ' Buttons new caption
    SubClass frmmain ' Sub Class the form

    FolName = GetFolder(frmmain.hwnd, "Choose Folder:")
    If Len(FolName) <= 0 Then Exit Sub
    FavFol = FixPath(FolName) & "Bookmarks\"
    MkDir FavFol ' Create the Favorites Folder
    
    tFile = FreeFile
    For I = 2 To tv1.Nodes.Count    ' Get all the node names from the treeview control
        rName = tv1.Nodes(I).Text   ' Get each node name
        StrB = ExportToIE(rName)    ' Build links for each recored set
        sLn = Split(StrB, vbCrLf)
        
        For J = LBound(sLn) To UBound(sLn)
            If Len(sLn(J)) > 0 Then     ' Check the length of the URL
                sFol = FavFol & rName   ' Sub folder name
                vData = Split(sLn(J), Chr(128)) ' Split out the URL name
                MkDir sFol ' create the sub folder
                Open sFol & "\" & vData(0) & ".url" For Output As #tFile
                    Print #tFile, "[InternetShortcut]"
                    Print #tFile, "URL=" & vData(1)
                    Print #tFile, "IconIndex=3"
                    Print #tFile, "IconFile=" & FixPath(App.Path) & App.EXEName & ".exe"
                Close #tFile
            End If
        Next
        DoEvents
    Next
    ' Clear up vars
    Erase vData
    Erase sLn
    sFol = ""
    FavFol = ""
    StrB = ""
    I = 0
    J = 0

End Sub

Private Sub mnufind_Click()
    frmserach.Show vbModal, frmmain
End Sub

Private Sub mnufindbook_Click()
    frmserach.Show vbModal, frmmain
End Sub

Private Sub mnufindcat_Click()
    frmfindcat.Show vbModal, frmmain
End Sub

Private Sub mnugenhem_Click()
    frmgenhtm.Show vbModal, frmmain
End Sub

Private Sub mnuicon_Click()
    SaveSetting "DMUrlMan", "Config", "View", "0"
    lstsites.View = lvwIcon
End Sub

Private Sub mnulastv_Click()
    CopyCommand m_SiteLastVistDate, lstsites.SelectedItem.SubItems(3)
End Sub

Private Sub mnumodbook_Click()
    If Len(TBookMark) = 0 Then MsgBox "You must first click on the bookmark you want like to modify.", vbInformation, frmmain.Caption: Exit Sub
    frmEdit.Show vbModal, frmmain
End Sub

Private Sub mnumove_Click()
    frmmoveto.Show vbModal, frmmain ' show moveto bookmark dialog
End Sub

Private Sub mnunew_Click()
Dim ans As Integer
On Error GoTo CanErr
    With CDialog1
    
        .CancelError = True
        .DialogTitle = "Create new Database"
        .Filter = "Microsoft Access Databases(*.mdb)|*.mdb|"
        .InitDir = Abspath
        .FileName = "Bookmarks.mdb"
        .ShowSave
        
        If Not GetFileExt(.FileName) = "MDB" Then .FileName = Left(.FileName, Len(.FileName) - 3) & "mdb"
        
        If FindFile(.FileName) Then
            ans = MsgBox(.FileName & vbCrLf & "Already exists do you want to replace this file with the new one?", vbYesNo Or vbQuestion, .DialogTitle)
            If ans = vbNo Then
                Exit Sub
            Else
                On Error Resume Next
                SetAttr .FileName, vbNormal
                Kill .FileName
                If Err Then
                    MsgBox Err.Description, vbCritical, "Error_" & Err.Number
                    Exit Sub
                Else
                    DBEngine.Workspaces(0).CreateDatabase .FileName, dbLangGeneral
                    MsgBox .FileName & vbCrLf & "Has been successfully created.", vbInformation, .DialogTitle
                    Exit Sub
                End If
            End If
        Else
            DBEngine.Workspaces(0).CreateDatabase .FileName, dbLangGeneral
            MsgBox .FileName & vbCrLf & "Has been successfully created.", vbInformation, .DialogTitle
        End If
CanErr:
        If Err = cdlCancel Then Err.Clear
        
    End With
    
End Sub

Private Sub mnunewbook_Click()
    If Len(Trim(RecoredName)) <= 0 Then
        MsgBox "You must first select a category.", vbInformation, frmmain.Caption
        Exit Sub
    Else
        FromWeb = False ' Book mark was not added form the web broswer
        frmAddUrl.Show vbModal, frmmain ' show add new bookmark dialog
    End If
End Sub

Private Sub mnunewcat_Click()
    frmaddcat.Show vbModal, frmmain
End Sub

Private Sub mnuopen_Click()
On Error GoTo CanErr
Dim ans As Integer

    With CDialog1
        .CancelError = True
        .DialogTitle = "Open Bookmarks Database"
        .Filter = "Microsoft Access Databases(*.mdb)|*.mdb|"
        .InitDir = Abspath
        .FileName = ""
        .ShowOpen
        If Not GetFileExt(.FileName) = "MDB" Then
            MsgBox "This is not a valid Access Database filename, or its format is not supported.", vbExclamation, .DialogTitle
            Exit Sub
        Else
            ans = MsgBox("Would you like this to be your default Datbase the next time you start the program?", vbYesNo Or vbQuestion, .DialogTitle)
            If ans = vbNo Then
                InitDB .FileName
                Exit Sub
            Else
                Config.mbkDatabase = .FileName
                SaveSetting "DMUrlMan", "Config", "Db", Config.mbkDatabase
                InitDB .FileName
            End If
        End If
        Exit Sub
CanErr:
        If Err = cdlCancel Then
            Err.Clear
        End If
        
    End With
    
End Sub

Private Sub mnupasteIE_Click()
    MsgBox "Feature not found", vbInformation
End Sub

Private Sub mnuping_Click()
    frmping.Show vbModal, frmmain
End Sub

Private Sub mnuplg_Click(Index As Integer)
    Select Case Index
        Case 0 ' Install plugin
            frmregdll.Show vbModal, frmmain
        Case Else
            RunAddins Index
    End Select
    
End Sub

Private Sub mnureadme_Click()
 On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, FixPath(App.Path) & "Help\Readme\Readme.htm", vbNullString, vbNullString, 1
    If Err Then MsgBox Err.Description, vbOKOnly, Err.Number

End Sub

Private Sub mnurename_Click()
    frmrename.Show vbModal, frmmain
End Sub

Private Sub mnureport_Click()
    SaveSetting "DMUrlMan", "Config", "View", "3"
    lstsites.View = lvwReport
End Sub

Private Sub mnurestore_Click()
    frmrestore.Show vbModal, frmmain
End Sub

Private Sub mnusavepage_Click()
Dim lzFile As String, FileExt As String
On Error Resume Next
    CDialog1.DialogTitle = "Save Web Page"
    CDialog1.Filter = "Hyper Text File(*.html)|*.html|"
    CDialog1.ShowSave
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

Private Sub mnushLnk_Click()
    frmref.Show vbModal, frmmain
End Sub

Private Sub mnuTip_Click()
On Error Resume Next
    If FindFile(FixPath(App.Path) & "Tips.ini") = False Then
        MsgBox "Unable to load the tips file.", vbCritical, "File Not Found"
        Exit Sub
    Else
        If Err Then Err.Clear
        frmtip.Show vbModal, frmmain
    End If
    
End Sub

Private Sub mnuvCats_Click()

    If mnuvCats.Checked Then
        mnuvCats.Checked = False
    ElseIf mnuvCats.Checked = False Then
        mnuvCats.Checked = True
    End If
    
    SaveSetting "DMUrlMan", "Config", "dmViewCats", CStr(mnuvCats.Checked)
    hideBar mnuvCats.Checked
    Form_Resize
    
End Sub

Private Sub mnuvdescription_Click()
    If mnuvdescription.Checked = True Then
        mnuvdescription.Checked = False
    ElseIf mnuvdescription.Checked = False Then
        mnuvdescription.Checked = True
    End If
    
    SaveSetting "DMUrlMan", "Config", "dmViewDesWnd", CStr(mnuvdescription.Checked)
    ShowDescription = mnuvdescription.Checked
    WebB.Visible = ShowDescription
    Form_Resize
    
End Sub

Private Sub mnuvtime_Click()
    If mnuvtime.Checked = True Then
        mnuvtime.Checked = False
    ElseIf mnuvtime.Checked = False Then
        mnuvtime.Checked = True
    End If
    SaveSetting "DMUrlMan", "Config", "dmViewTime", CStr(mnuvtime.Checked)
    StatusBar1.Panels(2).Visible = mnuvtime.Checked
    
End Sub

Private Sub mnuwensnap_Click()
    If mnuwensnap.Checked Then
        mnuwensnap.Checked = False
    ElseIf mnuwensnap.Checked = False Then
        mnuwensnap.Checked = True
    End If
    
    SaveSetting "DMUrlMan", "Config", "dmViewWebView", CStr(mnuwensnap.Checked)
    WebView = mnuwensnap.Checked
    
    If Not HideCatlist Then
        picshot.Visible = False
    Else
        picshot.Visible = WebView
    End If
    
    Form_Resize

End Sub

Private Sub picsh_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        OpenSite
    End If
    
End Sub

Private Sub picshbut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then Exit Sub
    picshbut.BorderStyle = 1
End Sub

Private Sub picshbut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then Exit Sub
    picshbut.BorderStyle = 0
    mnuvCats_Click
    hideBar True
    Form_Resize
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim mCol As Long

    Select Case UCase(Button.Key)
        Case "OPEN_BOOK"
            mnuopen_Click
        Case "BK_DB"
            mnubak_Click
        Case "RES_BK"
            mnurestore_Click
        Case "EDIT_BOOK"
            mnumodbook_Click
        Case "DEL_BOOK"
            mnudelbook_Click
        Case "GET_BOOK"
            mnupasteIE_Click
        Case "MOVE_BOOK"
            mnumove_Click ' Show MoveTo bookmark dialog
        Case "FIND_BOOK"
            mnufindbook_Click
        Case "BOOK_STATE"
            mnubkstats_Click
        Case "SAV_WEB"
            mnusavepage_Click
        Case "SH_LINK"
            mnushLnk_Click
        Case "WHO_IS"
            RunAddins 1
        Case "CFG_NOW"
            mnuconf_Click
        Case "SH_CUT"
            mnucbk_Click
        Case "VIEW_STAT"
            UpdateViewStat SiteID, RecoredName, Button.Value
            
            If Button.Value = tbrPressed Then
                mCol = Val(Config.NewItems)
            Else
                mCol = vbBlack
            End If
            
            lstsites.ListItems(LstIndex).ForeColor = mCol
            lstsites.ListItems(LstIndex).ListSubItems(1).ForeColor = mCol
            lstsites.ListItems(LstIndex).ListSubItems(2).ForeColor = mCol
            lstsites.ListItems(LstIndex).ListSubItems(3).ForeColor = mCol
            lstsites.ListItems(LstIndex).ListSubItems(4).ForeColor = mCol
        Case "EXIT"
            mnuexit_Click
        Case "ABOUT"
            manuabout_Click
        Case "HELP"
            mnucont_Click
    End Select
    
End Sub

Private Sub tBar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
        Case "NEW_BOOK"
            mnunewbook_Click
        Case "NEW_DB"
            mnunew_Click
        Case "BOOK_NAME"
            mnucpybook_Click
        Case "BOOK_URL"
            mnubkloc_Click
        Case "BOOK_ADD_DATE"
            mnudate_Click
        Case "BOOK_LAST_VIEW"
            mnulastv_Click
        Case "BOOK_HIT"
            mnubkhits_Click
        Case "COPY_HTML"
            mnubkhtm_Click
        Case "MNU_REP"
            mnureport_Click
        Case "MNU_ICO"
            mnuicon_Click
    End Select
    
End Sub

Private Sub tbarfind_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrB As String, iRet As Long
    
    If Trim(txtserach.Tag) = "SR" Or Len(Trim(txtserach.Text)) <= 0 Then Exit Sub
    mSerachPattern = Trim(txtserach.Text)
    
    StrB = String(128, Chr$(0))
    iRet = GetPrivateProfileString("Serach" & SerachOption, "URL", "", StrB, 128, SerachLst)
    StrB = Left(StrB, iRet)
    StrB = Replace(StrB, "$SERACH$", mSerachPattern)
    
    Select Case Val(Config.defBrowser) ' Find what's the default web broswer
        Case IE
            'ChDir WebBor.IE ' Move to the programs folder
            'tRet = WinExec("IEXPLORE.EXE " & StrB, 3)
             tRet = WinExec(ShortPath(WebBor.IE) & "\IEXPLORE.EXE " & StrB, 1)
        Case Netscape
            'ChDir WebBor.Netscape ' Move to the programs folder
            'tRet = WinExec("Netscape.exe " & StrB, 3)
            tRet = WinExec(ShortPath(WebBor.Netscape) & "\Netscape.exe " & StrB, 1)
        Case Opera
            'ChDir WebBor.Opera ' Move to the programs folder
            'tRet = WinExec("Opera.exe " & StrB, 3)
            tRet = WinExec(ShortPath(WebBor.Opera) & "\Opera.exe " & StrB, 1)
        Case Mozilla
            tRet = WinExec(ShortPath(WebBor.Mozilla) & "\Mozilla.exe " & StrB, 1)
        Case FireFox
            tRet = WinExec(ShortPath(WebBor.FireFox) & "\firefox.exe " & StrB, 1)
    End Select
    
    ChDir Abspath ' Move back to the programs folder
    
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark.", vbInformation, frmmain.Caption
    End If
    
    mSerachPattern = ""
    StrB = ""
    iRet = 0
    iRet = 0
End Sub

Private Sub tbarfind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Trim(txtserach.Tag) = "SR" Or Len(Trim(txtserach.Text)) <= 0 Then Exit Sub
    mSerachPattern = Trim(txtserach.Text)
    
    tbarfind.ToolTipText = "Find " & mSerachPattern
    
End Sub

Private Sub tbarserach_ButtonClick(ByVal Button As MSComctlLib.Button)
    PopupMenu frmPopUpmenu.mnuser
End Sub

Private Sub tBarWebBro_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tRet As Long
On Error Resume Next

    Select Case UCase(Button.Key)
        Case "WEB_IE"
            'ChDir WebBor.IE
            'tRet = WinExec("IEXPLORE.EXE " & TBookURL, 3)
            tRet = WinExec(ShortPath(WebBor.IE) & "\IEXPLORE.EXE " & TBookURL, 1)
        Case "WEB_NS"
            'ChDir WebBor.Netscape
            tRet = WinExec(ShortPath(WebBor.Netscape) & "\Netscape.exe " & TBookURL, 1)
            'tRet = WinExec("Netscape.exe " & TBookURL, 3)
        Case "WEB_OP"
            'ChDir WebBor.Opera
            tRet = WinExec(ShortPath(WebBor.Opera) & "\Opera.exe " & TBookURL, 1)
            'tRet = WinExec("Opera.exe " & TBookURL, 3)
        Case "WEB_MZ"
            'ChDir WebBor.Mozilla
            'tRet = WinExec("Mozilla.exe " & TBookURL, 3)
            tRet = WinExec(ShortPath(WebBor.Mozilla) & "\Mozilla.exe " & TBookURL, 1)
        Case "WEB_FOX"
            'ChDir WebBor.FireFox
            tRet = WinExec(ShortPath(WebBor.FireFox) & "\firefox.exe " & TBookURL, 1)
           ' tRet = WinExec("firefox.exe " & TBookURL, 3)
    End Select
    
    ChDir Abspath ' Move back to the programs folder
    
    If tRet = 2 Then
        MsgBox "There was an error while opeing the bookmark." _
        & DoubleCRLF & "Please check to see that the web broswer your opening the bookmark is installed.", vbInformation, frmmain.Caption
    End If
    
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels(2).Text = Time
End Sub

Private Sub Tray1_MouseUp(Button As Integer)
On Error Resume Next
    If Button = vbRightButton Then
        PopupMenu frmPopUpmenu.mnutray
        Exit Sub
    Else
        UpdateForm
    End If
    
End Sub

Private Sub tv1_Click()
On Error Resume Next

    If DbError Then Exit Sub
    
    If tv1.Nodes.Count <= 0 Then Exit Sub
    TvSelIdx = tv1.SelectedItem.Index
    
    tBarWebBro.Buttons(1).Enabled = False
    tBarWebBro.Buttons(2).Enabled = False
    tBarWebBro.Buttons(3).Enabled = False
    tBarWebBro.Buttons(4).Enabled = False
    tBarWebBro.Buttons(5).Enabled = False
    
    RecoredName = ""    ' Clear recoredname buffer
    BookCount = 0
    lstsites.ListItems.Clear
    If tv1.SelectedItem.Key = "bookmark" Then
        WebB.Navigate "about:blank"
        picsh.Picture = Nothing
        picsh.MousePointer = Normal
        picsh.Enabled = False
         'Category bar
        CatBar.Buttons(1).Enabled = True    ' Enable new cat button
        CatBar.Buttons(2).Enabled = False   ' Disable delete button
        CatBar.Buttons(3).Enabled = False   ' Disable rename table button
        CatBar.Buttons(4).Enabled = True
        ' Top toolbar
        tBar1.Buttons(1).ButtonMenus(1).Enabled = False ' Disbale new bookmark button
        
        tBar1.Buttons(7).Enabled = False    ' modify bookmark button
        tBar1.Buttons(8).Enabled = False    ' delete bookmark button
        tBar1.Buttons(9).Enabled = False    ' copy bookmark button
        tBar1.Buttons(11).Enabled = False   ' Disbale move bookmark
        tBar1.Buttons(14).Enabled = False   ' Disbale bookmark status button
        tBar1.Buttons(15).Enabled = False   ' Disbale bookmark save web page button
        tBar1.Buttons(16).Enabled = False   ' Disable bookmark share link
        tBar1.Buttons(18).Enabled = False   ' Disable create short cut button
        tBar1.Buttons(21).Enabled = False
        

        ' Menu items
        mnunewcat.Enabled = True    ' Enable new cat button
        mnudelcat.Enabled = False   ' Disable delete cat button
        mnurename.Enabled = False   ' Disable rename cat button
        mnunewbook.Enabled = False  ' Disbale new bookmark button
        mnucopy.Enabled = False     ' Disable the copy button
        mnumodbook.Enabled = False  ' Disbale Modify bookmark
        mnudelbook.Enabled = False  ' Disable Delete bookmark
        mnumove.Enabled = False     ' Disable moveto bookmark button
        mnucbk.Enabled = False      ' Disable create short cut button
        mnubkstats.Enabled = False  ' Disable bookmark status menu
        mnusavepage.Enabled = False ' Disbale save webpage menu
        mnuping.Caption = "&Ping"
        mnushLnk.Enabled = False    ' Disable share bookmark menu button
        StatusBar1.Panels(1).Text = "Categorys (" & TableCount & ")   Bookmarks (" & BookCount & ")"
        Exit Sub
    Else
        RecoredName = Trim(tv1.SelectedItem.Key)    ' Get the name form the treeview control
        TvIndex = tv1.SelectedItem.Index    ' Get the index number form the treeview control
        ' Category bar
        CatBar.Buttons(1).Enabled = False    ' Disbale new cat button
        CatBar.Buttons(2).Enabled = True    ' Enable delete button
        CatBar.Buttons(3).Enabled = True    ' Enable rename table button
        CatBar.Buttons(4).Enabled = True    ' Enable Gen Html button
        ' Top toolbar
        tBar1.Buttons(1).ButtonMenus(1).Enabled = True ' Enable new bookmark button
        tBar1.Buttons(7).Enabled = False    ' Disable modify bookmark button
        tBar1.Buttons(8).Enabled = False    ' Disable delete bookmark button
        tBar1.Buttons(9).Enabled = False    ' Disable copy bookmark button
        tBar1.Buttons(11).Enabled = False   ' Disable move bookmark
        tBar1.Buttons(14).Enabled = False   ' Disable bookmark status button
        tBar1.Buttons(15).Enabled = False   ' Disable bookmark save web page button
        tBar1.Buttons(16).Enabled = False   ' Disable bookmark share link
        tBar1.Buttons(18).Enabled = False   ' Disable create short cut button
        tBar1.Buttons(21).Enabled = False
        
        ' Menu items
        mnunewcat.Enabled = False    ' Disable new cat button
        mnudelcat.Enabled = True    ' Enable delete cat button
        mnurename.Enabled = True    ' Enable rename cat button
        mnunewbook.Enabled = True   ' Enable new bookmark button
        mnucopy.Enabled = False     ' Disable the copy button
        mnumodbook.Enabled = False  ' Diable Modify bookmark button
        mnudelbook.Enabled = False  ' Disable delete bookmark button
        mnumove.Enabled = False     ' Disable moveto bookmark button
        mnucbk.Enabled = False      ' Disable create short cut button
        mnubkstats.Enabled = False  ' Disable bookmark status menu
        mnusavepage.Enabled = False ' Disbale save webpage menu
        mnushLnk.Enabled = False    ' Disable share bookmark menu button
        mnuping.Caption = "&Ping"
        LoadSites RecoredName       ' Load in the bookmarks
        StatusBar1.Panels(1).Text = "Bookmarks Found (" & BookCount & ")"
        RemoveSelection lstsites ' Removes the selection of the listview
    End If
    
End Sub

Private Sub tv1_KeyUp(KeyCode As Integer, Shift As Integer)
    tv1_Click
    If Len(RecoredName) <= 0 Then
        Exit Sub
    ElseIf KeyCode = 46 Then
        frmPopUpmenu.DelCat
    End If
    
End Sub


Private Sub tv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DbError Then Exit Sub
    If Not Button = vbRightButton Then Exit Sub
    If tv1.Nodes.Count <= 0 Then Exit Sub
    If tv1.SelectedItem.Index = 1 Then
        frmPopUpmenu.mnuaddcat.Enabled = True
        frmPopUpmenu.mnudelcat.Enabled = False
        frmPopUpmenu.mnupasscat.Enabled = False
        frmPopUpmenu.mnufindincat.Enabled = False
        frmPopUpmenu.mnufindincat.Caption = "&Find in Category"
        frmPopUpmenu.mnugenhtmlpage.Enabled = False
        PopupMenu frmPopUpmenu.mnucat
    Else
        tv1_Click
        frmPopUpmenu.mnufindincat.Enabled = True
        frmPopUpmenu.mnufindincat.Caption = "&Find in " & RecoredName
        frmPopUpmenu.mnuaddcat.Enabled = False
        frmPopUpmenu.mnudelcat.Enabled = True
        frmPopUpmenu.mnupasscat.Enabled = True
        frmPopUpmenu.mnufindcat.Enabled = True
        frmPopUpmenu.mnugenhtmlpage.Enabled = True
        PopupMenu frmPopUpmenu.mnucat
    End If
    
    
End Sub



Private Sub txtserach_Click()
    If Not UCase(txtserach.Tag) = "SR" Then
        Exit Sub
    Else
        txtserach.Tag = ""
        txtserach.Text = ""
    End If
    
End Sub

Private Sub txtserach_GotFocus()
    txtserach.BackColor = Config.Hightlight
End Sub

Private Sub txtserach_LostFocus()
    txtserach.BackColor = vbWhite
End Sub

Private Sub WebB_TitleChange(ByVal Text As String)
On Error Resume Next
Dim iStart As Long

    iStart = InStr(1, Text, "ref:RATE(", vbTextCompare)
    
    If Not iStart = 1 Then
        Exit Sub
    Else
        If Val(Mid(Text, iStart + 9, Len(Text) - iStart - 9)) = 0 Then
            frmRateing.lblrate(1).Caption = "Bookmark not yet Rated."
        Else
            frmRateing.lblrate(1).Caption = Mid(Text, iStart + 9, Len(Text) - iStart - 9)
        End If
        
        frmRateing.Show vbModal, frmmain
        WebB.Navigate GetTempFolder & "tmpxtygtp4.html"
    End If
End Sub



