VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList iLstGameSid 
      Left            =   6960
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "sid"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04DE
            Key             =   "game"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0594
            Key             =   "justgame"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLstColHeads 
      Left            =   7680
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":062E
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A80
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED2
            Key             =   "up"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   8280
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Start >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Height          =   135
      Left            =   120
      TabIndex        =   23
      Top             =   6600
      Width           =   9855
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   5
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   36
      Top             =   120
      Width           =   9855
      Begin VB.ListBox lstRenamed 
         Enabled         =   0   'False
         Height          =   2205
         Left            =   360
         TabIndex        =   59
         Top             =   3120
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.ListBox lstDupes 
         Enabled         =   0   'False
         Height          =   2205
         Left            =   360
         TabIndex        =   58
         Top             =   720
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Details..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblVerify 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   9015
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   4
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   28
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdCompare 
         Caption         =   "&Remove from List"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   56
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Set D&B Fields..."
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   55
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Play Sid &File"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   54
         Top             =   3720
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvCompare 
         Height          =   3255
         Left            =   0
         TabIndex        =   52
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iLstGameSid"
         SmallIcons      =   "iLstGameSid"
         ColHdrIcons     =   "iLstColHeads"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "GB64 Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GB64 Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sid Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "GB64 Musician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sid Author"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GB64 Year/Pub"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sid Copyright"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "GB64 DB Fields and Sid-File details comparison:"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   53
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   3
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   27
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdMissing 
         Caption         =   "Add &Selected..."
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   51
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdMissing 
         Caption         =   "&Remove"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   50
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdMissing 
         Caption         =   "Play Sid &File"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   49
         Top             =   3720
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvMissing 
         Height          =   3255
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iLstGameSid"
         SmallIcons      =   "iLstGameSid"
         ColHdrIcons     =   "iLstColHeads"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sid Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Author"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Copyright"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Sids not already in GB64 (these will be added to the ""Music"" table, not linked to any game):"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   7575
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   2
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   26
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdGamesNoSids 
         Caption         =   "&Remove game(s) from List"
         Height          =   375
         Index           =   3
         Left            =   7680
         TabIndex        =   46
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton cmdGamesNoSids 
         Caption         =   "&Link Sid (below) to Game (above)"
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   45
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdGamesNoSids 
         Caption         =   "Set D&B Fields..."
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   44
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdGamesNoSids 
         Caption         =   "Play Sid &File"
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   41
         Top             =   2640
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvGamesNoSids 
         Height          =   2175
         Left            =   0
         TabIndex        =   39
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iLstGameSid"
         SmallIcons      =   "iLstGameSid"
         ColHdrIcons     =   "iLstColHeads"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Game Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DB Musician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DB Group"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DB Nick"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DB Photo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DB Year"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DB Publisher"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvCache 
         Height          =   1935
         Left            =   0
         TabIndex        =   40
         Top             =   3120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iLstGameSid"
         SmallIcons      =   "iLstGameSid"
         ColHdrIcons     =   "iLstColHeads"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sid Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Author"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Copyright"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPrompt 
         Caption         =   "HVSC:"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   43
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Games with no Sid File:"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   1
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   25
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdInvalid 
         Caption         =   "&Re-link Selected"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   20
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvalid 
         Caption         =   "&Delete Selected"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   19
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvalid 
         Caption         =   "Set D&B Fields..."
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   18
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvalid 
         Caption         =   "Play Sid &File"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvalid 
         Caption         =   "&Pick Sid File..."
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   5880
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvInvalidLinks 
         Height          =   4095
         Left            =   0
         TabIndex        =   15
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iLstGameSid"
         SmallIcons      =   "iLstGameSid"
         ColHdrIcons     =   "iLstColHeads"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DB Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DB Musician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DB Group"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DB Nick"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DB Photo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Not Found"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Relocated File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Author (from sid file)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1324
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   35
         Top             =   120
         Width           =   6135
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   0
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   9855
      TabIndex        =   24
      Top             =   120
      Width           =   9855
      Begin VB.OptionButton optAction 
         Caption         =   "Clea&n Music Table"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   57
         ToolTipText     =   "Compare GB64 Musicians with sid-file authors"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.OptionButton optAction 
         Caption         =   "&Compare GB64/HVSC"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   12
         ToolTipText     =   "Compare GB64 Musicians with sid-file authors"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Add &missing sids"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   11
         ToolTipText     =   "Add sid-files that aren't already in the database"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Find &sids for games"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         ToolTipText     =   "Find sid-files for games that have no sid-file linked to them"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton optAction 
         Caption         =   "&Fix invalid links"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Find database entries that have an invalid sid-file linked to them"
         Top             =   4200
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdCache 
         Caption         =   "&Rebuild Cache File"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtPhotos 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   285
         Index           =   3
         Left            =   4920
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   285
         Index           =   2
         Left            =   4920
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtHVSC 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   2160
         Width           =   4815
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   285
         Index           =   1
         Left            =   4920
         TabIndex        =   3
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtSidplayer 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   0
         TabIndex        =   0
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label lblPrompt 
         Caption         =   $"frmMain.frx":13C0
         Height          =   615
         Index           =   5
         Left            =   0
         TabIndex        =   34
         Top             =   3240
         Width           =   5895
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Musician Photos folder:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   33
         Top             =   2520
         Width           =   5895
      End
      Begin VB.Label lblPrompt 
         Caption         =   "HVSC Root Folder (e.g. C:\C64Music):"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   32
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Sidplayer Executable File (e.g. sidplay2w.exe):"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Label lblPrompt 
         Caption         =   "GB64 Database File (.mdb file):"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label lblPrompt 
         Caption         =   "This tool will help you update your GB64 database with the latest High Voltage Sid Collection (HVSC)."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   6255
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   6840
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
' HVSC2GB64 - Tool to update GB64 database with
' the latest version of the HVSC
'----------------------------------------------
' This program is free software; you can
' redistribute it and/or modify it under the
' terms of the GNU General Public License as
' published by the Free Software Foundation;
' either version 2 of the License, or (at your
' option) any later version.
'
' This program is distributed in the hope that
' it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR
' PURPOSE. See the GNU General Public License
' for more details.
'
' You should have received a copy of the GNU
' General Public License along with this
' program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330,
' Boston, MA 02111-1307 USA
'==============================================

Option Explicit

Public m_blnQuittable As Boolean
Public m_lngState As Long

Private WithEvents m_xobjFinder As GBFuncs.clsFunctions
Attribute m_xobjFinder.VB_VarHelpID = -1

Private m_strLastPath As String

Private m_blnCacheFailed As Boolean

Private m_blnDoGVerify As Boolean
Private m_blnDoMVerify As Boolean

Private Type MusicTableType
    strFilename As String
    lngGAID As Long
    strName As String
End Type
Private Function DoCache() As String

    DoCache = ""
    m_strLastPath = ""
    Set m_xobjFinder = New GBFuncs.clsFunctions
    g_blnCacheError = False
    Call m_xobjFinder.FindFiles(g_strHVSCPathWBS, "*.sid")
    If g_blnCacheError = True Then
        DoCache = "The cache was built, but there was an error adding one or more of the sid files, so the cache may not be complete!"
    End If
    Set m_xobjFinder = Nothing
    Exit Function

End Function

Private Sub FindInvalidLinkedSids()

    Dim Litem As ListItem
    Dim strFile As String
    Dim strPath As String
    Dim lngIdx As Long
    Dim lngUBound As Long
    Dim strSIDName As String
    Dim strSIDAuthor As String
    Dim strSIDCopyright As String
    Dim lngSIDSongs As Long
    Dim lngSIDDefSong As Long
    Dim strUpper As String
    Dim strUpper2 As String

    Dim fbott As Long
    Dim ftmp As Long
    Dim ftop As Long
    Dim hit As Boolean

    'attempt to find relocated sids from the invalid list
    Screen.MousePointer = vbHourglass
    lblStatus = "Attempting to relocate moved sid files..."
    pbStatus.Min = 0
    pbStatus.Value = 0
    pbStatus.Max = lvInvalidLinks.ListItems.Count
    DoEvents
    lvInvalidLinks.Visible = False
    lvInvalidLinks.Enabled = False
    
    lngUBound = g_lngCacheTotal - 1
    With lvInvalidLinks
        For Each Litem In .ListItems
            Call g_xobjFunc.GetFileAndPath(Litem.SubItems(6), strPath, strFile)
            
            

            hit = False
            ftop = lngUBound
            fbott = 0
            ftmp = Int((ftop + fbott) / 2)
            strUpper = UCase$(strFile)
            Do While (Not hit) And (ftop >= fbott)
                lngIdx = ftmp
                With ga_udtCache(lngIdx)
                    strUpper2 = UCase$(.strFile)
                    If strUpper2 = strUpper Then
                        hit = True
                    ElseIf strUpper < strUpper2 Then
                        ftop = ftmp - 1
                    Else
                        fbott = ftmp + 1
                    End If
                    ftmp = Int((ftop + fbott) / 2)
                    If hit Then
                        Litem.SubItems(7) = .strPathFile
                        Litem.SubItems(8) = .strAuthor
                        Exit Do
                    End If
                End With
            Loop
            
            
            
            
'            For lngIdx = 0 To lngUBound
'                If ga_udtCache(lngIdx).strFile = strFile Then
'                    found it
'                    LItem.SubItems(7) = ga_udtCache(lngIdx).strPathFile
'                    LItem.SubItems(8) = ga_udtCache(lngIdx).strAuthor
'                    Exit For
'                End If
'            Next lngIdx
            
            
            
            
            
            pbStatus.Value = pbStatus.Value + 1
            DoEvents
        Next Litem
    End With
    pbStatus.Value = pbStatus.Max
    lvInvalidLinks.Sorted = False
    lvInvalidLinks.Sorted = True
    lvInvalidLinks.Visible = True
    lvInvalidLinks.Enabled = True

    Screen.MousePointer = vbNormal

End Sub

Private Function FixUnknownNamedSids(lngFixedNames_out As Long) As Long

    'Fixes all <?> named sids

    Dim rsGames As DAO.Recordset
    Dim fldGAID As DAO.Field
    Dim fldName As DAO.Field
    Dim fldSid As DAO.Field
    Dim strPathFile As String
    Dim strPath As String
    Dim strFile As String
    Dim lngIdx As Long
    Dim lngMax As Long
    Dim lngTotalGames As Long
    Dim lngTotalMusic As Long
    Dim a_strGames() As String
    Dim a_Music() As MusicTableType
    Dim strUpper As String
    Dim ftop As Long
    Dim fbott As Long
    Dim ftmp As Long
    Dim blnDel As Boolean
    Dim lngMusicIdx As Long
    Dim a_lngDupes() As Long
    Dim lngTotalDupes As Long
    Dim lngDupeIdx As Long
    Dim strDesiredName As String

    Screen.MousePointer = vbHourglass

    'cache the Music table
    lngTotalGames = 0
    ReDim a_strGames(0)
    Set rsGames = g_dbGB64.OpenRecordset("SELECT DISTINCTROW SidFilename FROM Games WHERE SidFilename <> """"", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            lblStatus = "Caching Games table..."
            pbStatus.Value = 0
            DoEvents
            Set fldSid = !SidFileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            lngTotalGames = .RecordCount
            ReDim a_strGames(lngTotalGames - 1)
            lngIdx = 0
            Do Until .EOF
                a_strGames(lngIdx) = UCase$(fldSid)
                With pbStatus
                    .Value = .Value + 1
                    DoEvents
                End With
                lngIdx = lngIdx + 1
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldSid = Nothing
        End If
        .Close
    End With
    Set rsGames = Nothing

    'cache the Music table
    lngTotalMusic = 0
    ReDim a_Music(0)
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Filename, GA_Id, Name FROM Music WHERE Filename <> """"", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            lblStatus = "Caching Music table..."
            pbStatus.Value = 0
            DoEvents
            Set fldGAID = !GA_Id
            Set fldSid = !FileName
            Set fldName = !Name
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            lngTotalMusic = .RecordCount
            ReDim a_Music(lngTotalMusic - 1)
            lngIdx = 0
            Do Until .EOF
                a_Music(lngIdx).strFilename = UCase$(fldSid)
                a_Music(lngIdx).lngGAID = fldGAID
                a_Music(lngIdx).strName = fldName
                With pbStatus
                    .Value = .Value + 1
                    DoEvents
                End With
                lngIdx = lngIdx + 1
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldSid = Nothing
            Set fldGAID = Nothing
            Set fldName = Nothing
        End If
        .Close
    End With
    Set rsGames = Nothing
    
    'quicksort arrays
    lblStatus = "Sorting Games table cache..."
    DoEvents
    QuickSortStringArray 0, lngTotalGames - 1, a_strGames()
    lblStatus = "Sorting Music table cache..."
    DoEvents
    QuickSortMusicTableArray 0, lngTotalMusic - 1, a_Music()
  
    'loop through hvsc cache and check each sid
    lngFixedNames_out = 0
    lngMax = g_lngCacheTotal - 1
    lblStatus = "Cleaning Music table..."
    pbStatus.Min = 0
    pbStatus.Value = 0
    pbStatus.Max = g_lngCacheTotal
    DoEvents
    For lngIdx = 0 To lngMax
        blnDel = False
        With ga_udtCache(lngIdx)
            'find the sid in the music table
            strUpper = UCase$(.strPathFile)
            ftop = lngTotalMusic - 1
            fbott = 0
            Do While (fbott <= ftop)
                ftmp = Int((fbott + ftop) / 2)
                If a_Music(ftmp).strFilename = strUpper Then
                    
                    'found the sid in the Music table
                    lngMusicIdx = ftmp
                    
                    'is it in the games table?
                    ftop = lngTotalGames - 1
                    fbott = 0
                    Do While (fbott <= ftop)
                        ftmp = Int((fbott + ftop) / 2)
                        If a_strGames(ftmp) = strUpper Then
                            
                            'found the sid in the Music table
                            blnDel = True
                            Exit Do
        
                        ElseIf strUpper < a_strGames(ftmp) Then
                            ftop = ftmp - 1
                        Else
                            fbott = ftmp + 1
                        End If
                    Loop
                    
                    'now check if it's a duplicate in the music table
                    lngTotalDupes = 0
                    ReDim a_lngDupes(0)
                    'before in the list
                    If lngMusicIdx > 0 Then
                        lngDupeIdx = lngMusicIdx - 1
                        Do While lngDupeIdx > -1
                            If a_Music(lngDupeIdx).strFilename = strUpper Then
                                ReDim Preserve a_lngDupes(lngTotalDupes)
                                a_lngDupes(lngTotalDupes) = a_Music(lngDupeIdx).lngGAID
                                lngTotalDupes = lngTotalDupes + 1
                                lngDupeIdx = lngDupeIdx - 1
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                    'after in the list
                    If lngMusicIdx < (lngTotalMusic - 1) Then
                        lngDupeIdx = lngMusicIdx + 1
                        Do While lngDupeIdx < lngTotalMusic
                            If a_Music(lngDupeIdx).strFilename = strUpper Then
                                ReDim Preserve a_lngDupes(lngTotalDupes)
                                a_lngDupes(lngTotalDupes) = a_Music(lngDupeIdx).lngGAID
                                lngTotalDupes = lngTotalDupes + 1
                                lngDupeIdx = lngDupeIdx + 1
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'do we need to remove any dupes?
                    If lngTotalDupes > 0 Then
                        lstDupes.AddItem .strName & " (" & .strPathFile & ") [" & lngTotalDupes & " dupes removed]"
                        For lngDupeIdx = 0 To lngTotalDupes - 1
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT GA_Id FROM Music WHERE GA_Id=" & a_lngDupes(lngDupeIdx), dbOpenDynaset)
                            With rsGames
                                .MoveFirst
                                .Delete
                                .Close
                            End With
                            Set rsGames = Nothing
                        Next lngDupeIdx
                    End If
                    
                    'delete the sid (it's in the games table)?
                    If blnDel Then
                        lstDupes.AddItem .strName & " (" & .strPathFile & ") [in games table]"
                        Set rsGames = g_dbGB64.OpenRecordset("SELECT GA_Id FROM Music WHERE GA_Id=" & a_Music(lngMusicIdx).lngGAID, dbOpenDynaset)
                        With rsGames
                            .MoveFirst
                            .Delete
                            .Close
                        End With
                        Set rsGames = Nothing
                    Else
                        'rename to that of the sid-name in the hvsc
                        If .strName = "<?>" Then
                            strDesiredName = Replace(.strFile, "_", " ")
                            If Right$(UCase$(strDesiredName), 4) = ".SID" Then
                                strDesiredName = Left$(strDesiredName, Len(strDesiredName) - 4)
                            End If
                            strDesiredName = "<" & strDesiredName & ">"
                        Else
                            strDesiredName = .strName
                        End If
                        If a_Music(lngMusicIdx).strName <> strDesiredName Then
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT Name FROM Music WHERE GA_Id=" & a_Music(lngMusicIdx).lngGAID, dbOpenDynaset)
                            With rsGames
                                .MoveFirst
                                .Edit
                                !Name = strDesiredName
                                .Update
                                .Close
                            End With
                            Set rsGames = Nothing
                            lstRenamed.AddItem "Renamed '" & a_Music(lngMusicIdx).strName & "' ==> '" & strDesiredName & "'  [" & .strPathFile & "]"
                            lngFixedNames_out = lngFixedNames_out + 1
                        End If
                    End If

                    'next sid in the cache
                    Exit Do

                ElseIf strUpper < a_Music(ftmp).strFilename Then
                    ftop = ftmp - 1
                Else
                    fbott = ftmp + 1
                End If
            Loop

        End With

        With pbStatus
            .Value = .Value + 1
            DoEvents
        End With

    Next lngIdx
  
  
  
    FixUnknownNamedSids = lstDupes.ListCount
  


    Screen.MousePointer = vbNormal



'    FixUnknownNamedSids = 0
'    Set rsGames = g_dbGB64.OpenRecordset("SELECT Name, Filename FROM Music WHERE Name = ""<?>"" AND Filename <> """"", dbOpenDynaset)
'    With rsGames
'        If Not .EOF Then
'            Screen.MousePointer = vbHourglass
'            lblStatus = "Fixing <?> sids..."
'            pbStatus.Value = 0
'            DoEvents
'            Set fldSid = !FileName
'            Set fldName = !Name
'            .MoveLast
'            .MoveFirst
'            pbStatus.Min = 0
'            pbStatus.Max = .RecordCount + 1
'            FixUnknownNamedSids = .RecordCount
'            Do Until .EOF
'                strPathFile = fldSid
'                Call g_xobjFunc.GetFileAndPath(strPathFile, strPath, strFile)
'                strFile = Trim$(strFile)
'                If Right$(UCase$(strFile), 4) = ".SID" Then
'                    strFile = Left$(strFile, Len(strFile) - 4)
'                    .Edit
'                    fldName = "<" & Replace(strFile, "_", " ") & ">"
'                    .Update
'                End If
'                With pbStatus
'                    .Value = .Value + 1
'                    DoEvents
'                End With
'                .MoveNext
'            Loop
'            pbStatus.Value = pbStatus.Max
'            Set fldSid = Nothing
'            Set fldName = Nothing
'            Screen.MousePointer = vbNormal
'        End If
'        .Close
'    End With
'    Set rsGames = Nothing

End Function

Private Sub GetListViewSettings()

    Dim lngSortOrder As Long
    Dim lngSortKey As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long

    'lvCompare
    With g_xobjReg
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareSK", lngSortKey, False, 0 'col 1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareSO", lngSortOrder, False, 0 'asc
        For lngIdx = 0 To 7
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareW" & Trim$(lngIdx), lngWidth, False, 2000 'col width
            lvCompare.ColumnHeaders(lngIdx + 1).Width = lngWidth
        Next lngIdx
        For lngIdx = 0 To 7
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareP" & Trim$(lngIdx), lngPos, False, lngIdx + 1 'col pos
            lvCompare.ColumnHeaders(lngIdx + 1).Position = lngPos
        Next lngIdx
    End With
    If lngSortKey < 0 Or lngSortKey > 8 Then
        lngSortKey = 0
    End If
    If lngSortOrder < 0 Or lngSortOrder > 1 Then
        lngSortOrder = 0
    End If
    lvCompare.Sorted = False
    lvCompare.SortKey = lngSortKey
    lvCompare.SortOrder = lngSortOrder
    lvCompare.Sorted = True
    Call SetSortIconsForCompare

    'lvGamesNoSids
    With g_xobjReg
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsSK", lngSortKey, False, 0 'col 1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsSO", lngSortOrder, False, 0 'asc
        For lngIdx = 0 To 6
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsW" & Trim$(lngIdx), lngWidth, False, 2000 'col width
            lvGamesNoSids.ColumnHeaders(lngIdx + 1).Width = lngWidth
        Next lngIdx
        For lngIdx = 0 To 6
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsP" & Trim$(lngIdx), lngPos, False, lngIdx + 1 'col pos
            lvGamesNoSids.ColumnHeaders(lngIdx + 1).Position = lngPos
        Next lngIdx
    End With
    If lngSortKey < 0 Or lngSortKey > 6 Then
        lngSortKey = 0
    End If
    If lngSortOrder < 0 Or lngSortOrder > 1 Then
        lngSortOrder = 0
    End If
    lvGamesNoSids.Sorted = False
    lvGamesNoSids.SortKey = lngSortKey
    lvGamesNoSids.SortOrder = lngSortOrder
    lvGamesNoSids.Sorted = True
    Call SetSortIconsForGamesNoSids

    'lvCache
    With g_xobjReg
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheSK", lngSortKey, False, 0 'col 1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheSO", lngSortOrder, False, 0 'asc
        For lngIdx = 0 To 4
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheW" & Trim$(lngIdx), lngWidth, False, 2000 'col width
            lvCache.ColumnHeaders(lngIdx + 1).Width = lngWidth
        Next lngIdx
        For lngIdx = 0 To 4
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheP" & Trim$(lngIdx), lngPos, False, lngIdx + 1 'col pos
            lvCache.ColumnHeaders(lngIdx + 1).Position = lngPos
        Next lngIdx
    End With
    If lngSortKey < 0 Or lngSortKey > 4 Then
        lngSortKey = 0
    End If
    If lngSortOrder < 0 Or lngSortOrder > 1 Then
        lngSortOrder = 0
    End If
    lvCache.Sorted = False
    lvCache.SortKey = lngSortKey
    lvCache.SortOrder = lngSortOrder
    lvCache.Sorted = True
    Call SetSortIconsForCache

    'lvInvalidLinks
    With g_xobjReg
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidSK", lngSortKey, False, 0 'col 1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidSO", lngSortOrder, False, 0 'asc
        For lngIdx = 0 To 8
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidW" & Trim$(lngIdx), lngWidth, False, 2000 'col width
            lvInvalidLinks.ColumnHeaders(lngIdx + 1).Width = lngWidth
        Next lngIdx
        For lngIdx = 0 To 8
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidP" & Trim$(lngIdx), lngPos, False, lngIdx + 1 'col pos
            lvInvalidLinks.ColumnHeaders(lngIdx + 1).Position = lngPos
        Next lngIdx
    End With
    If lngSortKey < 0 Or lngSortKey > 8 Then
        lngSortKey = 0
    End If
    If lngSortOrder < 0 Or lngSortOrder > 1 Then
        lngSortOrder = 0
    End If
    lvInvalidLinks.Sorted = False
    lvInvalidLinks.SortKey = lngSortKey
    lvInvalidLinks.SortOrder = lngSortOrder
    lvInvalidLinks.Sorted = True
    Call SetSortIconsForInvalid
    
    'lvMissing
    With g_xobjReg
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingSK", lngSortKey, False, 0 'col 1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingSO", lngSortOrder, False, 0 'asc
        For lngIdx = 0 To 4
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingW" & Trim$(lngIdx), lngWidth, False, 2000 'col width
            lvMissing.ColumnHeaders(lngIdx + 1).Width = lngWidth
        Next lngIdx
        For lngIdx = 0 To 4
            .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingP" & Trim$(lngIdx), lngPos, False, lngIdx + 1 'col pos
            lvMissing.ColumnHeaders(lngIdx + 1).Position = lngPos
        Next lngIdx
    End With
    If lngSortKey < 0 Or lngSortKey > 4 Then
        lngSortKey = 0
    End If
    If lngSortOrder < 0 Or lngSortOrder > 1 Then
        lngSortOrder = 0
    End If
    lvMissing.Sorted = False
    lvMissing.SortKey = lngSortKey
    lvMissing.SortOrder = lngSortOrder
    lvMissing.Sorted = True
    Call SetSortIconsForMissing



End Sub

Private Function GetNewSidFile(strFile_in As String) As String

    Dim strPath As String
    Dim strFile As String
    Dim strLastDir As String

    On Error GoTo Exit_Here

    With cDialog
        .CancelError = True
        .DefaultExt = "sid"
        .DialogTitle = "Pick SID File..."
        .Filter = "SID;Sid Files"
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        If strFile = "" Then
            g_xobjReg.GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "LastPickSidPath", strLastDir, True, g_strHVSCPath
            If g_xobjFunc.DirExists(strLastDir) Then
                .InitDir = strLastDir
            Else
                .InitDir = g_strHVSCPath
            End If
            .FileName = ""
        Else
            Call g_xobjFunc.GetFileAndPath(strFile_in, strPath, strFile)
            .FileName = strFile
            .InitDir = strPath
        End If
        
        .ShowOpen

        If Left$(UCase$(.FileName), Len(g_strHVSCPathWBS)) = UCase$(g_strHVSCPathWBS) Then
            g_xobjFunc.GetFileAndPath .FileName, strPath, strFile
            g_xobjReg.SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "LastPickSidPath", strPath, True
            GetNewSidFile = Right$(.FileName, Len(.FileName) - Len(g_strHVSCPathWBS))
        Else
            MsgBox "Please select a SID file within the HVSC folders!", vbExclamation, gc_strAPPNAME
            GetNewSidFile = ""
        End If
        
    End With

    Exit Function

Exit_Here:
    GetNewSidFile = ""

End Function

Private Sub IntroStartCheck()

    cmdNext.Enabled = ((Trim$(txtDatabase.Text) <> "") And _
                      (Trim$(txtSidplayer.Text) <> "") And _
                      (Trim$(txtHVSC.Text) <> "") And _
                      (Trim$(txtPhotos.Text) <> ""))
                
End Sub


Private Function OpenAndCheckDatabase(strDBFile_in) As String

    'return error message on fail or "" on success

    Dim rsGames As DAO.Recordset
    Dim lngYearIdx As Long

    On Error GoTo Err_Handler

    'open the database
    Set g_dbGB64 = OpenDatabase(strDBFile_in)

    'grab the database version info
    Set rsGames = g_dbGB64.OpenRecordset("Config", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            .MoveFirst
            g_lngDBMajorVersion = !MajorVersion
            g_lngDBMinorVersion = !MinorVersion
            .Close
        Else
            .Close
            g_dbGB64.Close
            OpenAndCheckDatabase = "Missing Version/Update info in Config Table!"
            Exit Function
        End If
    End With
    Set rsGames = Nothing
    g_sngDBVersion = g_lngDBMajorVersion + (CSng(g_lngDBMinorVersion / 10))

    'update the database if needed
    If g_sngDBVersion < gc_sngREQ_DB_VERSION Then
        OpenAndCheckDatabase = "Database version is too low.  Update your database with the latest version of the GameBase Frontend!"
        Exit Function
    ElseIf g_sngDBVersion > gc_sngREQ_DB_VERSION Then
        OpenAndCheckDatabase = "Database version is too high.  There may be a newer version of this tool on the GB64 website!"
        Exit Function
    End If

    'read in the years lookup table
    'years field
    ga_strYearLookup(1) = "????"
    ga_strYearLookup(2) = "19??"
    ga_strYearLookup(3) = "197?"
    ga_strYearLookup(4) = "198?"
    ga_strYearLookup(5) = "199?"
    ga_strYearLookup(6) = "20??"
    ga_strYearLookup(7) = "200?"
    ga_strYearLookup(8) = "201?"
    ga_strYearLookup(9) = "202?"
    For lngYearIdx = 10 To gc_lngYEARLOOKUPMAX
        ga_strYearLookup(lngYearIdx) = (gc_lngEARLIESTYEAR - 10) + lngYearIdx
    Next lngYearIdx

    OpenAndCheckDatabase = ""
    Exit Function

Err_Handler:
    OpenAndCheckDatabase = "There was an error whilst opening the specified database file: -" & vbCrLf & vbCrLf & Err.Description & "  [Error " & Err.Number & "," & Err.LastDllError & "]"

End Function

Private Sub PlaySid(strFilename_in, ByVal blnShowError_in As Boolean)

    On Error Resume Next

    If g_xobjFunc.FileExists(g_strSidplayerExe, False, "") = False Then
        If blnShowError_in Then
            MsgBox "Sidplayer not found: """ & g_strSidplayerExe & """", vbExclamation, gc_strAPPNAME
        End If
    Else
        If g_xobjFunc.FileExists(strFilename_in, False, "") = False Then
            If blnShowError_in Then
                MsgBox "Sid File not found: """ & strFilename_in & """", vbExclamation, gc_strAPPNAME
            End If
        Else
            Shell g_strSidplayerExe & " """ & strFilename_in & """", vbNormalNoFocus
        End If
    End If

End Sub

Private Function ReadInCache() As String

    'read in the cache file

    Dim intFN As Integer
    Dim strFile As String
    Dim a_strTemp() As String
    Dim lngUBound As Long
    Dim lngLineIdx As Long
    Dim strLine As String
    
    On Error GoTo Err_Handler
    
    'init
    g_lngCacheTotal = 0
    ReDim ga_udtCache(0)
    lblStatus = "Preparing to load the cache file..."
    lblStatus.Visible = True
    pbStatus.Min = 0
    pbStatus.Value = 0
    pbStatus.Max = 1
    pbStatus.Visible = True
    DoEvents

    'open the file
    intFN = FreeFile()
    Open (g_strAppPathWithBackSlash & "cache.txt") For Input As #intFN

    'read in the first line of the cache file and get the total number of lines/sids
    Line Input #intFN, strLine
    g_lngCacheTotal = CLng(Trim$(strLine))
    If g_lngCacheTotal < 1 Then
        ReadInCache = "There was an error finding the number of sids in the cache file.  Possibly corrupt file."
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    'read in each line
    lngUBound = g_lngCacheTotal
    g_lngCacheTotal = g_lngCacheTotal + 1
    ReDim ga_udtCache(lngUBound)
    pbStatus.Max = g_lngCacheTotal
    lblStatus = "Loading the cache file..."
    DoEvents
    lngLineIdx = 0
    Do Until EOF(intFN)
        Line Input #intFN, strLine
        strLine = Trim$(strLine)
        If Len(strLine) > 0 Then
            a_strTemp = Split(strLine, gc_strDELIM)
            If UBound(a_strTemp) = 7 Then
                With ga_udtCache(lngLineIdx)
                    .strPathFile = a_strTemp(0)
                    .strFile = a_strTemp(1)
                    .strPath = a_strTemp(2)
                    .strName = a_strTemp(3)
                    .strAuthor = a_strTemp(4)
                    .strCopyright = a_strTemp(5)
                    .strSongs = a_strTemp(6)
                    .strDefSong = a_strTemp(7)
                    .lngID = lngLineIdx
                    .blnFlag = False
                End With
                lngLineIdx = lngLineIdx + 1
            End If
            pbStatus.Value = pbStatus.Value + 1
        End If
    Loop
    Close #intFN
    pbStatus.Value = pbStatus.Max

    If lngLineIdx < 100 Then
        ReadInCache = "There was an error loading the cache file.  Possibly corrupt file."
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    'adjust the num sid files
    ReDim Preserve ga_udtCache(lngLineIdx - 1)
    g_lngCacheTotal = lngLineIdx

    lblStatus = "Sorting..."
    DoEvents
    If optAction(0).Value = True Then
        'filename only sort needed
        QuickSortCacheArrayByFile 0, g_lngCacheTotal - 1
    Else
        'pathfilename sort needed
        QuickSortCacheArray 0, g_lngCacheTotal - 1
    End If

    ReadInCache = ""
    Exit Function

Err_Handler:
    ReadInCache = "There was an error whilst reading the cache file of your HVSC: -" & vbCrLf & vbCrLf & Err.Description & "  [Error " & Err.Number & "," & Err.LastDllError & "]"
    Screen.MousePointer = vbNormal

End Function

Private Function ReLinkInvalidSids() As String

    Dim Litem As ListItem
    Dim lngCount As Long
    Dim lngSelCount As Long
    Dim rsGames As DAO.Recordset
    Dim lngID As Long
    Dim blnGame As Boolean
    Dim a_lngRemove() As Long
    Dim lngIdx As Long

    On Error GoTo Err_Handler

    Screen.MousePointer = vbHourglass
    DoEvents
    lngCount = 0
    lngSelCount = 0
    ReDim a_lngRemove(0)
    For Each Litem In lvInvalidLinks.ListItems
        With Litem
            If .Selected = True Then
                lngSelCount = lngSelCount + 1
                If .SubItems(7) <> "" Then
                    blnGame = (Left$(.Key, 1) = "G")
                    lngID = Right$(.Key, Len(.Key) - 1)
                    If blnGame Then
                        Set rsGames = g_dbGB64.OpenRecordset("SELECT SidFilename FROM Games WHERE GA_ID=" & lngID, dbOpenDynaset)
                        m_blnDoGVerify = True
                    Else
                        Set rsGames = g_dbGB64.OpenRecordset("SELECT Filename AS SidFilename FROM Music WHERE GA_ID=" & lngID, dbOpenDynaset)
                        m_blnDoMVerify = True
                    End If
                    If Not rsGames.EOF Then
                        rsGames.MoveFirst
                        rsGames.Edit
                        rsGames!SidFileName = .SubItems(7)
                        rsGames.Update
                        ReDim Preserve a_lngRemove(lngCount)
                        a_lngRemove(lngCount) = Litem.Index
                        lngCount = lngCount + 1
                    End If
                    rsGames.Close
                    Set rsGames = Nothing
                End If
            End If
        End With
    Next Litem

    'remove all items
    With lvInvalidLinks
        .Enabled = False
        .Visible = False
        With .ListItems
            For lngIdx = (lngCount - 1) To 0 Step -1
                .Remove a_lngRemove(lngIdx)
            Next lngIdx
        End With
        .Enabled = True
        .Visible = True
    End With

    If lngCount < lngSelCount Then
        ReLinkInvalidSids = "Re-linked " & Trim$(lngCount) & " out of " & Trim$(lngSelCount) & " selected items." & vbCrLf & vbCrLf & "Items that do not have a relocated sid-file linked to them were skipped."
    Else
        ReLinkInvalidSids = ""
    End If

    Screen.MousePointer = vbNormal
    Exit Function

Err_Handler:
    ReLinkInvalidSids = "There was an error whilst re-linking the selected items: -" & vbCrLf & vbCrLf & Err.Description & "  [Error " & Err.Number & "," & Err.LastDllError & "]"
    Screen.MousePointer = vbNormal

End Function
Private Function DeleteInvalidSids() As String

    Dim Litem As ListItem
    Dim lngCount As Long
    Dim lngSelCount As Long
    Dim rsGames As DAO.Recordset
    Dim lngID As Long
    Dim blnGame As Boolean
    Dim a_lngRemove() As Long
    Dim lngIdx As Long

    On Error GoTo Err_Handler
    
    Screen.MousePointer = vbHourglass
    DoEvents
    lngCount = 0
    lngSelCount = 0
    ReDim a_lngRemove(0)
    For Each Litem In lvInvalidLinks.ListItems
        With Litem
            If .Selected = True Then
                lngSelCount = lngSelCount + 1
                blnGame = (Left$(.Key, 1) = "G")
                lngID = Right$(.Key, Len(.Key) - 1)
                If Not blnGame Then
                    Set rsGames = g_dbGB64.OpenRecordset("SELECT Filename FROM Music WHERE GA_ID=" & lngID, dbOpenDynaset)
                    If Not rsGames.EOF Then
                        rsGames.MoveFirst
                        rsGames.Delete
                        ReDim Preserve a_lngRemove(lngCount)
                        a_lngRemove(lngCount) = Litem.Index
                        lngCount = lngCount + 1
                        m_blnDoMVerify = True
                    End If
                    rsGames.Close
                    Set rsGames = Nothing
                End If
            End If
        End With
    Next Litem

    'remove all items
    With lvInvalidLinks
        .Enabled = False
        .Visible = False
        With .ListItems
            For lngIdx = (lngCount - 1) To 0 Step -1
                .Remove a_lngRemove(lngIdx)
            Next lngIdx
        End With
        .Enabled = True
        .Visible = True
    End With

    If lngCount < lngSelCount Then
        DeleteInvalidSids = "Deleted " & Trim$(lngCount) & " out of " & Trim$(lngSelCount) & " selected items." & vbCrLf & vbCrLf & "All Games were skipped.  Only Sids can be deleted."
    Else
        DeleteInvalidSids = ""
    End If

    Screen.MousePointer = vbNormal
    Exit Function

Err_Handler:
    DeleteInvalidSids = "There was an error whilst deleting the selected items: -" & vbCrLf & vbCrLf & Err.Description & "  [Error " & Err.Number & "," & Err.LastDllError & "]"
    Screen.MousePointer = vbNormal

End Function

Private Sub RemoveSelected(lv_in As ListView)

    Dim Litem As ListItem
    Dim lngTotal As Long
    Dim lngMax As Long
    Dim lngIdx As Long
    Dim a_remove() As Long

    Screen.MousePointer = vbHourglass
    DoEvents
    Enabled = False
    ReDim a_remove(0)
    lngTotal = 0

    With lv_in
        .Enabled = False
        .Visible = False
        For Each Litem In .ListItems
            With Litem
                If .Selected Then
                    ReDim Preserve a_remove(lngTotal)
                    a_remove(lngTotal) = .Index
                    lngTotal = lngTotal + 1
                End If
            End With
        Next Litem
        lngMax = lngTotal - 1
        For lngIdx = lngMax To 0 Step -1
            .ListItems.Remove a_remove(lngIdx)
        Next lngIdx
        .Enabled = True
        .Visible = True
    End With
    
    Enabled = True
    Screen.MousePointer = vbNormal

End Sub

Private Function SetSortIconsForInvalid()

    Dim lngIdx As Long

    With lvInvalidLinks
        'set the sort icon on the sorted column
        If .SortOrder = lvwAscending Then
            .ColumnHeaders(.SortKey + 1).Icon = "down"
        Else
            .ColumnHeaders(.SortKey + 1).Icon = "up"
        End If
        'blank out the others
        For lngIdx = 0 To 7
            If lngIdx <> .SortKey Then
                .ColumnHeaders(lngIdx + 1).Icon = "blank"
            End If
        Next lngIdx
    End With

End Function
Private Function SetSortIconsForCache()

    Dim lngIdx As Long

    With lvCache
        'set the sort icon on the sorted column
        If .SortOrder = lvwAscending Then
            .ColumnHeaders(.SortKey + 1).Icon = "down"
        Else
            .ColumnHeaders(.SortKey + 1).Icon = "up"
        End If
        'blank out the others
        For lngIdx = 0 To 4
            If lngIdx <> .SortKey Then
                .ColumnHeaders(lngIdx + 1).Icon = "blank"
            End If
        Next lngIdx
    End With

End Function
Private Function SetSortIconsForMissing()

    Dim lngIdx As Long

    With lvMissing
        'set the sort icon on the sorted column
        If .SortOrder = lvwAscending Then
            .ColumnHeaders(.SortKey + 1).Icon = "down"
        Else
            .ColumnHeaders(.SortKey + 1).Icon = "up"
        End If
        'blank out the others
        For lngIdx = 0 To 4
            If lngIdx <> .SortKey Then
                .ColumnHeaders(lngIdx + 1).Icon = "blank"
            End If
        Next lngIdx
    End With

End Function
Private Function SetSortIconsForCompare()

    Dim lngIdx As Long

    With lvCompare
        'set the sort icon on the sorted column
        If .SortOrder = lvwAscending Then
            .ColumnHeaders(.SortKey + 1).Icon = "down"
        Else
            .ColumnHeaders(.SortKey + 1).Icon = "up"
        End If
        'blank out the others
        For lngIdx = 0 To 7
            If lngIdx <> .SortKey Then
                .ColumnHeaders(lngIdx + 1).Icon = "blank"
            End If
        Next lngIdx
    End With

End Function
Private Function SetSortIconsForGamesNoSids()

    Dim lngIdx As Long

    With lvGamesNoSids
        'set the sort icon on the sorted column
        If .SortOrder = lvwAscending Then
            .ColumnHeaders(.SortKey + 1).Icon = "down"
        Else
            .ColumnHeaders(.SortKey + 1).Icon = "up"
        End If
        'blank out the others
        For lngIdx = 0 To 6
            If lngIdx <> .SortKey Then
                .ColumnHeaders(lngIdx + 1).Icon = "blank"
            End If
        Next lngIdx
    End With

End Function
Private Function VerifyAvailableFiles() As String

    Dim fldName As Field
    Dim fldFilename As Field
    Dim fldID As Field
    Dim fldCCode As Field
    Dim fldFA As Field
    Dim fldSA As Field
    Dim blnFound As Boolean
    Dim strFile As String
    Dim strFile2 As String
    Dim blnBlankFilename As Boolean
    Dim lngMissingSids As Long
    Dim lngTotalSids As Long
    Dim intFileNum As Integer
    Dim rsGames As DAO.Recordset

    On Error GoTo Failed

    'create the file
    intFileNum = FreeFile()
    If g_xobjFunc.FileExists(g_strAppPathWithBackSlash & "verify.txt", False, "") Then
        Kill g_strAppPathWithBackSlash & "verify.txt"
    End If
    Open g_strAppPathWithBackSlash & "verify.txt" For Output As #intFileNum
    Print #intFileNum, Format$(Now, "dddd dd mmmm yyyy Hh:mm:ss") & vbCrLf & g_strDBFile & vbCrLf & vbCrLf

    'verify sids
    If m_blnDoGVerify = True Then
        strFile = ""
        lngMissingSids = 0
        lngTotalSids = 0
        'check games table
        lblStatus = "Verifying sids in Games table..."
        Set rsGames = g_dbGB64.OpenRecordset("SELECT GA_Id, Name, SidFilename, CCode, FA, SA FROM Games ORDER BY Name", dbOpenDynaset)
        With rsGames
            If Not .EOF Then
                Set fldName = !Name
                Set fldFilename = !SidFileName
                Set fldID = !GA_Id
                Set fldCCode = !CCode
                Set fldFA = !FA
                Set fldSA = !SA
                .MoveLast
                .MoveFirst
                pbStatus.Min = 0
                pbStatus.Max = .RecordCount
                .MoveFirst
                Do Until .EOF
                    'find the file
                    blnFound = False
                    blnBlankFilename = False
                    If fldFilename <> "" Then
                        lngTotalSids = lngTotalSids + 1
                        If g_xobjFunc.FileExists(g_strHVSCPathWBS & fldFilename, False, "") Then
                            blnFound = True
                        End If
                    Else
                        blnBlankFilename = True
                    End If
                    'update the record
                    If blnFound Then
                        .Edit
                        fldSA = True
                        If fldFA = True Then
                            fldCCode = 1 'both exist
                        Else
                            fldCCode = 3 'only sid exists
                        End If
                        .Update
                    Else
                        .Edit
                        fldSA = False
                        If fldFA = True Then
                            fldCCode = 2 'only game exists
                        Else
                            fldCCode = 4 'none exist
                        End If
                        .Update
                        If Not blnBlankFilename Then
                            strFile = strFile & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                            lngMissingSids = lngMissingSids + 1
                        End If
                    End If
                    'update status and move to next record
                    pbStatus.Value = .AbsolutePosition
                    .MoveNext
                    DoEvents
                Loop
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldName = Nothing
                Set fldFilename = Nothing
                Set fldID = Nothing
                Set fldCCode = Nothing
                Set fldFA = Nothing
                Set fldSA = Nothing
            End If
        End With
        rsGames.Close
        Set rsGames = Nothing
        If strFile <> "" Then
            strFile = "Games table..." & vbCrLf & vbCrLf & strFile
        End If
    End If

    If m_blnDoMVerify = True Then

        'check music table
        strFile2 = ""
        lblStatus = "Verifying tunes in Music table..."
        Set rsGames = g_dbGB64.OpenRecordset("SELECT GA_Id, Name, Filename, SA FROM Music ORDER BY Name", dbOpenDynaset)
        With rsGames
            If Not .EOF Then
                Set fldName = !Name
                Set fldFilename = !FileName
                Set fldID = !GA_Id
                Set fldSA = !SA
                .MoveLast
                .MoveFirst
                pbStatus.Min = 0
                pbStatus.Max = .RecordCount
                lngTotalSids = lngTotalSids + .RecordCount
                .MoveFirst
                Do Until .EOF
                    'find the file
                    blnFound = False
                    blnBlankFilename = False
                    If fldFilename <> "" Then
                        If g_xobjFunc.FileExists(g_strHVSCPathWBS & fldFilename, False, "") Then
                            blnFound = True
                        End If
                    Else
                        blnBlankFilename = True
                    End If
                    'update the record
                    If blnFound Then
                        .Edit
                        fldSA = True
                        .Update
                    Else
                        .Edit
                        fldSA = False
                        .Update
                        If Not blnBlankFilename Then
                            strFile2 = strFile2 & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                        End If
                        lngMissingSids = lngMissingSids + 1
                    End If
                    'update status and move to next record
                    pbStatus.Value = .AbsolutePosition
                    .MoveNext
                    DoEvents
                Loop
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldName = Nothing
                Set fldFilename = Nothing
                Set fldID = Nothing
                Set fldSA = Nothing
            End If
        End With
        rsGames.Close
        Set rsGames = Nothing
        If strFile2 <> "" Then
            strFile2 = "Music table..." & vbCrLf & vbCrLf & strFile2
        End If

    End If

    'write to file
    If lngMissingSids > 0 Then
        If strFile <> "" And strFile2 <> "" Then
            Print #intFileNum, "Sids not found: " & lngMissingSids & "/" & lngTotalSids & vbCrLf & g_strHVSCPathWBS & vbCrLf & vbCrLf & strFile & vbCrLf & strFile2 & vbCrLf & vbCrLf
        ElseIf strFile <> "" Then
            Print #intFileNum, "Sids not found: " & lngMissingSids & "/" & lngTotalSids & vbCrLf & g_strHVSCPathWBS & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf
        Else
            Print #intFileNum, "Sids not found: " & lngMissingSids & "/" & lngTotalSids & vbCrLf & g_strHVSCPathWBS & vbCrLf & vbCrLf & strFile2 & vbCrLf & vbCrLf
        End If
    Else
        Print #intFileNum, "Sids not found: " & lngMissingSids & "/" & lngTotalSids & vbCrLf & vbCrLf & vbCrLf
    End If

    Close #intFileNum

    lblVerify = "Verify complete.  Found " & Trim$(lngTotalSids - lngMissingSids) & "/" & Trim$(lngTotalSids) & ".  Missing: " & Trim$(lngMissingSids)
    If m_blnDoGVerify = True And m_blnDoMVerify = True Then
        lblVerify = lblVerify & "  [Games and Music tables checked]"
    ElseIf m_blnDoGVerify = True Then
        lblVerify = lblVerify & "  [Only Games table checked]"
    Else
        lblVerify = lblVerify & "  [Only Music table checked]"
    End If
    lblVerify.Enabled = True
    lblVerify.Visible = True
    If lngMissingSids > 0 Then
        cmdVerify.Enabled = True
        cmdVerify.Visible = True
    End If
    
    VerifyAvailableFiles = ""
    Exit Function

Failed:
    VerifyAvailableFiles = "There was an error verifying the linked sid files: -" & vbCrLf & vbCrLf & Err.Description & "  (Err " & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Sub cmdAbort_Click()

    If MsgBox("Are you sure you want to quit?", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdCache_Click()

    'rebuild the cache

    Dim lngIdx As Long
    Dim strError As String

    With g_xobjFunc
        If Not .DirExists(Trim$(txtHVSC)) Then
            MsgBox "The specified HVSC Root folder was not found!", vbInformation, gc_strAPPNAME
            txtHVSC.SetFocus
            Exit Sub
        End If
        If (Mid$(Trim$(txtHVSC), 2, 1) <> ":") And (Mid$(Trim$(txtHVSC), 3, 1) <> "\") Then
            MsgBox "Absolute paths only please!", vbInformation, gc_strAPPNAME
            txtHVSC.SetFocus
            Exit Sub
        End If
    End With

    'set the paths
    g_strHVSCPath = Trim$(txtHVSC)
    g_strHVSCPathWBS = g_strHVSCPath
    If Right$(g_strHVSCPathWBS, 1) <> "\" Then
        g_strHVSCPathWBS = g_strHVSCPathWBS & "\"
        g_lngLenHVSCPathWBS = Len(g_strHVSCPathWBS)
    End If

    'disable the controls
    For lngIdx = 0 To 3
        cmdBrowse(lngIdx).Enabled = False
    Next lngIdx
    txtDatabase.Enabled = False
    txtSidplayer.Enabled = False
    txtHVSC.Enabled = False
    txtPhotos.Enabled = False
    cmdNext.Enabled = False
    cmdAbort.Enabled = False
    lblStatus.Visible = True
    picWiz(0).Enabled = False
    cmdCache.Enabled = False

    'run the cache builder
    MsgBox "About to create the cache file.  This could take a while, depending on the speed of your computer.", vbInformation, gc_strAPPNAME
    m_blnQuittable = False
    m_blnCacheFailed = False
    strError = BuildCachedDatabase()
    m_blnQuittable = True
    If strError <> "" Then
        m_blnCacheFailed = True
        MsgBox strError, vbCritical, gc_strAPPNAME
    Else
        MsgBox "Cache rebuilt successfully!" & vbCrLf & vbCrLf & "Total Sids added to cache: " & g_lngTotalSids, vbInformation, gc_strAPPNAME
    End If

    'enable the controls
    For lngIdx = 0 To 3
        cmdBrowse(lngIdx).Enabled = True
    Next lngIdx
    txtDatabase.Enabled = True
    txtSidplayer.Enabled = True
    txtHVSC.Enabled = True
    txtPhotos.Enabled = True
    cmdAbort.Enabled = True
    lblStatus.Visible = False
    picWiz(0).Enabled = True
    cmdCache.Enabled = True
    IntroStartCheck

End Sub

Private Sub cmdCompare_Click(Index As Integer)

    Dim blnMultiple As Boolean
    Dim Litem As ListItem
    Dim strFile As String
    Dim strError As String
    Dim lngGAID As Long
    Dim lngOldID As Long
    Dim rsGames As DAO.Recordset
    Dim blnGame As Boolean
    
    Set Litem = lvCompare.SelectedItem

    Select Case Index
        Case 0
            'play sid
            If Not Litem Is Nothing Then
                strFile = g_strHVSCPathWBS & Litem.SubItems(7)
                Call PlaySid(strFile, True)
            Else
                MsgBox "Please select a single item from the list.", vbInformation, gc_strAPPNAME
            End If
            lvCompare.SetFocus

        Case 1
            'set db fields
            If AtLeastOneItemSelectedInListView(lvCompare, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list.", vbInformation, gc_strAPPNAME
            Else
                With g_udtRecordSelector
                    .blnOKClicked = False
                    .strSQL = "SELECT MU_Id as ID, Musician as Name FROM Musicians ORDER BY Musician"
                    If blnMultiple Then
                        .lngID = 39
                    Else
                        .lngID = CLng(Litem.Tag)
                    End If
                    .blnGNPChanged = False
                    lngOldID = .lngID
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    frmRecordSelector.Show vbModal, Me
                    Screen.MousePointer = vbDefault
                End With
                If g_udtRecordSelector.blnOKClicked Then
                    'save new musician for selected games/sids
                    If Not blnMultiple Then
                        blnGame = (Left$(Litem.Key, 1) = "G")
                        lngGAID = CLng(Right$(Litem.Key, Len(Litem.Key) - 1))
                        If blnGame Then
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        Else
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        End If
                        With rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                !MU_Id = g_udtRecordSelector.lngID
                                .Update
                                Litem.SubItems(3) = g_udtRecordSelector.strReturnedName
                                Litem.Tag = Trim$(g_udtRecordSelector.lngID)
                            Else
                                MsgBox "Database record not found - Musician not updated for selected item!", vbCritical, gc_strAPPNAME
                            End If
                            .Close
                        End With
                        Set rsGames = Nothing
                    Else
                        Screen.MousePointer = vbHourglass
                        DoEvents
                        lvCompare.Enabled = False
                        lvCompare.Visible = False
                        Enabled = False
                        For Each Litem In lvCompare.ListItems
                            With Litem
                                If .Selected Then
                                    blnGame = (Left$(.Key, 1) = "G")
                                    lngGAID = CLng(Right$(.Key, Len(.Key) - 1))
                                    If blnGame Then
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    Else
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    End If
                                    If Not rsGames.EOF Then
                                        rsGames.MoveFirst
                                        rsGames.Edit
                                        rsGames!MU_Id = g_udtRecordSelector.lngID
                                        rsGames.Update
                                        .SubItems(3) = g_udtRecordSelector.strReturnedName
                                        .Tag = Trim$(g_udtRecordSelector.lngID)
                                    Else
                                        MsgBox "Database record not found - Musician not updated for selected item [" & .SubItems(1) & "]!", vbCritical, gc_strAPPNAME
                                    End If
                                    rsGames.Close
                                    Set rsGames = Nothing
                                End If
                            End With
                        Next Litem
                        lvCompare.Enabled = True
                        lvCompare.Visible = True
                        Enabled = True
                        Screen.MousePointer = vbNormal
                    End If
                End If
            End If
            lvCompare.SetFocus
        Case 2
            'remove selected
            If AtLeastOneItemSelectedInListView(lvCompare, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list to remove.", vbInformation, gc_strAPPNAME
            Else
                Call RemoveSelected(lvCompare)
            End If
            lvCompare.SetFocus
    End Select
    
End Sub

Private Sub cmdGamesNoSids_Click(Index As Integer)

    Dim blnMultiple As Boolean
    Dim Litem As ListItem
    Dim LItem2 As ListItem
    Dim strFile As String
    Dim strError As String
    Dim blnGame As Boolean
    Dim lngOldID As Long
    Dim lngGAID As Long
    Dim rsGames As DAO.Recordset
    
    Set Litem = lvGamesNoSids.SelectedItem
    Set LItem2 = lvCache.SelectedItem

    Select Case Index
        Case 0
            'play sid
            If Not Litem Is Nothing Then
                strFile = g_strHVSCPathWBS & LItem2.SubItems(1) & "\" & LItem2.SubItems(2)
                Call PlaySid(strFile, True)
            Else
                MsgBox "Please select a single item from the bottom (HVSC) list.", vbInformation, gc_strAPPNAME
            End If
            lvCache.SetFocus

        Case 1
            'set db fields
            If AtLeastOneItemSelectedInListView(lvGamesNoSids, blnMultiple) = False Then
                MsgBox "Please select at least one item from the top (games) list.", vbInformation, gc_strAPPNAME
            Else
                With g_udtRecordSelector
                    .blnOKClicked = False
                    .strSQL = "SELECT MU_Id as ID, Musician as Name FROM Musicians ORDER BY Musician"
                    If blnMultiple Then
                        .lngID = 39
                    Else
                        .lngID = CLng(Litem.Tag)
                    End If
                    .blnGNPChanged = False
                    lngOldID = .lngID
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    frmRecordSelector.Show vbModal, Me
                    Screen.MousePointer = vbDefault
                End With
                If g_udtRecordSelector.blnOKClicked Then
                    'save new musician for selected games/sids
                    If Not blnMultiple Then
                        blnGame = (Left$(Litem.Key, 1) = "G")
                        lngGAID = CLng(Right$(Litem.Key, Len(Litem.Key) - 1))
                        If blnGame Then
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        Else
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        End If
                        With rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                !MU_Id = g_udtRecordSelector.lngID
                                .Update
                                Litem.SubItems(1) = g_udtRecordSelector.strReturnedName
                                Litem.SubItems(2) = g_udtRecordSelector.strGroup
                                Litem.SubItems(3) = g_udtRecordSelector.strNick
                                Litem.SubItems(4) = g_udtRecordSelector.strPhoto
                                Litem.Tag = Trim$(g_udtRecordSelector.lngID)
                                Litem.Ghosted = (g_udtRecordSelector.lngID = 39)
                            Else
                                MsgBox "Database record not found - Musician not updated for selected item!", vbCritical, gc_strAPPNAME
                            End If
                            .Close
                        End With
                        Set rsGames = Nothing
                    Else
                        Screen.MousePointer = vbHourglass
                        DoEvents
                        lvGamesNoSids.Enabled = False
                        lvGamesNoSids.Visible = False
                        Enabled = False
                        For Each Litem In lvGamesNoSids.ListItems
                            With Litem
                                If .Selected Then
                                    blnGame = (Left$(.Key, 1) = "G")
                                    lngGAID = CLng(Right$(.Key, Len(.Key) - 1))
                                    If blnGame Then
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    Else
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    End If
                                    If Not rsGames.EOF Then
                                        rsGames.MoveFirst
                                        rsGames.Edit
                                        rsGames!MU_Id = g_udtRecordSelector.lngID
                                        rsGames.Update
                                        .SubItems(2) = g_udtRecordSelector.strReturnedName
                                        .Tag = Trim$(g_udtRecordSelector.lngID)
                                    Else
                                        MsgBox "Database record not found - Musician not updated for selected item [" & .SubItems(1) & "]!", vbCritical, gc_strAPPNAME
                                    End If
                                    rsGames.Close
                                    Set rsGames = Nothing
                                End If
                            End With
                        Next Litem
                        If g_udtRecordSelector.blnGNPChanged Then
                            For Each Litem In lvGamesNoSids.ListItems
                                With Litem
                                    If .Tag = g_udtRecordSelector.lngID Then
                                        .SubItems(1) = g_udtRecordSelector.strReturnedName
                                        .SubItems(2) = g_udtRecordSelector.strGroup
                                        .SubItems(3) = g_udtRecordSelector.strNick
                                        .SubItems(4) = g_udtRecordSelector.strPhoto
                                        .Tag = Trim$(g_udtRecordSelector.lngID)
                                        .Ghosted = (g_udtRecordSelector.lngID = 39)
                                    End If
                                End With
                            Next Litem
                        End If
                        lvGamesNoSids.Enabled = True
                        lvGamesNoSids.Visible = True
                        Enabled = True
                        Screen.MousePointer = vbNormal
                    End If
                End If
            End If
            lvGamesNoSids.SetFocus

        Case 2

            're-link selected
            If LItem2 Is Nothing Then
                MsgBox "Please select one item from the bottom (hvsc) list.", vbInformation, gc_strAPPNAME
            Else
                If AtLeastOneItemSelectedInListView(lvGamesNoSids, blnMultiple) = False Then
                    MsgBox "Please select at least one item from the top (games) list.", vbInformation, gc_strAPPNAME
                Else
                    If Not blnMultiple Then
                        
                        If Litem Is Nothing Then
                            MsgBox "Please properly select at least one item from the top (games) list.", vbInformation, gc_strAPPNAME
                        Else
                            'link the selected game
                            lngGAID = CLng(Right$(Litem.Key, Len(Litem.Key) - 1))
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT SidFilename FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                            With rsGames
                                If Not .EOF Then
                                    .MoveFirst
                                    .Edit
                                    !SidFileName = LItem2.SubItems(1) & "\" & LItem2.SubItems(2)
                                    .Update
                                    'remove the selected game
                                    Call lvGamesNoSids.ListItems.Remove(Litem.Index)
                                    m_blnDoGVerify = True
                                Else
                                    MsgBox "Database record not found - Sid file not updated for selected item!", vbCritical, gc_strAPPNAME
                                End If
                                .Close
                            End With
                            Set rsGames = Nothing
                        End If
                    Else
                        If MsgBox("Link the selected sid to ALL the selected games?", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbYes Then
                            'link the selected games
                            Screen.MousePointer = vbHourglass
                            DoEvents
                            lvGamesNoSids.Enabled = False
                            lvGamesNoSids.Visible = False
                            Enabled = False
                            For Each Litem In lvGamesNoSids.ListItems
                                With Litem
                                    If .Selected Then
                                        lngGAID = CLng(Right$(.Key, Len(.Key) - 1))
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT SidFilename FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                        If Not rsGames.EOF Then
                                            rsGames.MoveFirst
                                            rsGames.Edit
                                            rsGames!SidFileName = LItem2.SubItems(1) & "\" & LItem2.SubItems(2)
                                            rsGames.Update
                                        End If
                                        rsGames.Close
                                        Set rsGames = Nothing
                                    End If
                                End With
                            Next Litem
                            m_blnDoGVerify = True
                            'remove the selected games from the list
                            Call RemoveSelected(lvGamesNoSids)
                            lvGamesNoSids.Enabled = True
                            lvGamesNoSids.Visible = True
                            Enabled = True
                            Screen.MousePointer = vbNormal
                        End If
                    End If
                    If lvGamesNoSids.ListItems.Count = 0 Then
                        cmdNext.SetFocus
                    Else
                        lvGamesNoSids.SetFocus
                    End If
                End If
            End If
            lvGamesNoSids.SetFocus
        
        Case 3
            'remove selected
            If AtLeastOneItemSelectedInListView(lvGamesNoSids, blnMultiple) = False Then
                MsgBox "Please select at least one item from the top (games) list to remove.", vbInformation, gc_strAPPNAME
            Else
                If MsgBox("Are you sure you want to remove all selected items from the top (games) list (this won't delete them from the database)?", vbYesNo Or vbQuestion, gc_strAPPNAME) = vbYes Then
                    Call RemoveSelected(lvGamesNoSids)
                End If
            End If
            lvGamesNoSids.SetFocus

    End Select

End Sub

Private Sub cmdInvalid_Click(Index As Integer)

    Dim blnMultiple As Boolean
    Dim Litem As ListItem
    Dim strFile As String
    Dim strError As String
    Dim lngOldID As Long
    Dim rsGames As DAO.Recordset
    Dim lngGAID As String
    Dim blnGame As Boolean
    
    Set Litem = lvInvalidLinks.SelectedItem

    Select Case Index
        Case 0
            'pick sid file
            If AtLeastOneItemSelectedInListView(lvInvalidLinks, blnMultiple) = False Then
                MsgBox "Please select a single item from the list.", vbInformation, gc_strAPPNAME
            Else
                If blnMultiple Then
                    MsgBox "Multi-select not allowed for this option." & vbCrLf & vbCrLf & "Please select only a SINGLE item from the list to set its Sid file.", vbInformation, gc_strAPPNAME
                Else
                    If Not Litem Is Nothing Then
                        If Trim$(Litem.SubItems(7)) <> "" Then
                            strFile = g_strHVSCPathWBS & Litem.SubItems(7)
                        Else
                            strFile = ""
                        End If
                        strFile = GetNewSidFile(strFile)
                        If strFile <> "" Then
                            'update the listview
                            Litem.SubItems(7) = strFile
                        End If
                    Else
                        MsgBox "Please select a single item from the list.", vbInformation, gc_strAPPNAME
                    End If
                End If
            End If
            lvInvalidLinks.SetFocus
            
        Case 1
            'play sid
            If AtLeastOneItemSelectedInListView(lvInvalidLinks, blnMultiple) = False Then
                MsgBox "Please select a single item from the list that has a valid sid file linked to it.", vbInformation, gc_strAPPNAME
            Else
                If blnMultiple Then
                    MsgBox "Multi-select not allowed for this option." & vbCrLf & vbCrLf & "Please select only a SINGLE item from the list that has a valid sid file linked to it.", vbInformation, gc_strAPPNAME
                Else
                    If Not Litem Is Nothing Then
                        strFile = g_strHVSCPathWBS & Litem.SubItems(7)
                        If Trim$(Litem.SubItems(7)) <> "" Then
                            Call PlaySid(strFile, True)
                        End If
                    Else
                        MsgBox "Please select a single item from the list that has a valid sid file linked to it.", vbInformation, gc_strAPPNAME
                    End If
                End If
            End If
            lvInvalidLinks.SetFocus
            
        Case 2
            'set db fields
            If AtLeastOneItemSelectedInListView(lvInvalidLinks, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list.", vbInformation, gc_strAPPNAME
            Else
                With g_udtRecordSelector
                    .blnOKClicked = False
                    .strSQL = "SELECT MU_Id as ID, Musician as Name FROM Musicians ORDER BY Musician"
                    If blnMultiple Then
                        .lngID = 39
                    Else
                        .lngID = CLng(Litem.Tag)
                    End If
                    .blnGNPChanged = False
                    lngOldID = .lngID
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    frmRecordSelector.Show vbModal, Me
                    Screen.MousePointer = vbDefault
                End With
                If g_udtRecordSelector.blnOKClicked Then
                    'save new musician for selected games/sids
                    If Not blnMultiple Then
                        blnGame = (Left$(Litem.Key, 1) = "G")
                        lngGAID = CLng(Right$(Litem.Key, Len(Litem.Key) - 1))
                        If blnGame Then
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        Else
                            Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                        End If
                        With rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                !MU_Id = g_udtRecordSelector.lngID
                                .Update
                                Litem.SubItems(2) = g_udtRecordSelector.strReturnedName
                                Litem.SubItems(3) = g_udtRecordSelector.strGroup
                                Litem.SubItems(4) = g_udtRecordSelector.strNick
                                Litem.SubItems(5) = g_udtRecordSelector.strPhoto
                                Litem.Tag = Trim$(g_udtRecordSelector.lngID)
                            Else
                                MsgBox "Database record not found - Musician not updated for selected item!", vbCritical, gc_strAPPNAME
                            End If
                            .Close
                        End With
                        Set rsGames = Nothing
                    Else
                        Screen.MousePointer = vbHourglass
                        DoEvents
                        lvInvalidLinks.Enabled = False
                        lvInvalidLinks.Visible = False
                        Enabled = False
                        For Each Litem In lvInvalidLinks.ListItems
                            With Litem
                                If .Selected Then
                                    blnGame = (Left$(.Key, 1) = "G")
                                    lngGAID = CLng(Right$(.Key, Len(.Key) - 1))
                                    If blnGame Then
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Games WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    Else
                                        Set rsGames = g_dbGB64.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id=" & lngGAID, dbOpenDynaset)
                                    End If
                                    If Not rsGames.EOF Then
                                        rsGames.MoveFirst
                                        rsGames.Edit
                                        rsGames!MU_Id = g_udtRecordSelector.lngID
                                        rsGames.Update
                                        .SubItems(2) = g_udtRecordSelector.strReturnedName
                                        .Tag = Trim$(g_udtRecordSelector.lngID)
                                    Else
                                        MsgBox "Database record not found - Musician not updated for selected item [" & .SubItems(1) & "]!", vbCritical, gc_strAPPNAME
                                    End If
                                    rsGames.Close
                                    Set rsGames = Nothing
                                End If
                            End With
                        Next Litem
                        If g_udtRecordSelector.blnGNPChanged Then
                            For Each Litem In lvInvalidLinks.ListItems
                                With Litem
                                    If .Tag = g_udtRecordSelector.lngID Then
                                        .SubItems(2) = g_udtRecordSelector.strReturnedName
                                        .SubItems(3) = g_udtRecordSelector.strGroup
                                        .SubItems(4) = g_udtRecordSelector.strNick
                                        .SubItems(5) = g_udtRecordSelector.strPhoto
                                        .Tag = Trim$(g_udtRecordSelector.lngID)
                                    End If
                                End With
                            Next Litem
                        End If
                        lvInvalidLinks.Enabled = True
                        lvInvalidLinks.Visible = True
                        Enabled = True
                        Screen.MousePointer = vbNormal
                    End If
                End If
            End If
            lvInvalidLinks.SetFocus

        Case 3
            'delete selected
            If AtLeastOneItemSelectedInListView(lvInvalidLinks, blnMultiple) = False Then
                MsgBox "Please select at least one ""Sid"" item from the list (games cannot be deleted).", vbInformation, gc_strAPPNAME
            Else
                If MsgBox("Are you sure you want to permanently delete the selected sids (not games) from the database?", vbQuestion Or vbYesNo) = vbYes Then
                    strError = DeleteInvalidSids()
                    If strError <> "" Then
                        MsgBox strError, vbExclamation, gc_strAPPNAME
                    End If
                    If lvInvalidLinks.ListItems.Count = 0 Then
                        cmdNext.SetFocus
                    Else
                        lvInvalidLinks.SetFocus
                    End If
                End If
            End If
            lvInvalidLinks.SetFocus

        Case 4
            're-link selected
            If AtLeastOneItemSelectedInListView(lvInvalidLinks, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list who's sid file has been relocated.", vbInformation, gc_strAPPNAME
                lvInvalidLinks.SetFocus
            Else
                strError = ReLinkInvalidSids()
                If strError <> "" Then
                    MsgBox strError, vbExclamation, gc_strAPPNAME
                End If
                If lvInvalidLinks.ListItems.Count = 0 Then
                    cmdNext.SetFocus
                Else
                    lvInvalidLinks.SetFocus
                End If
            End If
            
    End Select

End Sub

Private Sub cmdMissing_Click(Index As Integer)

    Dim blnMultiple As Boolean
    Dim Litem As ListItem
    Dim strFile As String
    Dim strError As String
    Dim lngGAID As Long
    Dim lngOldID As Long
    Dim rsGames As DAO.Recordset
    
    Set Litem = lvMissing.SelectedItem

    Select Case Index
        Case 0
            'play sid
            If Not Litem Is Nothing Then
                strFile = g_strHVSCPathWBS & Litem.SubItems(1) & "\" & Litem.SubItems(2)
                Call PlaySid(strFile, True)
            Else
                MsgBox "Please select a single item from the list.", vbInformation, gc_strAPPNAME
            End If
            lvMissing.SetFocus

        Case 1
            'remove selected
            If AtLeastOneItemSelectedInListView(lvMissing, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list to remove.", vbInformation, gc_strAPPNAME
            Else
                If MsgBox("Are you sure you want to remove all selected items from the list (and NOT add them to GB64's Music table?", vbYesNo Or vbQuestion, gc_strAPPNAME) = vbYes Then
                    Call RemoveSelected(lvMissing)
                End If
            End If
            lvMissing.SetFocus
 
        Case 2
            'add selected (and set gb64 musician fields)
            If AtLeastOneItemSelectedInListView(lvMissing, blnMultiple) = False Then
                MsgBox "Please select at least one item from the list.", vbInformation, gc_strAPPNAME
            Else
                With g_udtRecordSelector
                    .blnOKClicked = False
                    .strSQL = "SELECT MU_Id as ID, Musician as Name FROM Musicians ORDER BY Musician"
                    .lngID = 39
                    .blnGNPChanged = False
                    lngOldID = .lngID
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    frmRecordSelector.Show vbModal, Me
                    Screen.MousePointer = vbDefault
                End With
                If g_udtRecordSelector.blnOKClicked Then
                    If Not blnMultiple Then
                        'add the sid to the music table
                        Set rsGames = g_dbGB64.OpenRecordset("SELECT * FROM Music", dbOpenDynaset)
                        With rsGames
                            .AddNew
                            !Name = Left$(Litem.Text, 255)
                            !FileName = Litem.SubItems(1) & "\" & Litem.SubItems(2)
                            !MU_Id = g_udtRecordSelector.lngID
                            !SFav = False
                            !SA = False
                            !Adult = False
                            .Update
                            .Close
                        End With
                        Set rsGames = Nothing
                        Call lvMissing.ListItems.Remove(Litem.Index)
                        m_blnDoMVerify = True
                    Else
                        'add all selected sids to the music table
                        Screen.MousePointer = vbHourglass
                        DoEvents
                        lvMissing.Enabled = False
                        lvMissing.Visible = False
                        Enabled = False
                        Set rsGames = g_dbGB64.OpenRecordset("SELECT * FROM Music", dbOpenDynaset)
                        For Each Litem In lvMissing.ListItems
                            With Litem
                                If .Selected Then
                                    rsGames.AddNew
                                    rsGames!Name = Left$(Litem.Text, 255)
                                    rsGames!FileName = .SubItems(1) & "\" & Litem.SubItems(2)
                                    rsGames!MU_Id = g_udtRecordSelector.lngID
                                    rsGames!SFav = False
                                    rsGames!SA = False
                                    rsGames!Adult = False
                                    rsGames.Update
                                End If
                            End With
                        Next Litem
                        rsGames.Close
                        Set rsGames = Nothing
                        Call RemoveSelected(lvMissing)
                        lvMissing.Enabled = True
                        lvMissing.Visible = True
                        Enabled = True
                        Screen.MousePointer = vbNormal
                        m_blnDoMVerify = True
                    End If
                End If
            End If
            lvMissing.SetFocus

    End Select

End Sub
Private Sub cmdNext_Click()

    Dim strError As String
    Dim lngIdx As Long
    Dim lngTotalInvalid As Long
    Dim lngTotalGamesNoSids As Long
    Dim lngTotalMissing As Long
    Dim lngTotalCompare As Long
    Dim blnDoVerify As Boolean
    Dim blnVN As Boolean
    Dim lngTotalFixed As Long
    Dim lngTotalFixedNames As Long

    If m_lngState = 0 Then
    
        'get the paths etc
        With g_xobjFunc
            If Not .FileExists(Trim$(txtDatabase), False, "") Then
                MsgBox "The specified GB64 Database file was not found!", vbInformation, gc_strAPPNAME
                txtDatabase.SetFocus
                Exit Sub
            End If
            If Not .FileExists(Trim$(txtSidplayer), False, "") Then
                MsgBox "The specified Sidplayer file was not found!", vbInformation, gc_strAPPNAME
                txtSidplayer.SetFocus
                Exit Sub
            End If
            If Not .DirExists(Trim$(txtHVSC)) Then
                MsgBox "The specified HVSC Root folder was not found!", vbInformation, gc_strAPPNAME
                txtHVSC.SetFocus
                Exit Sub
            End If
            If Not .DirExists(Trim$(txtPhotos)) Then
                MsgBox "The specified Musician Photos folder was not found!", vbInformation, gc_strAPPNAME
                txtPhotos.SetFocus
                Exit Sub
            End If
            If (Mid$(Trim$(txtHVSC), 2, 1) <> ":") And (Mid$(Trim$(txtHVSC), 3, 1) <> "\") Then
                MsgBox "Absolute paths only please!", vbInformation, gc_strAPPNAME
                txtHVSC.SetFocus
                Exit Sub
            End If
            If (Mid$(Trim$(txtPhotos), 2, 1) <> ":") And (Mid$(Trim$(txtPhotos), 3, 1) <> "\") Then
                MsgBox "Absolute paths only please!", vbInformation, gc_strAPPNAME
                txtPhotos.SetFocus
                Exit Sub
            End If
            If (Mid$(Trim$(txtSidplayer), 2, 1) <> ":") And (Mid$(Trim$(txtSidplayer), 3, 1) <> "\") Then
                MsgBox "Absolute paths only please!", vbInformation, gc_strAPPNAME
                txtSidplayer.SetFocus
                Exit Sub
            End If
            If (Mid$(Trim$(txtDatabase), 2, 1) <> ":") And (Mid$(Trim$(txtDatabase), 3, 1) <> "\") Then
                MsgBox "Absolute paths only please!", vbInformation, gc_strAPPNAME
                txtDatabase.SetFocus
                Exit Sub
            End If
            strError = OpenAndCheckDatabase(Trim$(txtDatabase))
            If strError <> "" Then
                MsgBox strError, vbInformation, gc_strAPPNAME
                txtDatabase.SetFocus
                Exit Sub
            End If
        End With

        'set the paths
        g_strPhotoPath = Trim$(txtPhotos)
        g_strDBFile = Trim$(txtDatabase)
        g_strHVSCPath = Trim$(txtHVSC)
        g_strHVSCPathWBS = g_strHVSCPath
        If Right$(g_strHVSCPathWBS, 1) <> "\" Then
            g_strHVSCPathWBS = g_strHVSCPathWBS & "\"
            g_lngLenHVSCPathWBS = Len(g_strHVSCPathWBS)
        End If
        g_strSidplayerExe = Trim$(txtSidplayer)

        'save these to the registry so we don't piss the user off :)
        With g_xobjReg
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Database", g_strDBFile, True
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Sidplayer", g_strSidplayerExe, True
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "HVSC", g_strHVSCPath, True
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Photos", g_strPhotoPath, True
        End With

        'do we need to rebuilt the cache?
        
        If g_xobjFunc.FileExists(g_strAppPathWithBackSlash & "cache.txt", False, "") = False Then
            cmdCache.Value = True
            If m_blnCacheFailed = True Then
                Exit Sub
            End If
        End If
        
        'disable the controls
        For lngIdx = 0 To 3
            cmdBrowse(lngIdx).Enabled = False
        Next lngIdx
        txtDatabase.Enabled = False
        txtSidplayer.Enabled = False
        txtHVSC.Enabled = False
        txtPhotos.Enabled = False
        cmdNext.Enabled = False
        cmdAbort.Enabled = False
        optAction(0).Enabled = False
        optAction(1).Enabled = False
        optAction(2).Enabled = False
        optAction(3).Enabled = False
        optAction(4).Enabled = False
        picWiz(0).Enabled = False
        cmdCache.Enabled = False
        
        ' if here, we're ready to read in the cache
        strError = ReadInCache()
        If strError <> "" Then
            MsgBox strError, vbCritical, gc_strAPPNAME
            For lngIdx = 0 To 3
                cmdBrowse(lngIdx).Enabled = True
            Next lngIdx
            txtDatabase.Enabled = True
            txtSidplayer.Enabled = True
            txtHVSC.Enabled = True
            txtPhotos.Enabled = True
            cmdNext.Enabled = True
            cmdAbort.Enabled = True
            optAction(0).Enabled = True
            optAction(1).Enabled = True
            optAction(2).Enabled = True
            optAction(3).Enabled = True
            optAction(4).Enabled = True
            picWiz(0).Enabled = True
            cmdCache.Enabled = True
        End If

        If optAction(0).Value = True Then

            'finding the missing invalid files
            m_blnQuittable = False
            lngTotalInvalid = FillInvalidLinkedSidsList()
            If lngTotalInvalid > 0 Then
                Call FindInvalidLinkedSids
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                'lvInvalidLinks.SetFocus
            Else
                lvInvalidLinks.Enabled = False
                For lngIdx = 0 To 4
                    cmdInvalid(lngIdx).Enabled = False
                Next lngIdx
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                'cmdNext.SetFocus
            End If
            pbStatus.Visible = False
            lblStatus.Visible = False
            m_blnQuittable = True
            m_lngState = 1
            picWiz(1).ZOrder 0
        
        ElseIf optAction(1).Value = True Then
        
            'manually find games that have no game file
            m_blnQuittable = False
            lngTotalGamesNoSids = FillGamesWithNoSidsList()
            If lngTotalGamesNoSids > 0 Then
                Call FillHVSCCacheList
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                lvGamesNoSids.SetFocus
            Else
                lvGamesNoSids.Enabled = False
                For lngIdx = 0 To 2
                    cmdGamesNoSids(lngIdx).Enabled = False
                Next lngIdx
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                cmdNext.SetFocus
            End If
            pbStatus.Visible = False
            lblStatus.Visible = False
            m_blnQuittable = True
            m_lngState = 1
            picWiz(2).ZOrder 0
        
        ElseIf optAction(2).Value = True Then
        
            'add missing sids to the music table
            m_blnQuittable = False
            lngTotalMissing = FillMissingSidsList()
            If lngTotalMissing > 0 Then
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                lvMissing.SetFocus
            Else
                lvMissing.Enabled = False
                For lngIdx = 0 To 2
                    cmdMissing(lngIdx).Enabled = False
                Next lngIdx
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                cmdNext.SetFocus
            End If
            pbStatus.Visible = False
            lblStatus.Visible = False
            m_blnQuittable = True
            m_lngState = 1
            picWiz(3).ZOrder 0
        
        ElseIf optAction(3).Value = True Then
        
            'compare gb64 musician to sid-author
            m_blnQuittable = False
            lngTotalCompare = FillCompareList()
            If lngTotalCompare > 0 Then
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                lvCompare.SetFocus
            Else
                lvCompare.Enabled = False
                For lngIdx = 0 To 1
                    cmdCompare(lngIdx).Enabled = False
                Next lngIdx
                cmdAbort.Enabled = True
                cmdNext.Enabled = True
                cmdNext.Caption = "&Done"
                cmdNext.SetFocus
            End If
            pbStatus.Visible = False
            lblStatus.Visible = False
            m_blnQuittable = True
            m_lngState = 1
            picWiz(4).ZOrder 0

        ElseIf optAction(4).Value = True Then
        
            'fix <?> sids
            m_blnQuittable = False
            lngTotalFixed = FixUnknownNamedSids(lngTotalFixedNames)
            m_blnQuittable = True
            Screen.MousePointer = vbNormal
            cmdAbort.Enabled = False
            cmdAbort.Visible = False
            pbStatus.Visible = False
            lblStatus.Visible = False
            cmdNext.Enabled = True
            cmdNext.Caption = "&Close"
            m_lngState = 2
            lblVerify.Visible = True
            lblVerify = "Finished.  " & lngTotalFixed & " duplicate sids removed (" & lngTotalFixedNames & " sid-names fixed)"
            picWiz(5).ZOrder 0
            If lngTotalFixed > 0 Then
                lstDupes.Enabled = True
                lstDupes.Visible = True
                lstDupes.SetFocus
            End If
            If lngTotalFixedNames > 0 Then
                lstRenamed.Enabled = True
                lstRenamed.Visible = True
                If lngTotalFixed = 0 Then
                    lstRenamed.Top = lstDupes.Top
                End If
            End If
            m_blnDoGVerify = False
            m_blnDoMVerify = False

        End If
    ElseIf m_lngState = 1 Then
        'finished - do the verify if we need to
        blnDoVerify = False
        blnVN = False
        If m_blnDoGVerify And m_blnDoMVerify Then

            blnDoVerify = True
            If MsgBox("Both the ""Games"" and ""Music"" tables must be verified." & vbCrLf & vbCrLf & "Click YES to verify now." & vbCrLf & vbCrLf & "If you click NO, you'll need to run ""Verify Available Files..."" from GameBase's Tools menu.", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbNo Then
                blnDoVerify = False
            End If
        
        ElseIf m_blnDoGVerify Then
        
            blnDoVerify = True
            If MsgBox("The ""Games"" table must be verified." & vbCrLf & vbCrLf & "Click YES to verify now." & vbCrLf & vbCrLf & "If you click NO, you'll need to run ""Verify Available Files..."" from GameBase's Tools menu.", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbNo Then
                blnDoVerify = False
            End If
        
        ElseIf m_blnDoMVerify Then
        
            blnDoVerify = True
            If MsgBox("The ""Music"" table must be verified." & vbCrLf & vbCrLf & "Click YES to verify now." & vbCrLf & vbCrLf & "If you click NO, you'll need to run ""Verify Available Files..."" from GameBase's Tools menu.", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbNo Then
                blnDoVerify = False
            End If

        Else
            'no verify
            blnVN = True
            lblVerify.Visible = True
            lblVerify = "Finished. (Verify was not needed)"

        End If
        If blnVN = False Then
            If blnDoVerify Then
                cmdAbort.Enabled = False
                cmdNext.Enabled = False
                lblStatus.Visible = True
                pbStatus.Visible = True
                Screen.MousePointer = vbHourglass
                picWiz(5).ZOrder 0
                lblVerify = "Verifying..."
                lblVerify.Visible = True
                DoEvents
                m_blnQuittable = False
                strError = VerifyAvailableFiles()
                m_blnQuittable = True
                If strError <> "" Then
                    MsgBox strError, vbCritical, gc_strAPPNAME
                End If
            Else
                lblVerify.Visible = True
                lblVerify = "Finished. (Verify was skipped.  You will need to do this before using the database in GameBase.)"
            End If
        End If
        Screen.MousePointer = vbNormal
        cmdAbort.Enabled = False
        cmdAbort.Visible = False
        pbStatus.Visible = False
        lblStatus.Visible = False
        cmdNext.Enabled = True
        cmdNext.Caption = "&Close"
        m_lngState = 2
        picWiz(5).ZOrder 0
        If cmdVerify.Enabled = True And cmdVerify.Visible = True Then
            cmdVerify.SetFocus
        End If
        m_blnDoGVerify = False
        m_blnDoMVerify = False

    Else
        'close
        Unload Me

    End If
    

End Sub

Private Sub cmdVerify_Click()

    If g_xobjFunc.FileExists(g_strAppPathWithBackSlash & "verify.txt", False, "") Then
        Call g_xobjFunc.ShellExecuteFile(g_strAppPathWithBackSlash & "verify.txt", "", False)
    Else
        MsgBox "The file was not found!" & vbCrLf & vbCrLf & g_strAppPathWithBackSlash & "verify.txt", vbExclamation, gc_strAPPNAME
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 5565    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 9120     'The required scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim lngMax As Long
    Dim strDBPath As String
    Dim strSPPath As String
    Dim strHVPath As String
    Dim strPHPath As String
    Dim lngAction As Long

    m_lngState = 0
    m_blnDoGVerify = False
    m_blnDoMVerify = False
    m_blnQuittable = True
    Set g_xobjFunc = New GBFuncs.clsFunctions
    Set g_xobjReg = New GBFuncs.clsRegistry

    g_strAPPVERSION = Format$(Val(App.Major), "0") & "." & Format$(Val(App.Minor), "0")
    Caption = gc_strAPPNAME & "  v" & g_strAPPVERSION & "  [GameBase DB v" & gc_strREQ_DB_VERSION & "]  " & gc_strCOPYRIGHT
    
    g_strAppPathWithBackSlash = App.Path
    If Right$(g_strAppPathWithBackSlash, 1) <> "\" Then
        g_strAppPathWithBackSlash = g_strAppPathWithBackSlash & "\"
    End If

    'get current size/position/state variables
    With g_xobjReg
        'main window size/position
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainWidth", sngRegWidth, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainHeight", sngRegHeight, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainTop", sngRegTop, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainLeft", sngRegLeft, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "Maximized", lngMax, False, 0
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Database", strDBPath, True, ""
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Sidplayer", strSPPath, True, ""
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "HVSC", strHVPath, True, ""
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Paths", "Photos", strPHPath, True, ""
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Options", "Action", lngAction, False, 0
    End With

    'ListView stuff
    GetListViewSettings

    'set the paths
    txtDatabase = strDBPath
    txtSidplayer = strSPPath
    txtHVSC = strHVPath
    txtPhotos = strPHPath
    IntroStartCheck

    'make sure not to big or small...
    If sngRegWidth < DesignScaleWidth Then
        sngRegWidth = DesignScaleWidth + (Width - ScaleWidth)
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight < DesignScaleHeight Then
        sngRegHeight = DesignScaleHeight + (Height - ScaleHeight)
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    '...or out of full screen view
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (sngRegHeight / 2)
        sngRegLeft = (Screen.Width / 2) - (sngRegWidth / 2)
    End If
    'set the position and size
    Top = sngRegTop
    Left = sngRegLeft
    Width = sngRegWidth
    Height = sngRegHeight
    If lngMax = 1 Then
        WindowState = vbMaximized
    End If
    
    If lngAction < 0 Or lngAction > 4 Then
        lngAction = 0
    End If
    optAction(lngAction).Value = True
    
    picWiz(0).ZOrder 0

End Sub

Private Function BuildCachedDatabase() As String

    'builds a cached database into a text file for certain parts of this wizard

    Dim intFN As Integer
    Dim strCacheFile As String
    Dim lngIdx As Long
    Dim strError As String

    On Error GoTo Err_Handler

    Screen.MousePointer = vbHourglass
    DoEvents

    'delete the file if it's already there
    strCacheFile = g_strAppPathWithBackSlash & "cache.txt"
    If g_xobjFunc.FileExists(strCacheFile, False, "") Then
        Kill strCacheFile
    End If

    'we're not adding all the sids to one big long string as in VB this would
    'take hours for 20000+ sids, so we'll use lots of smaller strings holding
    '500 sids each, then write them in order to the cache.txt file
    g_lngTotalCacheStrings = 1
    g_lngCurrentCacheString = 0
    g_lngTotalSids = 0
    g_lngCurrentCacheItem = 0
    For lngIdx = 0 To 999
        g_strCache(lngIdx) = ""
    Next lngIdx
    strError = DoCache()

    'was there an error?
    If strError <> "" Then
        BuildCachedDatabase = strError
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    'write the strings to the cache file
    intFN = FreeFile()
    Open strCacheFile For Output As #intFN
    Print #intFN, g_lngTotalSids
    For lngIdx = 0 To g_lngTotalCacheStrings - 1
        Print #intFN, g_strCache(lngIdx)
    Next lngIdx
    Close #intFN

    'all done
    Screen.MousePointer = vbNormal
    BuildCachedDatabase = ""
    Exit Function

Err_Handler:
    BuildCachedDatabase = "There was an error whilst building the cache of your HVSC: -" & vbCrLf & vbCrLf & Err.Description & "  [Error " & Err.Number & "," & Err.LastDllError & "]"
    'enable controls
    Screen.MousePointer = vbNormal

End Function


Private Function AtLeastOneItemSelectedInListView(lv_in As ListView, blnMultiple_out As Boolean) As Boolean

    'returns the total number of items selected in the passed in listview
    'also outputs whether multiple (two or more) items are selected (in the blnMultiple_out var)

    Dim Litem As ListItem
    Dim lngCount As Long

    AtLeastOneItemSelectedInListView = False
    lngCount = 0
    blnMultiple_out = False
    With lv_in
        For Each Litem In .ListItems
            If Litem.Selected Then
                AtLeastOneItemSelectedInListView = True
                lngCount = lngCount + 1
                If lngCount = 2 Then
                    blnMultiple_out = True
                    Exit Function
                End If
            End If
        Next Litem
    End With

End Function

Private Function FillInvalidLinkedSidsList() As Long

    'Fills the "Invalid Sid Links" listbox

    Dim rsGames As DAO.Recordset
    Dim fldGAID As DAO.Field
    Dim fldName As DAO.Field
    Dim fldMusician As DAO.Field
    Dim fldMUID As DAO.Field
    Dim fldGroup As DAO.Field
    Dim fldNick As DAO.Field
    Dim fldPhoto As DAO.Field
    Dim fldSid As DAO.Field
    Dim Litem As ListItem
    Dim lngSortKey As Long
    Dim lngSortOrder As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long

    'games table
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Games.GA_Id, Games.Name, Games.SidFilename, Musicians.MU_Id, Musicians.Musician, Musicians.Grp, Musicians.Nick, Musicians.Photo  FROM Games, Musicians WHERE (Games.MU_Id=Musicians.MU_Id) AND (NOT Games.SidFilename = """")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Reading in invalid linked sids (Games table)..."
            pbStatus.Value = 0
            DoEvents
            lvInvalidLinks.Visible = False
            lvInvalidLinks.Enabled = False
            Set fldGAID = !GA_Id
            Set fldName = !Name
            Set fldMusician = !Musician
            Set fldMUID = !MU_Id
            Set fldGroup = !Grp
            Set fldNick = !Nick
            Set fldPhoto = !Photo
            Set fldSid = !SidFileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            Do Until .EOF
                If Not g_xobjFunc.FileExists(g_strHVSCPathWBS & fldSid, False, "") Then
                    '| Type (game/sid) | Name | Musician in DB | DB Group | DB Nick | DB Photo | incorrect filename | New found filename (if found) | Author in Sid file |
                    Set Litem = lvInvalidLinks.ListItems.Add(, "G" & fldGAID, " Game", 2, 2)
                    With Litem
                        .SubItems(1) = fldName
                        .SubItems(2) = fldMusician
                        .SubItems(3) = fldGroup
                        .SubItems(4) = fldNick
                        .SubItems(5) = fldPhoto
                        .SubItems(6) = fldSid
                        .SubItems(7) = ""
                        .SubItems(8) = ""
                        .Tag = fldMUID
                    End With
                    Set Litem = Nothing
                End If
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldGAID = Nothing
            Set fldName = Nothing
            Set fldMusician = Nothing
            Set fldMUID = Nothing
            Set fldGroup = Nothing
            Set fldNick = Nothing
            Set fldPhoto = Nothing
            Set fldSid = Nothing
            lvInvalidLinks.Visible = True
            lvInvalidLinks.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        .Close
    End With
    Set rsGames = Nothing

    'music table
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Music.GA_Id, Music.Name, Music.Filename, Musicians.MU_Id, Musicians.Musician, Musicians.Grp, Musicians.Nick, Musicians.Photo  FROM Music, Musicians WHERE (Music.MU_Id=Musicians.MU_Id) AND (NOT Music.Filename = """")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Reading in invalid linked sids (Music table)..."
            pbStatus.Value = 0
            DoEvents
            lvInvalidLinks.Visible = False
            lvInvalidLinks.Enabled = False
            Set fldGAID = !GA_Id
            Set fldName = !Name
            Set fldMusician = !Musician
            Set fldMUID = !MU_Id
            Set fldGroup = !Grp
            Set fldNick = !Nick
            Set fldPhoto = !Photo
            Set fldSid = !FileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            Do Until .EOF
                If Not g_xobjFunc.FileExists(g_strHVSCPathWBS & fldSid, False, "") Then
                    '| Type (game/sid) | Name | Musician in DB | DB Group | DB Nick | DB Photo | incorrect filename | New found filename (if found) | Author in Sid file |
                    Set Litem = lvInvalidLinks.ListItems.Add(, "M" & fldGAID, " Sid", 1, 1)
                    With Litem
                        .SubItems(1) = fldName
                        .SubItems(2) = fldMusician
                        .SubItems(3) = fldGroup
                        .SubItems(4) = fldNick
                        .SubItems(5) = fldPhoto
                        .SubItems(6) = fldSid
                        .SubItems(7) = ""
                        .SubItems(8) = ""
                        .Tag = fldMUID
                    End With
                    Set Litem = Nothing
                End If
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldGAID = Nothing
            Set fldName = Nothing
            Set fldMusician = Nothing
            Set fldMUID = Nothing
            Set fldGroup = Nothing
            Set fldNick = Nothing
            Set fldPhoto = Nothing
            Set fldSid = Nothing
            lvInvalidLinks.Visible = True
            lvInvalidLinks.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        .Close
    End With
    Set rsGames = Nothing

    FillInvalidLinkedSidsList = lvInvalidLinks.ListItems.Count

End Function
Private Function FillCompareList() As Long

    'Fills the "Compare" listbox

    Dim rsGames As DAO.Recordset
    Dim fldGAID As DAO.Field
    Dim fldName As DAO.Field
    Dim fldMusician As DAO.Field
    Dim fldMUID As DAO.Field
    Dim fldGroup As DAO.Field
    Dim fldNick As DAO.Field
    Dim fldSid As DAO.Field
    Dim fldPub As DAO.Field
    Dim fldYear As DAO.Field
    Dim Litem As ListItem
    Dim lngSortKey As Long
    Dim lngSortOrder As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long
    Dim lngMax As Long
    Dim blnFound As Boolean
    Dim strUpper As String
    Dim strUpper2 As String
    
    Dim fbott As Long
    Dim ftmp As Long
    Dim ftop As Long
    Dim hit As Boolean

    lngMax = g_lngCacheTotal - 1

    'games table
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Games.GA_Id, Games.Name, Games.SidFilename, Musicians.MU_Id, Musicians.Musician, Musicians.Grp, Musicians.Nick, Musicians.Photo, Years.YE_Id, Publishers.Publisher FROM Games, Musicians, Years, Publishers WHERE (Games.MU_Id=Musicians.MU_Id) AND (Games.PU_Id=Publishers.PU_Id) AND (Games.YE_Id=Years.YE_Id) AND (NOT Games.SidFilename = """")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Reading in Sid details (Games table)..."
            pbStatus.Value = 0
            DoEvents
            lvCompare.Visible = False
            lvCompare.Enabled = False
            Set fldGAID = !GA_Id
            Set fldName = !Name
            Set fldMusician = !Musician
            Set fldMUID = !MU_Id
            Set fldGroup = !Grp
            Set fldNick = !Nick
            Set fldPub = !Publisher
            Set fldYear = !YE_Id
            Set fldSid = !SidFileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            Do Until .EOF
                '| Type (game/sid) | GB64 Name | Sid Name | GB64 Musician | Sid Author | GB64 Year/Pub | Sid Copyright | File |
                Set Litem = lvCompare.ListItems.Add(, "G" & fldGAID, " Game", 2, 2)
                With Litem
                    .SubItems(1) = fldName
                    .SubItems(3) = fldMusician
                    .SubItems(5) = ga_strYearLookup(fldYear) & " " & fldPub
                    .SubItems(7) = fldSid
                    .Tag = fldMUID
                End With
                'find the sid details
                
                blnFound = False
                hit = False
                ftop = lngMax
                fbott = 0
                strUpper = UCase$(fldSid)
'                If strUpper = "DAGLISH_BEN\KRAKOUT.SID" Then
'                    Stop
'                End If
                Do While (fbott <= ftop)
                    ftmp = Int((fbott + ftop) / 2)
                    lngIdx = ftmp
                    With ga_udtCache(lngIdx)
                        strUpper2 = UCase$(.strPathFile)
                        If strUpper2 = strUpper Then
                            blnFound = True
                            Litem.SubItems(2) = .strName
                            Litem.SubItems(4) = .strAuthor
                            Litem.SubItems(6) = .strCopyright
                            Exit Do
                        ElseIf strUpper < strUpper2 Then
                            ftop = ftmp - 1
                        Else
                            fbott = ftmp + 1
                        End If
                    End With
                Loop
                If Not blnFound Then
                    With Litem
                        .SubItems(2) = ""
                        .SubItems(4) = ""
                        .SubItems(6) = ""
                        .SubItems(7) = .SubItems(7) & " [Not Found]"
                        .Ghosted = True
                    End With
                End If
                Set Litem = Nothing
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldGAID = Nothing
            Set fldName = Nothing
            Set fldMusician = Nothing
            Set fldMUID = Nothing
            Set fldGroup = Nothing
            Set fldNick = Nothing
            Set fldPub = Nothing
            Set fldYear = Nothing
            Set fldSid = Nothing
            lvCompare.Visible = True
            lvCompare.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        .Close
    End With
    Set rsGames = Nothing

    'music table
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Music.GA_Id, Music.Name, Music.Filename AS SidFilename, Musicians.MU_Id, Musicians.Musician, Musicians.Grp, Musicians.Nick, Musicians.Photo FROM Music, Musicians WHERE (Music.MU_Id=Musicians.MU_Id) AND (NOT Music.Filename = """")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Reading in Sid details (Music table)..."
            pbStatus.Value = 0
            DoEvents
            lvCompare.Visible = False
            lvCompare.Enabled = False
            Set fldGAID = !GA_Id
            Set fldName = !Name
            Set fldMusician = !Musician
            Set fldMUID = !MU_Id
            Set fldGroup = !Grp
            Set fldNick = !Nick
            Set fldSid = !SidFileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            Do Until .EOF
                '| Type (game/sid) | GB64 Name | Sid Name | GB64 Musician | Sid Author | GB64 Year/Pub | Sid Copyright | File |
                Set Litem = lvCompare.ListItems.Add(, "M" & fldGAID, " Sid", 1, 1)
                With Litem
                    .SubItems(1) = fldName
                    .SubItems(3) = fldMusician
                    .SubItems(5) = ""
                    .SubItems(7) = fldSid
                    .Tag = fldMUID
                End With
                'find the sid details
                
                blnFound = False
                hit = False
                ftop = lngMax
                fbott = 0
                ftmp = Int((ftop + fbott) / 2)
                strUpper = UCase$(fldSid)
                Do While (Not hit) And (ftop >= fbott)
                    lngIdx = ftmp
                    With ga_udtCache(lngIdx)
                        strUpper2 = UCase$(.strPathFile)
                        If strUpper2 = strUpper Then
                            hit = True
                        ElseIf strUpper < strUpper2 Then
                            ftop = ftmp - 1
                        Else
                            fbott = ftmp + 1
                        End If
                        ftmp = Int((ftop + fbott) / 2)
                        If hit Then
                            blnFound = True
                            Litem.SubItems(2) = .strName
                            Litem.SubItems(4) = .strAuthor
                            Litem.SubItems(6) = .strCopyright
                            Exit Do
                        End If
                    End With
                Loop
                If Not blnFound Then
                    With Litem
                        .SubItems(2) = ""
                        .SubItems(4) = ""
                        .SubItems(6) = ""
                        .SubItems(7) = .SubItems(7) & " [Not Found]"
                        .Ghosted = True
                    End With
                End If
                Set Litem = Nothing
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldGAID = Nothing
            Set fldName = Nothing
            Set fldMusician = Nothing
            Set fldMUID = Nothing
            Set fldGroup = Nothing
            Set fldNick = Nothing
            Set fldSid = Nothing
            lvCompare.Visible = True
            lvCompare.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        .Close
    End With
    Set rsGames = Nothing
    FillCompareList = lvCompare.ListItems.Count

End Function
Private Sub QuickSortCacheArray(inLow As Long, inHi As Long)
  
   Dim pivot As CacheArrayType
   Dim tmpSwap As CacheArrayType
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = ga_udtCache((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (UCase$(ga_udtCache(tmpLow).strPathFile) < UCase$(pivot.strPathFile) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (UCase$(pivot.strPathFile) < UCase$(ga_udtCache(tmpHi).strPathFile) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = ga_udtCache(tmpLow)
         ga_udtCache(tmpLow) = ga_udtCache(tmpHi)
         ga_udtCache(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortCacheArray inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortCacheArray tmpLow, inHi
  
End Sub
Private Sub QuickSortMusicTableArray(inLow As Long, inHi As Long, a_Music() As MusicTableType)
  
   Dim pivot As MusicTableType
   Dim tmpSwap As MusicTableType
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = a_Music((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (UCase$(a_Music(tmpLow).strFilename) < UCase$(pivot.strFilename) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (UCase$(pivot.strFilename) < UCase$(a_Music(tmpHi).strFilename) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = a_Music(tmpLow)
         a_Music(tmpLow) = a_Music(tmpHi)
         a_Music(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortMusicTableArray inLow, tmpHi, a_Music()
   If (tmpLow < inHi) Then QuickSortMusicTableArray tmpLow, inHi, a_Music()
  
End Sub
Private Sub QuickSortStringArray(inLow As Long, inHi As Long, a_strings() As String)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = a_strings((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (UCase$(a_strings(tmpLow)) < UCase$(pivot) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (UCase$(pivot) < UCase$(a_strings(tmpHi)) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = a_strings(tmpLow)
         a_strings(tmpLow) = a_strings(tmpHi)
         a_strings(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortStringArray inLow, tmpHi, a_strings()
   If (tmpLow < inHi) Then QuickSortStringArray tmpLow, inHi, a_strings()
  
End Sub
Private Sub QuickSortCacheArrayByFile(inLow As Long, inHi As Long)
  
   Dim pivot As CacheArrayType
   Dim tmpSwap As CacheArrayType
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = ga_udtCache((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (UCase$(ga_udtCache(tmpLow).strFile) < UCase$(pivot.strFile) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (UCase$(pivot.strFile) < UCase$(ga_udtCache(tmpHi).strFile) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = ga_udtCache(tmpLow)
         ga_udtCache(tmpLow) = ga_udtCache(tmpHi)
         ga_udtCache(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortCacheArrayByFile inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortCacheArrayByFile tmpLow, inHi
  
End Sub
Private Function FillMissingSidsList() As Long

    'Fills the "Missing Sids" listbox

    Dim rsMusic As DAO.Recordset
    Dim rsGames As DAO.Recordset
    Dim fldSid As DAO.Field
    Dim Litem As ListItem
    Dim lngSortKey As Long
    Dim lngSortOrder As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long
    Dim lngMax As Long
    Dim blnFound As Boolean
    Dim strUpper As String
    Dim strUpper2 As String

    Dim fbott As Long
    Dim ftmp As Long
    Dim ftop As Long
    Dim hit As Boolean

    'find all missing sids
    Screen.MousePointer = vbHourglass
    lblStatus = "Finding all HVSC Sids not linked to the GB64 database (this may take a while)..."
    pbStatus.Value = 0
    DoEvents
    lvMissing.Visible = False
    lvMissing.Enabled = False
    lngMax = g_lngCacheTotal - 1
    pbStatus.Min = 0
    pbStatus.Max = g_lngCacheTotal
    lngMax = g_lngCacheTotal - 1

    Set rsGames = g_dbGB64.OpenRecordset("SELECT DISTINCTROW SidFilename AS Filename FROM Games WHERE (NOT SidFilename="""") UNION SELECT DISTINCTROW Filename FROM Music WHERE (NOT Filename="""")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Finding all HVSC Sids not linked to the GB64 database..."
            pbStatus.Value = 0
            DoEvents
            lvMissing.Visible = False
            lvMissing.Enabled = False
            Set fldSid = !FileName
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            lngMax = g_lngCacheTotal - 1
            Do Until .EOF
                hit = False
                ftop = lngMax
                fbott = 0
                ftmp = Int((ftop + fbott) / 2)
                strUpper = UCase$(fldSid)
                Do While (Not hit) And (ftop >= fbott)
                    lngIdx = ftmp
                    With ga_udtCache(lngIdx)
                    strUpper2 = UCase$(.strPathFile)
                        If strUpper2 = strUpper Then
                            hit = True
                        ElseIf strUpper < strUpper2 Then
                            ftop = ftmp - 1
                        Else
                            fbott = ftmp + 1
                        End If
                        ftmp = Int((ftop + fbott) / 2)
                        If hit Then
                            .blnFlag = True
                            Exit Do
                        End If
                    End With
                Loop
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldSid = Nothing
        End If
        .Close
    End With
    Set rsGames = Nothing

    'now fill
    lblStatus = "Finding all HVSC Sids not linked to the GB64 database..."
    pbStatus.Value = 0
    DoEvents
    For lngIdx = 0 To lngMax
        With ga_udtCache(lngIdx)
            If Not .blnFlag Then
                Set Litem = lvMissing.ListItems.Add(, "S" & .lngID, .strName, 1, 1)
                Litem.SubItems(1) = .strPath
                Litem.SubItems(2) = .strFile
                Litem.SubItems(3) = .strAuthor
                Litem.SubItems(4) = .strCopyright
                Set Litem = Nothing
            End If
        End With
    Next lngIdx
    lvMissing.Visible = True
    lvMissing.Enabled = True
    pbStatus.Value = pbStatus.Max
    Screen.MousePointer = vbNormal
    
'    Set rsMusic = g_dbGB64.OpenRecordset("SELECT DISTINCTROW Filename FROM Music WHERE NOT Filename="""" UNION SELECT DISTINCTROW SidFilename AS Filename FROM Games WHERE NOT SidFilename=""""", dbOpenSnapshot)
'    For lngIdx = 0 To lngMax
'        blnFound = True
'        With rsMusic
'            .FindFirst "Filename = '" & ga_udtCache(lngIdx).strPathFile & "'"
'            If .NoMatch Then
'                'add to listview
'                With ga_udtCache(lngIdx)
'                    '| Sid Name | HVSC Path | Filename | Author in Sid File | Sid Copyright
'                    Set LItem = lvMissing.ListItems.Add(, "S" & .lngID, .strName, 1, 1)
'                    LItem.SubItems(1) = .strPath
'                    LItem.SubItems(2) = .strFile
'                    LItem.SubItems(3) = .strAuthor
'                    LItem.SubItems(4) = .strCopyright
'                    Set LItem = Nothing
'                End With
'            End If
'        End With
'        With pbStatus
'            .Value = .Value + 1
'        End With
'        DoEvents
'    Next lngIdx
'    rsMusic.Close
'    Set rsMusic = Nothing
'    pbStatus.Value = pbStatus.Max
'    lvMissing.Visible = True
'    lvMissing.Enabled = True
'    Screen.MousePointer = vbNormal

    FillMissingSidsList = lvMissing.ListItems.Count

End Function

Private Function FillGamesWithNoSidsList() As Long

    'Fills the "Games No Sids" listbox

    Dim rsGames As DAO.Recordset
    Dim fldGAID As DAO.Field
    Dim fldName As DAO.Field
    Dim fldMusician As DAO.Field
    Dim fldMUID As DAO.Field
    Dim fldGroup As DAO.Field
    Dim fldNick As DAO.Field
    Dim fldPhoto As DAO.Field
    Dim fldSid As DAO.Field
    Dim fldPub As DAO.Field
    Dim fldYear As DAO.Field
    Dim Litem As ListItem
    Dim lngSortKey As Long
    Dim lngSortOrder As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long

    'games table
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Games.GA_Id, Games.Name, Musicians.MU_Id, Musicians.Musician, Musicians.Grp, Musicians.Nick, Musicians.Photo, Years.YE_Id, Publishers.Publisher FROM Games, Musicians, Years, Publishers WHERE (Games.MU_Id=Musicians.MU_Id) AND (Games.PU_Id=Publishers.PU_Id) AND (Games.YE_Id=Years.YE_Id) AND (Games.SidFilename = """")", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            Screen.MousePointer = vbHourglass
            lblStatus = "Reading in Games with no linked sid file..."
            pbStatus.Value = 0
            DoEvents
            lvGamesNoSids.Visible = False
            lvGamesNoSids.Enabled = False
            Set fldGAID = !GA_Id
            Set fldName = !Name
            Set fldMusician = !Musician
            Set fldMUID = !MU_Id
            Set fldGroup = !Grp
            Set fldNick = !Nick
            Set fldPhoto = !Photo
            Set fldPub = !Publisher
            Set fldYear = !YE_Id
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount + 1
            Do Until .EOF
                '| Name | Musician | Group | Nick | Photo | Year | Publisher |
                Set Litem = lvGamesNoSids.ListItems.Add(, "G" & fldGAID, fldName, 3, 3)
                With Litem
                    .SubItems(1) = fldMusician
                    .SubItems(2) = fldGroup
                    .SubItems(3) = fldNick
                    .SubItems(4) = fldPhoto
                    .SubItems(5) = ga_strYearLookup(fldYear)
                    .SubItems(6) = fldPub
                    .Ghosted = (fldMUID = 39)
                    .Tag = fldMUID
                End With
                Set Litem = Nothing
                With pbStatus
                    .Value = .Value + 1
                End With
                .MoveNext
            Loop
            pbStatus.Value = pbStatus.Max
            Set fldGAID = Nothing
            Set fldName = Nothing
            Set fldMusician = Nothing
            Set fldMUID = Nothing
            Set fldGroup = Nothing
            Set fldNick = Nothing
            Set fldPhoto = Nothing
            Set fldPub = Nothing
            Set fldYear = Nothing
            lvGamesNoSids.Visible = True
            lvGamesNoSids.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        .Close
    End With
    Set rsGames = Nothing

    FillGamesWithNoSidsList = lvGamesNoSids.ListItems.Count

End Function
Private Function FillHVSCCacheList() As Long

    'Fills the "HVSC Cache" listbox

    Dim Litem As ListItem
    Dim lngSortKey As Long
    Dim lngSortOrder As Long
    Dim lngWidth As Long
    Dim lngIdx As Long
    Dim lngPos As Long
    Dim lngMax As Long

    'games table
    Screen.MousePointer = vbHourglass
    lblStatus = "Reading in the HVSC..."
    pbStatus.Value = 0
    DoEvents
    lvCache.Visible = False
    lvCache.Enabled = False
    pbStatus.Min = 0
    pbStatus.Max = g_lngCacheTotal
    lngMax = g_lngCacheTotal - 1
    For lngIdx = 0 To lngMax
        '| Sid Name | HVSC Path | FileName | Author in Sid File | Copyright in Sid file
        With ga_udtCache(lngIdx)
            Set Litem = lvCache.ListItems.Add(, "S" & .lngID, .strName, 1, 1)
        End With
        With Litem
            .SubItems(1) = ga_udtCache(lngIdx).strPath
            .SubItems(2) = ga_udtCache(lngIdx).strFile
            .SubItems(3) = ga_udtCache(lngIdx).strAuthor
            .SubItems(4) = ga_udtCache(lngIdx).strCopyright
        End With
        Set Litem = Nothing
        With pbStatus
            .Value = .Value + 1
        End With
    Next lngIdx
    pbStatus.Value = pbStatus.Max
    lvCache.Visible = True
    lvCache.Enabled = True
    Screen.MousePointer = vbNormal

    FillHVSCCacheList = lvCache.ListItems.Count

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If m_blnQuittable = False Then
        If UnloadMode = vbFormControlMenu Then
            Cancel = True
        End If
    End If

End Sub

Private Sub Form_Resize()

    Dim lngIdx As Long

    If WindowState <> vbMinimized Then
        'widths
        If ScaleWidth > 7650 Then
            cmdNext.Left = ScaleWidth - 120 - cmdNext.Width
            With pbStatus
                .Left = cmdAbort.Left + cmdAbort.Width + 120
                .Width = ScaleWidth - (.Left * 2)
                lblStatus.Width = .Width
                lblStatus.Left = .Left
                fraLine.Width = ScaleWidth - 240
                For lngIdx = 0 To 5
                    picWiz(lngIdx).Width = fraLine.Width
                Next lngIdx
            End With
        End If
        'heights
        If ScaleHeight > 4905 Then
            With cmdAbort
                .Top = ScaleHeight - .Height - 120
                cmdNext.Top = .Top
            End With
            With pbStatus
                .Top = ScaleHeight - .Height - 120
                lblStatus.Top = .Top - lblStatus.Height
                fraLine.Top = lblStatus.Top - 240
                For lngIdx = 0 To 5
                    picWiz(lngIdx).Height = fraLine.Top - picWiz(lngIdx).Top
                Next lngIdx
            End With
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim lngIdx As Long
    Dim lngAction As Long

    If m_blnDoGVerify Or m_blnDoMVerify Then
        MsgBox "The GB64 database has been changed.  You will need to run ""Verify Available Files..."" in the GameBase frontend before your Sid-File links will be properly recognised.", vbInformation, gc_strAPPNAME
    End If

    'save current size/position/state variables
    If WindowState <> vbMinimized Then
        With g_xobjReg
            'main window size/position
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainWidth", Width, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainHeight", Height, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainTop", Top, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "MainLeft", Left, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Window", "Maximized", IIf((WindowState = vbMaximized), 1, 0), False
        End With
    End If
    
    'save listview settings
    With g_xobjReg
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidSK", lvInvalidLinks.SortKey, False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidSO", lvInvalidLinks.SortOrder, False
        For lngIdx = 0 To 8
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidW" & Trim$(lngIdx), lvInvalidLinks.ColumnHeaders(lngIdx + 1).Width, False
        Next lngIdx
        For lngIdx = 0 To 8
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "InvalidP" & Trim$(lngIdx), lvInvalidLinks.ColumnHeaders(lngIdx + 1).Position, False
        Next lngIdx

        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsSK", lvGamesNoSids.SortKey, False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsSO", lvGamesNoSids.SortOrder, False
        For lngIdx = 0 To 6
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsW" & Trim$(lngIdx), lvGamesNoSids.ColumnHeaders(lngIdx + 1).Width, False
        Next lngIdx
        For lngIdx = 0 To 6
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "GamesNoSidsP" & Trim$(lngIdx), lvGamesNoSids.ColumnHeaders(lngIdx + 1).Position, False
        Next lngIdx

        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheSK", lvCache.SortKey, False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheSO", lvCache.SortOrder, False
        For lngIdx = 0 To 4
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheW" & Trim$(lngIdx), lvCache.ColumnHeaders(lngIdx + 1).Width, False
        Next lngIdx
        For lngIdx = 0 To 4
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CacheP" & Trim$(lngIdx), lvCache.ColumnHeaders(lngIdx + 1).Position, False
        Next lngIdx
    
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingSK", lvMissing.SortKey, False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingSO", lvMissing.SortOrder, False
        For lngIdx = 0 To 4
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingW" & Trim$(lngIdx), lvMissing.ColumnHeaders(lngIdx + 1).Width, False
        Next lngIdx
        For lngIdx = 0 To 4
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "MissingP" & Trim$(lngIdx), lvMissing.ColumnHeaders(lngIdx + 1).Position, False
        Next lngIdx
    
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareSK", lvCompare.SortKey, False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareSO", lvCompare.SortOrder, False
        For lngIdx = 0 To 7
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareW" & Trim$(lngIdx), lvCompare.ColumnHeaders(lngIdx + 1).Width, False
        Next lngIdx
        For lngIdx = 0 To 7
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Lists", "CompareP" & Trim$(lngIdx), lvCompare.ColumnHeaders(lngIdx + 1).Position, False
        Next lngIdx
    
    End With
    
    'save misc options
    With g_xobjReg
        If optAction(0).Value = True Then
            lngAction = 0
        ElseIf optAction(1).Value = True Then
            lngAction = 1
        ElseIf optAction(2).Value = True Then
            lngAction = 2
        ElseIf optAction(3).Value = True Then
            lngAction = 3
        Else
            lngAction = 4
        End If
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\HVSCWiz\Options", "Action", lngAction, False
    End With

    Set g_xobjFunc = Nothing
    Set g_xobjReg = Nothing
    Set g_dbGB64 = Nothing

End Sub


Private Sub lvCache_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvCache
        If (ColumnHeader.Index - 1) = .SortKey Then
            'reverse sort order
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            'set asc sort order on new column
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

    Call SetSortIconsForCache
    Screen.MousePointer = vbNormal
    
End Sub


Private Sub lvCache_DblClick()

    Dim Litem As ListItem
    Dim strFile As String

    Set Litem = lvCache.SelectedItem
    If Not Litem Is Nothing Then
        strFile = g_strHVSCPathWBS & Litem.SubItems(1) & "\" & Litem.SubItems(2)
        Call PlaySid(strFile, False)
    End If

End Sub


Private Sub lvCompare_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvCompare
        If (ColumnHeader.Index - 1) = .SortKey Then
            'reverse sort order
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            'set asc sort order on new column
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

    Call SetSortIconsForCompare
    Screen.MousePointer = vbNormal

End Sub

Private Sub lvCompare_DblClick()

    Dim Litem As ListItem
    Dim strFile As String

    Set Litem = lvCompare.SelectedItem
    If Not Litem Is Nothing Then
        strFile = g_strHVSCPathWBS & Litem.SubItems(7)
        If InStr(1, Litem.SubItems(7), "[not found]", vbTextCompare) = 0 Then
            Call PlaySid(strFile, False)
        End If
    End If

End Sub

Private Sub lvGamesNoSids_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvGamesNoSids
        If (ColumnHeader.Index - 1) = .SortKey Then
            'reverse sort order
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            'set asc sort order on new column
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

    Call SetSortIconsForGamesNoSids
    Screen.MousePointer = vbNormal

End Sub


Private Sub lvInvalidLinks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvInvalidLinks
        If (ColumnHeader.Index - 1) = .SortKey Then
            'reverse sort order
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            'set asc sort order on new column
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

    Call SetSortIconsForInvalid
    Screen.MousePointer = vbNormal

End Sub


Private Sub lvInvalidLinks_DblClick()

    Dim Litem As ListItem
    Dim strFile As String

    Set Litem = lvInvalidLinks.SelectedItem
    If Not Litem Is Nothing Then
        strFile = g_strHVSCPathWBS & Litem.SubItems(7)
        If Trim$(Litem.SubItems(7)) <> "" Then
            Call PlaySid(strFile, False)
        End If
    End If

End Sub

Private Sub lvMissing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvMissing
        If (ColumnHeader.Index - 1) = .SortKey Then
            'reverse sort order
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            'set asc sort order on new column
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

    Call SetSortIconsForMissing
    Screen.MousePointer = vbNormal

End Sub


Private Sub lvMissing_DblClick()

    Dim Litem As ListItem
    Dim strFile As String

    Set Litem = lvMissing.SelectedItem
    If Not Litem Is Nothing Then
        strFile = g_strHVSCPathWBS & Litem.SubItems(1) & "\" & Litem.SubItems(2)
        Call PlaySid(strFile, False)
    End If

End Sub


Private Sub m_xobjFinder_FoundFile(ByVal strPathAndFile As String)

'Public g_strCache(999) As String
'Public g_lngTotalCacheStrings As Long
'Public g_lngCurrentCacheString As Long
'Public g_lngTotalSids As Long

    'this is only for building the cache
    Dim strFile As String
    Dim strPath As String
    
    Dim strSIDName As String
    Dim strSIDAuthor As String
    Dim strSIDCopyright As String
    Dim lngSIDTotalSongs As Long
    Dim lngSIDDefaultSong As Long
    

    On Error GoTo Err_Handler

    'which file are we using?
    If g_lngCurrentCacheItem = 499 Then '500 sids per string
        g_lngCurrentCacheItem = 0
        g_lngCurrentCacheString = g_lngCurrentCacheString + 1
        g_lngTotalCacheStrings = g_lngTotalCacheStrings + 1
    End If

    'get the file and path
    Call g_xobjFunc.GetFileAndPath(strPathAndFile, strPath, strFile)
    strPath = Right$(strPath, Len(strPath) - g_lngLenHVSCPathWBS)
    If m_strLastPath <> strPath Then
        m_strLastPath = "Adding: " & strPath
        lblStatus = m_strLastPath
        DoEvents
    End If
    
    'add it to the cache string
    g_strCache(g_lngCurrentCacheString) = g_strCache(g_lngCurrentCacheString) & strPath & "\" & strFile & gc_strDELIM & strFile & gc_strDELIM & strPath & gc_strDELIM
    
    'get the sid info
    If g_xobjFunc.GetSidInfo(strPathAndFile, strSIDName, strSIDAuthor, strSIDCopyright, lngSIDTotalSongs, lngSIDDefaultSong) Then
        'woohoo
        g_strCache(g_lngCurrentCacheString) = g_strCache(g_lngCurrentCacheString) & strSIDName & gc_strDELIM & strSIDAuthor & gc_strDELIM & strSIDCopyright & gc_strDELIM & lngSIDTotalSongs & gc_strDELIM & lngSIDDefaultSong & vbCrLf
    Else
        'dang!
        g_strCache(g_lngCurrentCacheString) = g_strCache(g_lngCurrentCacheString) & "[Could not read]" & gc_strDELIM & "[Could not read]" & gc_strDELIM & "[Could not read]" & gc_strDELIM & "[Could not read]" & gc_strDELIM & "[Could not read]" & vbCrLf
    End If
    
    'And another sid bites the dust!
    g_lngTotalSids = g_lngTotalSids + 1
    g_lngCurrentCacheItem = g_lngCurrentCacheItem + 1

    Exit Sub

Err_Handler:
    'ignore the sid!
    g_blnCacheError = True

End Sub

Private Sub picWiz_Resize(Index As Integer)

    Dim lngIdx As Long

    If Index = 0 Then
        'intro
        
        
    ElseIf Index = 1 Then
        'invalid sid links
        lvInvalidLinks.Width = picWiz(Index).Width
        For lngIdx = 0 To 4
            cmdInvalid(lngIdx).Top = picWiz(Index).Height - cmdInvalid(lngIdx).Height
        Next lngIdx
        lvInvalidLinks.Height = cmdInvalid(0).Top - 120 - lvInvalidLinks.Top

    ElseIf Index = 2 Then
        'matchup game files
        cmdGamesNoSids(0).Top = ((picWiz(Index).Height) / 2) - (cmdGamesNoSids(0).Height / 2) + lblPrompt(7).Height
        cmdGamesNoSids(1).Top = cmdGamesNoSids(0).Top
        cmdGamesNoSids(2).Top = cmdGamesNoSids(0).Top
        cmdGamesNoSids(3).Top = cmdGamesNoSids(0).Top
        lvGamesNoSids.Height = cmdGamesNoSids(0).Top - 120 - lvGamesNoSids.Top
        lvCache.Top = cmdGamesNoSids(0).Top + cmdGamesNoSids(0).Height + 120
        lvCache.Height = picWiz(Index).Height - lvCache.Top
        lblPrompt(8).Top = lvCache.Top - 240
        lvGamesNoSids.Width = picWiz(Index).Width
        lvCache.Width = lvGamesNoSids.Width
        cmdGamesNoSids(3).Left = picWiz(Index).Width - cmdGamesNoSids(3).Width
        cmdGamesNoSids(2).Left = cmdGamesNoSids(3).Left - 120 - cmdGamesNoSids(2).Width
        cmdGamesNoSids(1).Left = cmdGamesNoSids(2).Left - 120 - cmdGamesNoSids(1).Width
        cmdGamesNoSids(0).Left = cmdGamesNoSids(1).Left - 120 - cmdGamesNoSids(0).Width
        
    ElseIf Index = 3 Then
        'add non-linked sids
        lvMissing.Width = picWiz(Index).Width
        cmdMissing(0).Top = picWiz(Index).Height - cmdMissing(0).Height
        cmdMissing(1).Top = cmdMissing(0).Top
        cmdMissing(2).Top = cmdMissing(0).Top
        lvMissing.Height = cmdMissing(0).Top - 120 - lvMissing.Top
        
    ElseIf Index = 4 Then
        'compare musicians/sid-authors
        lvCompare.Width = picWiz(Index).Width
        cmdCompare(0).Top = picWiz(Index).Height - cmdCompare(0).Height
        cmdCompare(1).Top = cmdCompare(0).Top
        cmdCompare(2).Top = cmdCompare(0).Top
        lvCompare.Height = cmdCompare(0).Top - 120 - lvCompare.Top
         
    ElseIf Index = 5 Then
        'finished
        

    End If

End Sub


Private Sub txtDatabase_Change()

    IntroStartCheck

End Sub


Private Sub txtDatabase_GotFocus()

    With txtDatabase
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtHVSC_Change()

    IntroStartCheck

End Sub


Private Sub txtHVSC_GotFocus()

    With txtHVSC
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtPhotos_Change()

    IntroStartCheck

End Sub


Private Sub txtPhotos_GotFocus()

    With txtPhotos
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtSidplayer_Change()

    IntroStartCheck
    
End Sub


Private Sub txtSidplayer_GotFocus()

    With txtSidplayer
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

