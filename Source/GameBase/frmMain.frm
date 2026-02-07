VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "gamebase64 - ""The Gamebase Collection"""
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   10290
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10290
   Visible         =   0   'False
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   3000
      MouseIcon       =   "frmMain.frx":08CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   2295
      ScaleMode       =   0  'User
      ScaleWidth      =   45
      TabIndex        =   139
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picSidebar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   600
      ScaleHeight     =   2895
      ScaleWidth      =   1575
      TabIndex        =   138
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      Begin VB.PictureBox picSBSplitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   3
         Left            =   0
         MouseIcon       =   "frmMain.frx":0A1C
         MousePointer    =   99  'Custom
         ScaleHeight     =   45
         ScaleWidth      =   1125
         TabIndex        =   149
         Top             =   2640
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox picSBSplitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   2
         Left            =   0
         MouseIcon       =   "frmMain.frx":0B6E
         MousePointer    =   99  'Custom
         ScaleHeight     =   45
         ScaleWidth      =   1125
         TabIndex        =   148
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox picSBSplitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   0
         MouseIcon       =   "frmMain.frx":0CC0
         MousePointer    =   99  'Custom
         ScaleHeight     =   45
         ScaleWidth      =   1125
         TabIndex        =   147
         Top             =   1440
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox picSBSplitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmMain.frx":0E12
         MousePointer    =   99  'Custom
         ScaleHeight     =   45
         ScaleWidth      =   1125
         TabIndex        =   146
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtSidebar 
         Height          =   285
         Index           =   3
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   145
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSidebar 
         Height          =   285
         Index           =   2
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   144
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSidebar 
         Height          =   285
         Index           =   1
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   143
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSidebar 
         Height          =   285
         Index           =   0
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   142
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ImageList imgListQL 
         Left            =   1320
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0F64
               Key             =   "play_game"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1017
               Key             =   "play_music"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10C8
               Key             =   "gemus_extra"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":117B
               Key             =   "music_extra"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":122C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1395
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1411
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1698
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":174C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":17DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":188D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":193D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":19DB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvLinks 
         Height          =   375
         Left            =   0
         TabIndex        =   141
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListQL"
         SmallIcons      =   "imgListQL"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Quick Links"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picListViews 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4215
      TabIndex        =   132
      Top             =   0
      Width           =   4215
      Begin MSComctlLib.ListView lvMain 
         Height          =   3015
         Index           =   2
         Left            =   0
         TabIndex        =   133
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilstMain16"
         SmallIcons      =   "ilstMain16"
         ColHdrIcons     =   "ilstCHeaders"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "The Gamebase Collection"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   3015
         Index           =   1
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilstMain16"
         SmallIcons      =   "ilstMain16"
         ColHdrIcons     =   "ilstCHeaders"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "The Gamebase Collection"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   3015
         Index           =   0
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilstMain16"
         SmallIcons      =   "ilstMain16"
         ColHdrIcons     =   "ilstCHeaders"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "The Gamebase Collection"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   3015
         Index           =   3
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilstMain16"
         SmallIcons      =   "ilstMain16"
         ColHdrIcons     =   "ilstCHeaders"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "The Gamebase Collection"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   3015
         Index           =   4
         Left            =   0
         TabIndex        =   137
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilstMain16"
         SmallIcons      =   "ilstMain16"
         ColHdrIcons     =   "ilstCHeaders"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "The Gamebase Collection"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.PictureBox picNoSSSmall 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   6480
      Picture         =   "frmMain.frx":1A79
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   60
      Top             =   11640
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox picNoSS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      Picture         =   "frmMain.frx":1C29
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   59
      Top             =   11280
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox picFull 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6840
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   54
      Top             =   11160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPleaseWait 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label lblPleaseWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ilstCHeaders 
      Left            =   1200
      Top             =   9720
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
            Picture         =   "frmMain.frx":1F7F
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D1
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2823
            Key             =   "blank"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstMain16 
      Left            =   600
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.Frame fraPanel 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3060
      Width           =   9135
      Begin VB.Frame fraButtons 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4290
         TabIndex        =   8
         Top             =   0
         Width           =   4815
         Begin VB.CommandButton cmdSid 
            Caption         =   "Play &Sid"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1820
            TabIndex        =   4
            Top             =   0
            Width           =   1450
         End
         Begin VB.CommandButton cmdRandom 
            Caption         =   "?"
            Height          =   315
            Left            =   0
            TabIndex        =   2
            ToolTipText     =   "Random Game select"
            Top             =   0
            Width           =   260
         End
         Begin VB.CommandButton cmdExtra 
            Caption         =   "&Extras"
            Height          =   315
            Left            =   3310
            TabIndex        =   5
            Top             =   0
            Width           =   1480
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "&Play Game"
            Enabled         =   0   'False
            Height          =   315
            Left            =   300
            TabIndex        =   3
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.ComboBox lstView 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   3990
      End
      Begin VB.CommandButton cmdCustViews 
         Caption         =   "..."
         Height          =   315
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Custom Views"
         Top             =   0
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6090
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   423
            MinWidth        =   423
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   423
            MinWidth        =   423
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMain.frx":2C75
   End
   Begin VB.PictureBox picSS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   4320
      ScaleHeight     =   3015
      ScaleWidth      =   4695
      TabIndex        =   53
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox picSSDown 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   57
         Top             =   480
         Visible         =   0   'False
         Width           =   255
         Begin VB.Image imgSSDown 
            Height          =   150
            Left            =   0
            MouseIcon       =   "frmMain.frx":2F8F
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":3299
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox picSSUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   255
         Begin VB.Image imgSSUp 
            Height          =   150
            Left            =   0
            MouseIcon       =   "frmMain.frx":3601
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":390B
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox picSSMain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3000
         Index           =   0
         Left            =   0
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   316
         TabIndex        =   55
         Top             =   0
         Width           =   4800
      End
   End
   Begin TabDlg.SSTab tabDetails 
      Height          =   2250
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   3969
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      Enabled         =   0   'False
      TabCaption(0)   =   " Game"
      TabPicture(0)   =   "frmMain.frx":3C74
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Info_Comment"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lineCommentTop"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lineCommentBot"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cDialog"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "picTabHack(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   " Version"
      TabPicture(1)   =   "frmMain.frx":3D24
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LineVCommentBot"
      Tab(1).Control(1)=   "LineVCommentTop"
      Tab(1).Control(2)=   "lbl_Info_VComment"
      Tab(1).Control(3)=   "picTabHack(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Personal"
      TabPicture(2)   =   "frmMain.frx":3DCA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picTabHack(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   " Memo"
      TabPicture(3)   =   "frmMain.frx":3E85
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtMemo"
      Tab(3).Control(1)=   "cmdMemoClear"
      Tab(3).Control(2)=   "cmdMemoSave"
      Tab(3).Control(3)=   "cmdMemoExpand"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Sid Info"
      TabPicture(4)   =   "frmMain.frx":3F47
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picHack"
      Tab(4).Control(1)=   "fraSidInfo"
      Tab(4).Control(2)=   "lblSidNotFound2"
      Tab(4).ControlCount=   3
      Begin VB.PictureBox picHack 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   -74970
         ScaleHeight     =   1455
         ScaleWidth      =   5100
         TabIndex        =   123
         Top             =   645
         Width           =   5100
         Begin VB.Label lbl_SFInfo_Photo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            MouseIcon       =   "frmMain.frx":3FEB
            MousePointer    =   99  'Custom
            TabIndex        =   131
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Photo:"
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   130
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Group:"
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   129
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Nick:"
            Height          =   195
            Index           =   30
            Left            =   240
            TabIndex        =   128
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Musician:"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   127
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl_SFInfo_Musician 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1440
            MouseIcon       =   "frmMain.frx":42F5
            MousePointer    =   99  'Custom
            TabIndex        =   126
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SFInfo_Nick 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1440
            MouseIcon       =   "frmMain.frx":45FF
            MousePointer    =   99  'Custom
            TabIndex        =   125
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SFInfo_Group 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1440
            MouseIcon       =   "frmMain.frx":4909
            MousePointer    =   99  'Custom
            TabIndex        =   124
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
      End
      Begin VB.PictureBox picTabHack 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   2
         Left            =   -74970
         ScaleHeight     =   1455
         ScaleWidth      =   10215
         TabIndex        =   112
         Top             =   480
         Width           =   10215
         Begin VB.Label lbl_Info_LP 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7560
            MouseIcon       =   "frmMain.frx":4C13
            MousePointer    =   99  'Custom
            TabIndex        =   122
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_TP 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7560
            MouseIcon       =   "frmMain.frx":4F1D
            MousePointer    =   99  'Custom
            TabIndex        =   121
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Rating 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   5160
            MouseIcon       =   "frmMain.frx":5227
            MousePointer    =   99  'Custom
            TabIndex        =   120
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Difficulty 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2760
            MouseIcon       =   "frmMain.frx":5531
            MousePointer    =   99  'Custom
            TabIndex        =   119
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_HighScore 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2760
            MouseIcon       =   "frmMain.frx":583B
            MousePointer    =   99  'Custom
            TabIndex        =   118
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rating:"
            Height          =   195
            Index           =   28
            Left            =   4215
            TabIndex        =   117
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   885
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Played:"
            Height          =   195
            Index           =   27
            Left            =   6360
            TabIndex        =   116
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Times Played:"
            Height          =   195
            Index           =   26
            Left            =   6360
            TabIndex        =   115
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Difficulty Level:"
            Height          =   195
            Index           =   25
            Left            =   1560
            TabIndex        =   114
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "High Score:"
            Height          =   195
            Index           =   24
            Left            =   1560
            TabIndex        =   113
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1185
         End
      End
      Begin VB.PictureBox picTabHack 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   -74970
         ScaleHeight     =   1455
         ScaleWidth      =   10215
         TabIndex        =   87
         Top             =   405
         Width           =   10215
         Begin VB.Label lbl_Info_Playable 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":5B45
            MousePointer    =   99  'Custom
            TabIndex        =   111
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_TitleScreen 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":5E4F
            MousePointer    =   99  'Custom
            TabIndex        =   110
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Original 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":6159
            MousePointer    =   99  'Custom
            TabIndex        =   109
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_VWebLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":6463
            MousePointer    =   99  'Custom
            TabIndex        =   108
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Playable:"
            Height          =   195
            Index           =   14
            Left            =   960
            TabIndex        =   107
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title Screen:"
            Height          =   195
            Index           =   20
            Left            =   6960
            TabIndex        =   106
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Original:"
            Height          =   195
            Index           =   18
            Left            =   6960
            TabIndex        =   105
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download:"
            Height          =   195
            Index           =   17
            Left            =   960
            TabIndex        =   104
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lbl_Info_Docs 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":676D
            MousePointer    =   99  'Custom
            TabIndex        =   103
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_LoadingScreen 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":6A77
            MousePointer    =   99  'Custom
            TabIndex        =   102
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_NTSC 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":6D81
            MousePointer    =   99  'Custom
            TabIndex        =   101
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_TDE 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":708B
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_GameLength 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":7395
            MousePointer    =   99  'Custom
            TabIndex        =   99
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_HSS 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   8355
            MouseIcon       =   "frmMain.frx":769F
            MousePointer    =   99  'Custom
            TabIndex        =   98
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Trainers 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":79A9
            MousePointer    =   99  'Custom
            TabIndex        =   97
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Cracker 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "frmMain.frx":7CB3
            MousePointer    =   99  'Custom
            TabIndex        =   96
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loading Screen:"
            Height          =   195
            Index           =   19
            Left            =   6960
            TabIndex        =   95
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Included Docs:"
            Height          =   195
            Index           =   21
            Left            =   6960
            TabIndex        =   94
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fixed for NTSC:"
            Height          =   195
            Index           =   16
            Left            =   960
            TabIndex        =   93
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "High Score Saver:"
            Height          =   195
            Index           =   22
            Left            =   6960
            TabIndex        =   92
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "True Drive Emul.:"
            Height          =   195
            Index           =   23
            Left            =   6960
            TabIndex        =   91
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Game Length:"
            Height          =   195
            Index           =   15
            Left            =   960
            TabIndex        =   90
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Trainers:"
            Height          =   195
            Index           =   13
            Left            =   960
            TabIndex        =   89
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cracking Group:"
            Height          =   195
            Index           =   12
            Left            =   960
            TabIndex        =   88
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1380
         End
      End
      Begin VB.PictureBox picTabHack 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   0
         Left            =   30
         ScaleHeight     =   1455
         ScaleWidth      =   10215
         TabIndex        =   61
         Top             =   405
         Width           =   10215
         Begin VB.Label lbl_Info_WebLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":7FBD
            MousePointer    =   99  'Custom
            TabIndex        =   86
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Rarity 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":82C7
            MousePointer    =   99  'Custom
            TabIndex        =   85
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Control 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":85D1
            MousePointer    =   99  'Custom
            TabIndex        =   84
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Artist 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":88DB
            MousePointer    =   99  'Custom
            TabIndex        =   83
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Developer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":8BE5
            MousePointer    =   99  'Custom
            TabIndex        =   82
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_License 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":8EEF
            MousePointer    =   99  'Custom
            TabIndex        =   81
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "License:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   80
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Developer:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   79
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Artist:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   78
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Control:"
            Height          =   195
            Index           =   9
            Left            =   5760
            TabIndex        =   77
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Rarity:"
            Height          =   195
            Index           =   10
            Left            =   5760
            TabIndex        =   76
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "More Info:"
            Height          =   195
            Index           =   11
            Left            =   5760
            TabIndex        =   75
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lbl_Info_Language 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":91F9
            MousePointer    =   99  'Custom
            TabIndex        =   74
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Language:"
            Height          =   195
            Index           =   8
            Left            =   5760
            TabIndex        =   73
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lbl_Info_Year 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":9503
            MousePointer    =   99  'Custom
            TabIndex        =   72
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Players 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":980D
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Genre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7155
            MouseIcon       =   "frmMain.frx":9B17
            MousePointer    =   99  'Custom
            TabIndex        =   70
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Musician 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":9E21
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Programmer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1635
            MouseIcon       =   "frmMain.frx":A12B
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_Info_Publisher 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1680
            MouseIcon       =   "frmMain.frx":A435
            MousePointer    =   99  'Custom
            TabIndex        =   67
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Players:"
            Height          =   195
            Index           =   7
            Left            =   5760
            TabIndex        =   66
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Game Genre:"
            Height          =   195
            Index           =   6
            Left            =   5760
            TabIndex        =   65
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Musician:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   64
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Programmer:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   63
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   62
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdMemoExpand 
         Caption         =   "Expand"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -66000
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdMemoSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -66000
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdMemoClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -66000
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtMemo 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74640
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   480
         Width           =   8415
      End
      Begin MSComDlg.CommonDialog cDialog 
         Left            =   1320
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraSidInfo 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   -69840
         TabIndex        =   36
         Top             =   645
         Width           =   4575
         Begin VB.Label lbl_SFInfo_Default 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4440
            MouseIcon       =   "frmMain.frx":A73F
            TabIndex        =   48
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Default:"
            Height          =   195
            Index           =   35
            Left            =   3480
            TabIndex        =   47
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   705
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Sid-File Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   0
            TabIndex        =   45
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright:"
            Height          =   195
            Index           =   38
            Left            =   0
            TabIndex        =   44
            Top             =   735
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Index           =   36
            Left            =   0
            TabIndex        =   43
            Top             =   255
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Author:"
            Height          =   195
            Index           =   37
            Left            =   0
            TabIndex        =   42
            Top             =   495
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Songs:"
            Height          =   195
            Index           =   34
            Left            =   2160
            TabIndex        =   41
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   705
         End
         Begin VB.Label lbl_SFInfo_Name 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":AA49
            TabIndex        =   40
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SFInfo_Author 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":AD53
            TabIndex        =   39
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SFInfo_Copyright 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":B05D
            TabIndex        =   38
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SFInfo_Songs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3120
            MouseIcon       =   "frmMain.frx":B367
            TabIndex        =   37
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
      End
      Begin VB.Label lblSidNotFound2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SID FILE NOT FOUND!!!"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69405
         TabIndex        =   46
         Top             =   1080
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lbl_Info_VComment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         MouseIcon       =   "frmMain.frx":B671
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1965
         UseMnemonic     =   0   'False
         Width           =   9780
      End
      Begin VB.Line LineVCommentTop 
         BorderColor     =   &H80000010&
         X1              =   -74880
         X2              =   -65160
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line LineVCommentBot 
         BorderColor     =   &H80000014&
         X1              =   -74880
         X2              =   -65160
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line lineCommentBot 
         BorderColor     =   &H80000014&
         X1              =   240
         X2              =   9960
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line lineCommentTop 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   9960
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Label lbl_Info_Comment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmMain.frx":B97B
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1965
         UseMnemonic     =   0   'False
         Width           =   9780
      End
   End
   Begin TabDlg.SSTab tabSid 
      Height          =   2205
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   3889
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Sid Info"
      TabPicture(0)   =   "frmMain.frx":BC85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPrompt(41)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPrompt(40)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPrompt(39)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPrompt(42)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblSidNotFound"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_SInfo_Musician"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_SInfo_Nick"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_SInfo_Group"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_SInfo_Photo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraSidFile"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.Frame fraSidFile 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4200
         TabIndex        =   21
         Top             =   720
         Width           =   4575
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Default:"
            Height          =   195
            Index           =   45
            Left            =   3480
            TabIndex        =   50
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   585
         End
         Begin VB.Label lbl_SInfo_Default 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4440
            MouseIcon       =   "frmMain.frx":BD29
            TabIndex        =   49
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SInfo_Songs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3120
            MouseIcon       =   "frmMain.frx":C033
            TabIndex        =   35
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SInfo_Copyright 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":C33D
            TabIndex        =   34
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SInfo_Author 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":C647
            TabIndex        =   33
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lbl_SInfo_Name 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmMain.frx":C951
            TabIndex        =   32
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Songs:"
            Height          =   195
            Index           =   44
            Left            =   2160
            TabIndex        =   26
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   585
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Author:"
            Height          =   195
            Index           =   47
            Left            =   0
            TabIndex        =   25
            Top             =   495
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Index           =   46
            Left            =   0
            TabIndex        =   24
            Top             =   255
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright:"
            Height          =   195
            Index           =   48
            Left            =   0
            TabIndex        =   23
            Top             =   735
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Sid-File Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   43
            Left            =   0
            TabIndex        =   22
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1065
         End
      End
      Begin VB.Label lbl_SInfo_Photo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         MouseIcon       =   "frmMain.frx":CC5B
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lbl_SInfo_Group 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1200
         MouseIcon       =   "frmMain.frx":CF65
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lbl_SInfo_Nick 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1200
         MouseIcon       =   "frmMain.frx":D26F
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lbl_SInfo_Musician 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1200
         MouseIcon       =   "frmMain.frx":D579
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblSidNotFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   27
         Top             =   1080
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Photo:"
         Height          =   195
         Index           =   42
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Musician:"
         Height          =   195
         Index           =   39
         Left            =   240
         TabIndex        =   19
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick:"
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   18
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         Height          =   195
         Index           =   41
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   945
      End
   End
   Begin VB.Image imgCup 
      Height          =   165
      Left            =   8760
      Picture         =   "frmMain.frx":D883
      ToolTipText     =   "Classic Game!"
      Top             =   3495
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPrequel 
      Height          =   165
      Left            =   9000
      MouseIcon       =   "frmMain.frx":DC0D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":DF17
      ToolTipText     =   "Prequel..."
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRelated 
      Height          =   165
      Left            =   9240
      MouseIcon       =   "frmMain.frx":DF9A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":E2A4
      ToolTipText     =   "Related Game..."
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgSequel 
      Height          =   165
      Left            =   9720
      MouseIcon       =   "frmMain.frx":E329
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":E633
      ToolTipText     =   "Sequel..."
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCup2 
      Height          =   165
      Left            =   8520
      Picture         =   "frmMain.frx":E6B6
      ToolTipText     =   "Classic Game!"
      Top             =   3495
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRatingFace 
      Height          =   165
      Left            =   8280
      MouseIcon       =   "frmMain.frx":EA40
      MousePointer    =   99  'Custom
      Top             =   3495
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgAdult 
      Height          =   165
      Left            =   9960
      MouseIcon       =   "frmMain.frx":ED4A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":F054
      ToolTipText     =   "Sequel..."
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCloneOf 
      Height          =   165
      Left            =   9480
      MouseIcon       =   "frmMain.frx":F0DD
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":F3E7
      ToolTipText     =   "Sequel..."
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   51
      Left            =   7320
      MouseIcon       =   "frmMain.frx":F46B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":F775
      Top             =   3495
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      MouseIcon       =   "frmMain.frx":FAE2
      MousePointer    =   99  'Custom
      TabIndex        =   150
      Top             =   3495
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   50
      Left            =   1200
      Picture         =   "frmMain.frx":FDEC
      Top             =   9360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   49
      Left            =   1200
      Picture         =   "frmMain.frx":FEB2
      Top             =   9240
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   48
      Left            =   1200
      Picture         =   "frmMain.frx":FF7B
      Top             =   9120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   47
      Left            =   1200
      Picture         =   "frmMain.frx":10045
      Top             =   9000
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   46
      Left            =   1200
      Picture         =   "frmMain.frx":1010F
      Top             =   8880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   45
      Left            =   1200
      Picture         =   "frmMain.frx":101D8
      Top             =   8760
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   44
      Left            =   1200
      Picture         =   "frmMain.frx":102A1
      Top             =   8640
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   43
      Left            =   1200
      Picture         =   "frmMain.frx":1036A
      Top             =   8520
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   42
      Left            =   1200
      Picture         =   "frmMain.frx":10433
      Top             =   8400
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   41
      Left            =   1200
      Picture         =   "frmMain.frx":104FA
      Top             =   8280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   40
      Left            =   1200
      Picture         =   "frmMain.frx":105C1
      Top             =   8160
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   39
      Left            =   1200
      Picture         =   "frmMain.frx":10688
      Top             =   8040
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   38
      Left            =   1200
      Picture         =   "frmMain.frx":1074E
      Top             =   7920
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   37
      Left            =   1200
      Picture         =   "frmMain.frx":10814
      Top             =   7800
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   36
      Left            =   1200
      Picture         =   "frmMain.frx":108DA
      Top             =   7680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   35
      Left            =   1200
      Picture         =   "frmMain.frx":109A0
      Top             =   7560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   34
      Left            =   1200
      Picture         =   "frmMain.frx":10A67
      Top             =   7440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   33
      Left            =   1200
      Picture         =   "frmMain.frx":10B2F
      Top             =   7320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   32
      Left            =   1200
      Picture         =   "frmMain.frx":10BF7
      Top             =   7200
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   31
      Left            =   1200
      Picture         =   "frmMain.frx":10CC0
      Top             =   7080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   30
      Left            =   1200
      Picture         =   "frmMain.frx":10D89
      Top             =   6960
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   29
      Left            =   1200
      Picture         =   "frmMain.frx":10E52
      Top             =   6840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   28
      Left            =   1200
      Picture         =   "frmMain.frx":10F1A
      Top             =   6720
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   27
      Left            =   1200
      Picture         =   "frmMain.frx":10FE2
      Top             =   6600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   26
      Left            =   1200
      Picture         =   "frmMain.frx":110A9
      Top             =   6480
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   25
      Left            =   1200
      Picture         =   "frmMain.frx":11170
      Top             =   6360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   24
      Left            =   360
      Picture         =   "frmMain.frx":11237
      Top             =   9240
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   23
      Left            =   360
      Picture         =   "frmMain.frx":112FD
      Top             =   9120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   22
      Left            =   360
      Picture         =   "frmMain.frx":113C3
      Top             =   9000
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   21
      Left            =   360
      Picture         =   "frmMain.frx":11489
      Top             =   8880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   20
      Left            =   360
      Picture         =   "frmMain.frx":1154E
      Top             =   8760
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   19
      Left            =   360
      Picture         =   "frmMain.frx":11614
      Top             =   8640
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   18
      Left            =   360
      Picture         =   "frmMain.frx":116DA
      Top             =   8520
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   17
      Left            =   360
      Picture         =   "frmMain.frx":1179F
      Top             =   8400
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   16
      Left            =   360
      Picture         =   "frmMain.frx":11864
      Top             =   8280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   15
      Left            =   360
      Picture         =   "frmMain.frx":1192A
      Top             =   8160
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   14
      Left            =   360
      Picture         =   "frmMain.frx":119EF
      Top             =   8040
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   13
      Left            =   360
      Picture         =   "frmMain.frx":11AB4
      Top             =   7920
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   12
      Left            =   360
      Picture         =   "frmMain.frx":11B79
      Top             =   7800
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   11
      Left            =   360
      Picture         =   "frmMain.frx":11C3E
      Top             =   7680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   10
      Left            =   360
      Picture         =   "frmMain.frx":11D02
      Top             =   7560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   9
      Left            =   360
      Picture         =   "frmMain.frx":11DC6
      Top             =   7440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   8
      Left            =   360
      Picture         =   "frmMain.frx":11E89
      Top             =   7320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   7
      Left            =   360
      Picture         =   "frmMain.frx":11F4D
      Top             =   7200
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   6
      Left            =   360
      Picture         =   "frmMain.frx":12010
      Top             =   7080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   5
      Left            =   360
      Picture         =   "frmMain.frx":120D3
      Top             =   6960
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   4
      Left            =   360
      Picture         =   "frmMain.frx":12195
      Top             =   6840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   3
      Left            =   360
      Picture         =   "frmMain.frx":12258
      Top             =   6720
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   2
      Left            =   360
      Picture         =   "frmMain.frx":12319
      Top             =   6600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":123D9
      Top             =   6480
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRRStore 
      Height          =   165
      Index           =   0
      Left            =   360
      Picture         =   "frmMain.frx":12497
      Top             =   6360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblSplitter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      TabIndex        =   140
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgNoCloneOf 
      Enabled         =   0   'False
      Height          =   165
      Left            =   6120
      Picture         =   "frmMain.frx":12804
      Top             =   10560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCloneOfStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5880
      Picture         =   "frmMain.frx":12882
      Top             =   10800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblWebLink 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3000
      TabIndex        =   58
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTab3N 
      Height          =   240
      Left            =   5520
      Picture         =   "frmMain.frx":12906
      Top             =   11520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgTab3Y 
      Enabled         =   0   'False
      Height          =   240
      Left            =   5160
      Picture         =   "frmMain.frx":129A6
      Top             =   11520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAdultStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5400
      Picture         =   "frmMain.frx":12A58
      Top             =   11040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNoAdultStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5160
      Picture         =   "frmMain.frx":12AE1
      Top             =   11040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgListMusic 
      Enabled         =   0   'False
      Height          =   240
      Left            =   7440
      Picture         =   "frmMain.frx":12B5F
      Top             =   10680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgListGameHasNone 
      Enabled         =   0   'False
      Height          =   240
      Left            =   7200
      Picture         =   "frmMain.frx":12BEB
      Top             =   10680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgListGameHasGame 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6720
      Picture         =   "frmMain.frx":12C53
      Top             =   10680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgListGameHasMusic 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6960
      Picture         =   "frmMain.frx":12CDD
      Top             =   10680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgListGameHasBoth 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6480
      Picture         =   "frmMain.frx":12D64
      Top             =   10680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgLoad 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6480
      Picture         =   "frmMain.frx":12E0A
      Top             =   10320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRelatedStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5640
      Picture         =   "frmMain.frx":12EA0
      Top             =   10800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgSequelStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5400
      Picture         =   "frmMain.frx":12F25
      Top             =   10800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgPrequelStore 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5160
      Picture         =   "frmMain.frx":12FA8
      Top             =   10800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNoSequel 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5640
      Picture         =   "frmMain.frx":1302B
      Top             =   10560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNoPrequel 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5400
      Picture         =   "frmMain.frx":130A9
      Top             =   10560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNoRelated 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5160
      Picture         =   "frmMain.frx":13126
      Top             =   10560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNoRating 
      Enabled         =   0   'False
      Height          =   165
      Left            =   5880
      Picture         =   "frmMain.frx":131A3
      Top             =   10560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRecent 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   1
      Left            =   7320
      Picture         =   "frmMain.frx":1321F
      Top             =   9840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRecent 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   0
      Left            =   7080
      Picture         =   "frmMain.frx":132AD
      Top             =   9840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgZoomCur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   240
      Left            =   6720
      MouseIcon       =   "frmMain.frx":13337
      MousePointer    =   99  'Custom
      Top             =   9840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFace 
      Enabled         =   0   'False
      Height          =   165
      Index           =   4
      Left            =   6120
      Picture         =   "frmMain.frx":13641
      Top             =   10320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgFace 
      Enabled         =   0   'False
      Height          =   165
      Index           =   3
      Left            =   5880
      Picture         =   "frmMain.frx":136C7
      Top             =   10320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgFace 
      Enabled         =   0   'False
      Height          =   165
      Index           =   2
      Left            =   5640
      Picture         =   "frmMain.frx":1374C
      Top             =   10320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgFace 
      Enabled         =   0   'False
      Height          =   165
      Index           =   1
      Left            =   5400
      Picture         =   "frmMain.frx":137D0
      Top             =   10320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgFace 
      Enabled         =   0   'False
      Height          =   165
      Index           =   0
      Left            =   5160
      Picture         =   "frmMain.frx":13855
      Top             =   10320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDummySS 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgNoPhoto 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1500
      Left            =   6000
      Picture         =   "frmMain.frx":138DC
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   4560
      MouseIcon       =   "frmMain.frx":14E5B
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgCurDown 
      Height          =   375
      Left            =   5880
      MouseIcon       =   "frmMain.frx":15165
      MousePointer    =   99  'Custom
      Top             =   9840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgCurUp 
      Height          =   375
      Left            =   5520
      MouseIcon       =   "frmMain.frx":1546F
      MousePointer    =   99  'Custom
      Top             =   9840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add"
         Begin VB.Menu mnuFileAddGame 
            Caption         =   "&Game..."
            Shortcut        =   +{INSERT}
         End
         Begin VB.Menu mnuFileAddSid 
            Caption         =   "&Sid to Music Table..."
            Shortcut        =   ^{INSERT}
         End
      End
      Begin VB.Menu mnuFileRem 
         Caption         =   "&Remove"
         Begin VB.Menu mnuFileRemove 
            Caption         =   "&Selected Game/Sid..."
            Shortcut        =   {DEL}
         End
         Begin VB.Menu mnuFileRemoveAll 
            Caption         =   "&All Games/Sids in Current View..."
            Enabled         =   0   'False
            Shortcut        =   +{DEL}
         End
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditGame 
         Caption         =   "G&ame Info"
         Enabled         =   0   'False
         Begin VB.Menu mnuEditGameCurrent 
            Caption         =   "&Current game..."
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuEditGameAll 
            Caption         =   "&All games in current view..."
            Shortcut        =   +^{F1}
         End
      End
      Begin VB.Menu mnuEditVersion 
         Caption         =   "&Version Info"
         Enabled         =   0   'False
         Begin VB.Menu mnuEditVersionCurrent 
            Caption         =   "&Current Game..."
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuEditVersionAll 
            Caption         =   "&All games in current view..."
            Shortcut        =   +^{F2}
         End
      End
      Begin VB.Menu mnuEditPersonal 
         Caption         =   "&Personal"
         Enabled         =   0   'False
         Begin VB.Menu mnuEditPersonalCurrent 
            Caption         =   "&Current Game..."
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuEditPersonalAll 
            Caption         =   "&All games in current view..."
            Shortcut        =   +^{F3}
         End
      End
      Begin VB.Menu mnuEditBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGameFile 
         Caption         =   "&Game File..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSidFile 
         Caption         =   "&Music File..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditScrnShots 
         Caption         =   "&Pictures..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEditExtras 
         Caption         =   "&Extras..."
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEditBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGemus 
         Caption         =   "GEMUS ""&key=value"" pairs..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMark 
         Caption         =   "&Mark Game for Linking"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditLinkSet 
         Caption         =   "Set Game &Link..."
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditLinkClear 
         Caption         =   "&Clear Game Link..."
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find in List..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindID 
         Caption         =   "Find by ID"
         Begin VB.Menu mnuEditFindIDGame 
            Caption         =   "&Game..."
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu mnuEditFindIDMusic 
            Caption         =   "&Music..."
            Shortcut        =   ^{F9}
         End
      End
      Begin VB.Menu mnuEditBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFav 
         Caption         =   "Add to &Favourites"
         Enabled         =   0   'False
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewInfo 
         Caption         =   "Show &Game Details"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuViewBar15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSidebar 
         Caption         =   "Hide Sidebar"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewSidebar 
         Caption         =   "Show Sidebar on Left"
         Index           =   1
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuViewSidebar 
         Caption         =   "Show Sidebar on Right"
         Index           =   2
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuViewBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewScreenshots 
         Caption         =   "&Hide Screenshots"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuViewScreenshots 
         Caption         =   "&Small Screenshots"
         Index           =   1
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuViewScreenshots 
         Caption         =   "&Full Size Screenshots"
         Index           =   2
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuViewBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSSSnap 
         Caption         =   "S&nap to Picture Size"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuViewBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewExtraColumns 
         Caption         =   "Extra Columns..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuViewBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewListList 
         Caption         =   "&List view"
      End
      Begin VB.Menu mnuViewListDetails 
         Caption         =   "&Details view"
      End
      Begin VB.Menu mnuViewBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFontCols 
         Caption         =   "List F&ont/Colours..."
      End
      Begin VB.Menu mnuViewBar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAddView 
         Caption         =   "&Add View..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuViewEditView 
         Caption         =   "&Edit View..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuViewRemoveView 
         Caption         =   "&Remove View..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewBar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewQuickFilter 
         Caption         =   "&Quick Filter..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuViewBar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCategories 
         Caption         =   "Ca&tegory Window"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewBar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFocusToList 
         Caption         =   "Set Focus to List"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuViewBar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewExportToTextFile 
         Caption         =   "E&xport to Text File"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuViewExportToRomLister 
         Caption         =   "Export to Ro&mLister..."
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuGemus 
      Caption         =   "&GEMUS"
      Begin VB.Menu mnuGemusEmu 
         Caption         =   "EmuList"
         Index           =   0
         Shortcut        =   +{F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   1
         Shortcut        =   +{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   2
         Shortcut        =   +{F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   3
         Shortcut        =   +{F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   4
         Shortcut        =   +{F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   5
         Shortcut        =   +{F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   6
         Shortcut        =   +{F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   7
         Shortcut        =   +{F8}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusEmu 
         Caption         =   ""
         Index           =   8
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGemusBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGemusManage 
         Caption         =   "&Manage Emulators..."
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuGemusBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGemusFileTypes 
         Caption         =   "&Supported Emulator File Types..."
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuGemusFields 
         Caption         =   "Set Default GEMUS Script &Fields..."
      End
      Begin VB.Menu mnuGemusBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGemusWizard 
         Caption         =   "Emulator Setup &Wizard..."
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu mnuGemusBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGemusExport 
         Caption         =   "&Export Emulator Definitions"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSounds 
         Caption         =   "Play &Sound on Classic Game"
      End
      Begin VB.Menu mnuToolsNSM 
         Caption         =   "Non Stop Music"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuToolsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsHideAdult 
         Caption         =   "&Hide Adult Games/Sids..."
      End
      Begin VB.Menu mnuToolsAdultPassword 
         Caption         =   "&Reset Password..."
      End
      Begin VB.Menu mnuToolsBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuToolsBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsDatabaseOptions 
         Caption         =   "Database Options..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuToolsPaths 
         Caption         =   "&Paths..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuToolsBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsVerify 
         Caption         =   "&Verify Available Files..."
      End
      Begin VB.Menu mnuToolsBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsImportPersonal 
         Caption         =   "&Import Personal Info..."
      End
   End
   Begin VB.Menu mnuRecent 
      Caption         =   "&Recent"
      Begin VB.Menu mnuRec 
         Caption         =   "0"
         Index           =   0
      End
      Begin VB.Menu mnuRec 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu mnuRec 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu mnuRec 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu mnuRec 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu mnuRec 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu mnuRec 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu mnuRec 
         Caption         =   "7"
         Index           =   7
      End
      Begin VB.Menu mnuRec 
         Caption         =   "8"
         Index           =   8
      End
      Begin VB.Menu mnuRec 
         Caption         =   "9"
         Index           =   9
      End
   End
   Begin VB.Menu mnuGBPick 
      Caption         =   "Game&Base"
      Begin VB.Menu mnuGB 
         Caption         =   "GB"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGamebase64 
         Caption         =   "&Gamebase64 Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHome 
         Caption         =   "&Gamebase64 Home Page"
      End
      Begin VB.Menu mnuHelpEmail 
         Caption         =   "&Email the Author"
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpDB 
         Caption         =   "Database Help..."
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpDB 
         Caption         =   "Database Website"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpDB 
         Caption         =   "Database Email"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu pmnuViews 
      Caption         =   "PopupViews"
      Visible         =   0   'False
      Begin VB.Menu pmnuView 
         Caption         =   "&Add View..."
         Index           =   0
      End
      Begin VB.Menu pmnuView 
         Caption         =   "&Edit View..."
         Index           =   1
      End
      Begin VB.Menu pmnuView 
         Caption         =   "&Remove View"
         Index           =   2
      End
      Begin VB.Menu pmnuView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu pmnuView 
         Caption         =   "&Quick Filter..."
         Index           =   4
      End
      Begin VB.Menu pmnuView 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu pmnuView 
         Caption         =   "&Sort Custom Views..."
         Index           =   6
      End
   End
   Begin VB.Menu pmnuSidPlayers 
      Caption         =   "PopupSidPlayers"
      Visible         =   0   'False
      Begin VB.Menu pmnuSidPlayer 
         Caption         =   "Default Sid Player"
         Index           =   0
      End
   End
   Begin VB.Menu pmnuEmulators 
      Caption         =   "PopupEmulators"
      Visible         =   0   'False
      Begin VB.Menu pmnuEmulator 
         Caption         =   "Default Emulator"
         Index           =   0
      End
   End
   Begin VB.Menu pmnuScrnShots 
      Caption         =   "PopupScrnShots"
      Visible         =   0   'False
      Begin VB.Menu pmnuScrnShot 
         Caption         =   "Zoom In"
         Index           =   0
      End
      Begin VB.Menu pmnuScrnShot 
         Caption         =   "Change Screenshot(s)..."
         Index           =   1
      End
   End
   Begin VB.Menu pmnuExtras 
      Caption         =   "PopupExtras"
      Visible         =   0   'False
      Begin VB.Menu pmnuExtra 
         Caption         =   "Extra"
         Index           =   0
      End
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
' Copyright (C) 1999-2010 James Burrows
' Email: gamebase.frontend@leesti.com
' Homepage: http://www.bu22.com/
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

Private ma_strGameBaseNames() As String

Private m_xobjSQLEngine As GBSql.clsQuery

Private Const mc_lngSECCOLHIDDEN As Long = -1
Private Const mc_strMENUBAR As String = "-"

Private Type CVFilterFixedType
    strFieldTable As String * 128
    strFieldName As String * 128
    strMusicFieldTable As String * 128
    strMusicFieldName As String * 128
    eOperator As FilterOperatorEnum
    eClauseType As FilterClauseTypeEnum
    strClauseData As String * 64
End Type
Private Type CVFilterFixedBufferType
    strBuffer As String * 580
End Type

Private m_blnNotesCut As Boolean
Private m_blnNotesCopy As Boolean
Private m_blnNotesPaste As Boolean

Private m_strPrequelTTT As String
Private m_strSequelTTT As String
Private m_strRelatedTTT As String
Private m_strCloneofTTT As String

'for window resizing
Private Type WindowSizeType
    lngTwipsX As Long
    lngTwipsY As Long
    sngWidth_Twips As Single
    lngWidth_Pixels As Long
    sngScaleWidth_Twips As Single
    lngScaleWidth_Pixels As Long
    sngFullSSHeight_Twips As Single
    lngFullSSHeight_Pixels As Long
    sngSmallSSHeight_Twips As Single
    lngSmallSSHeight_Pixels As Long
    sngFullSSWidth_Twips As Single
    lngFullSSWidth_Pixels As Long
    sngSmallSSWidth_Twips As Single
    lngSmallSSWidth_Pixels As Long
    sngHeight_FullSSDetails_Twips As Single
    sngHeight_FullSSNoDetails_Twips As Single
    sngHeight_SmallSSDetails_Twips As Single
    sngHeight_SmallSSNoDetails_Twips As Single
    sngScaleHeight_FullSSDetails_Twips As Single
    sngScaleHeight_FullSSNoDetails_Twips As Single
    sngScaleHeight_SmallSSDetails_Twips As Single
    sngScaleHeight_SmallSSNoDetails_Twips As Single
    lngHeight_FullSSDetails_Pixels As Long
    lngHeight_FullSSNoDetails_Pixels As Long
    lngHeight_SmallSSDetails_Pixels As Long
    lngHeight_SmallSSNoDetails_Pixels As Long
    sngScaleHeightLessSSDetails_Twips As Single
    sngScaleHeightLessSSNoDetails_Twips As Single
    lngScreenHeight_Pixels As Long
    lngScreenWidth_Pixels As Long
End Type
Private m_udtMinSizes As WindowSizeType
Private Const mc_TwipsBetweenTABDETAILSAndSBSTATUS As Single = 10
Private Const mc_TwipsBetweenFRAPANELAndTABDETAILS As Single = 405
Private Const mc_TwipsBetweenLVMAINAndFRAPANEL As Single = 45
Private Const mc_TwipsBetweenFRAPANELAndSBSTATUS As Single = 375 '15
Private Const mc_TwipsBetweenLVMAINAndpicSSMain As Single = 45
Private m_sngScrnShotWidth As Single
Private m_sngScrnShotHeight As Single
Private m_blnResizeEnable As Boolean
Private m_sngRegWidth As Single
Private m_sngRegHeight As Single
Private m_sngRegTop As Single
Private m_sngRegLeft As Single
Private m_blnRegMaximized As Boolean
Private m_blnRegDetails As Boolean

Private Const mc_SCRNSHOTSHIDE As Long = 0
Private Const mc_SCRNSHOTSSMALL As Long = 1
Private Const mc_SCRNSHOTSBIG As Long = 2

Private m_ScrnShotNavButtonWidth As Long
Private m_ScrnShotNavButtonHeight As Long

Private SoundArray() As Byte    'for playing default sound from resource file

'ui strings
Public m_strAddToFavs As String
Public m_strRemoveFromFavs As String

'find/find next/find by id
Private m_strFindInList As String
Private m_lngFindCommentInList As Long
Private m_lngFindIDGame As Long
Private m_lngFindIDMusic As Long

Private m_lngRegScrnShots As Long
Private m_sngStatusBarOffset As Single
Private m_sngCachedPhotoHeight As Single
Private m_sngCachedPhotoWidth As Single

Private m_blnIsDraggingSidebarSplitter As Boolean
Private m_lngRegSidebar As Long
Private m_sngSplitterWidth As Single
Private m_sngSidebarWidth As Single
Private Const mc_sngDEFAULTSIDEBARWIDTH As Single = 1575
Private Const mc_sngMINSIDEBARWIDTH As Single = 500
Private m_sngPreviousSplitterLeft As Single
Private m_sngLastListViewSideBarAvailableWidth As Single

'recent menu
Private Type RecentMenuType
    blnIsGame As Boolean
    strKey As String
    strName As String
End Type
Private ma_udtRecent(9) As RecentMenuType

'general flags/variables
Private m_blnFormIsLoading As Boolean
Private m_blnActivated As Boolean
Private m_blnDoneTabDetails As Boolean
Private m_blnLastScrnShotWasNone As Boolean
Private m_blnLastPhotoWasNone As Boolean
Private m_lngLastViewIdx As Long
Private m_blnSetTheView As Boolean
Private m_blnSaveViewPrefs As Boolean
Private m_blnLastViewIsQuick As Boolean
Private m_blnDragging As Boolean
Private m_strDragFile As String
Private m_blnNotesChanged As Boolean
Private m_blnNotesByCode As Boolean
Private m_ShowingPhoto As Boolean

'sounds on classic games
Public m_blnRegSounds As Boolean
Private m_strLastSelectedKey As String

'non stop music
Private m_blnRegNonStopMusic As Boolean
Private m_blnSettingView As Boolean
Private m_strLastSidFilePlayed As String

'marked games
Private Type MarkedGameType
    strName As String
    lngID As Long
End Type
Private m_udtMarkedGame As MarkedGameType

'picture viewer
Private ma_strScrnShots() As String
Private m_lngScrnShotCount As Long
Private m_lngScrnShotLoaded As Long
Private m_lngScrnShotVisible As Long
Private m_lngScrnShotMaxVisible As Long
Private m_lngScrnShotTop As Long
Private m_blnScrnShotNone As Boolean
Public m_blnScrnShotBugWorkaround As Boolean
Private m_longScrnShotZoomIdx As Long
Private m_blnPicExists As Boolean

'gfx on the details tabs
Private m_sngTabIconsTop As Single
Private m_sngTabIconsStart As Single
Private m_sngTabIconsX As Single
Private m_sngTabIconsGap As Single
Private m_sngTabIconsCupGap As Single
Private m_sngHalfTabWidth As Single
Private m_sngCup1X As Single
Private m_sngCup2X As Single
Private m_sngTabIconsCupTop As Single
Private m_sngTabIconsCupWidth As Single

'error messages
Private m_strErrorItemNotFound As String
Private m_strErrorItemNotFoundCurrent As String
Private m_strErrorRequestRefresh As String

'classic game sounds array
Private ma_strSounds() As String
Private m_lngMaxSounds As Long

'for filtered views
Private Enum FilteredViewTypeEnum
    e_FilteredView_ID
    e_FilteredView_Text
    e_FilteredView_Other
End Enum
Private Type FilteredViewType
    eFilterType As FilteredViewTypeEnum
    strFieldName As String
    strFieldTable As String
    strMusicFieldName As String
    strMusicFieldTable As String
    strClauseData As String
    strTitle As String
End Type

'help menu
Private m_strDBHelpFile As String
Private m_strDBHelpEmail As String
Private m_strDBHelpURL As String

'sidebar quicklinks
Private m_strQLPlayGame As String
Private m_strQLPlayMusic As String
Private Type QuickLinksIconType
    strExtension As String
    lngIconIdx As Long
End Type
Private ma_udtQLNameIcons() As QuickLinksIconType
Private m_lngTotalQLNameIcons As Long
Private ma_udtQLIcons() As QuickLinksIconType
Private m_lngTotalQLIcons As Long
Private m_lngQLIconIdxPlayGame As Long
Private m_lngQLIconIdxPlayMusic As Long
Private m_lngQLIconIdxURLExtra As Long
Private m_lngQLIconIdxGemusExtra As Long
Private m_lngQLIconIdxMusicExtra As Long
Private m_lngQLIconIdxDefaultExtra As Long
Private m_lngQLIconIdxPrequel As Long
Private m_lngQLIconIdxSequel As Long
Private m_lngQLIconIdxRelated As Long
Private m_lngQLIconIdxCloneOf As Long
Private m_lngQLIconIdxWeblink As Long
Private m_lngQLIconIdxVersionWeblink As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Function FindQLIconByExtension(strLowerCaseExtension_in As String)

    Dim lngIdx As Long
    
    For lngIdx = 0 To m_lngTotalQLIcons - 1
        If ma_udtQLIcons(lngIdx).strExtension = strLowerCaseExtension_in Then
            FindQLIconByExtension = lngIdx
            Exit Function
        End If
    Next lngIdx
    FindQLIconByExtension = -1

End Function
Private Function FindQLNameIconByName(strLowerCaseName_in As String)

    Dim lngIdx As Long
    
    For lngIdx = 0 To m_lngTotalQLNameIcons - 1
        If ma_udtQLNameIcons(lngIdx).strExtension = strLowerCaseName_in Then
            FindQLNameIconByName = lngIdx
            Exit Function
        End If
    Next lngIdx
    FindQLNameIconByName = -1

End Function
Public Sub GetFirstLoadStuff(strMessage_out As String, blnGemus_out As Boolean)

    On Error GoTo Err_Handler

    
    g_dbGamebase.Execute ("UPDATE Config SET FirstLoadMessage="""" WHERE FirstLoadMessage IS NULL")
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT FirstLoadMessage, FirstLoadGemusAsk FROM Config", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            strMessage_out = !FirstLoadMessage
            blnGemus_out = (!FirstLoadGemusAsk <> 0)
            .Close
        End If
    End With
    Set g_rsGames = Nothing

    strMessage_out = Trim$(strMessage_out)
    If Len(strMessage_out) > gc_lngMAXMEMOTEXTLENGTH Then
        strMessage_out = Left$(strMessage_out, gc_lngMAXMEMOTEXTLENGTH)
    End If

    Exit Sub

Err_Handler:
    strMessage_out = ""
    blnGemus_out = False
    
End Sub
Public Sub LoadDatabaseConfig()

    On Error GoTo Err_Handler

    g_dbGamebase.Execute ("UPDATE Config SET DatabaseName="""" WHERE DatabaseName IS NULL")
    g_dbGamebase.Execute ("UPDATE Config SET GameBaseWindowTitle="""" WHERE GameBaseWindowTitle IS NULL")
    
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT DatabaseName, GameBaseWindowTitle FROM Config", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            If Trim$(!DatabaseName) <> "" Then
                g_strDatabaseName = Trim$(!DatabaseName)
            End If
            If Trim$(!GameBaseWindowTitle) <> "" Then
                g_strWINDOWCAPTION = Trim$(!GameBaseWindowTitle)
            End If
            .Close
        End If
    End With
    Set g_rsGames = Nothing

    mnuToolsDatabaseOptions.Caption = "'" & g_strDatabaseName & "' " & g_xobjFunc.GetINIMessage("MainMenus", "Menu_Tools_Database_Options", "Op&tions") & "..."
    mnuToolsPaths.Caption = "'" & g_strDatabaseName & "' " & g_xobjFunc.GetINIMessage("MainMenus", "Menu_Tools_Paths", "&Paths") & "..."

    Exit Sub

Err_Handler:
    
End Sub
Private Sub LoadCustomGfx()

    Dim lngIdx As Long
    Dim strPathWBS As String
    Dim strPath As String
    Dim strFile As String
    Dim strLower As String
    Dim lngArrayIdx As Long
    Dim lngDotPos As Long
    Dim a_strExtensions() As String
    Dim lngTotalExtensions As Long
    Dim lngExtensionIdx As Long

    'reads in the GFX from the GameBase/Gfx folder

    'NoPhoto_100x100.gif (not Transp)
    'PictureDown_15x10.gif (Transp = any colour)
    'PictureUp_15x10.gif (Transp = any colour)

    'ListGameHasBoth_16x16.gif (mask purple #FF00FF)
    'ListGameHasGame_16x16.gif (mask purple #FF00FF)
    'ListGameHasMusic_16x16.gif (mask purple #FF00FF)
    'ListGameHasNone_16x16.gif (mask purple #FF00FF)
    'ListMusic_16x16.gif (mask purple #FF00FF)

    'ClassicGame_16x11.gif (Transp = any colour)

    'NoPrecursor_13x11.gif (Transp = any colour)
    'NoSequel_13x11.gif (Transp = any colour)
    'NoRelated_13x11.gif (Transp = any colour)
    'NoCloneOf_13x11.gif (Transp = any colour)
    'Related_13x11.gif (Transp = any colour)
    'CloneOf_13x11.gif (Transp = any colour)
    'Sequel_13x11.gif (Transp = any colour)
    'Precursor_13x11.gif (Transp = any colour)

    'NotAdult_13x11.gif (Transp = any colour)
    'Adult_13x11.gif (Transp = any colour)

    'VeryGoodRating_13x11.gif (Transp = any colour)
    'QuiteGoodRating_13x11.gif (Transp = any colour)
    'AverageRating_13x11.gif (Transp = any colour)
    'PoorRating_13x11.gif (Transp = any colour)
    'TerribleRating_13x11.gif (Transp = any colour)
    'UnknownRating_13x11.gif (Transp = any colour)

    'GameTab1_16x16.gif (Transp = any colour)
    'GameTab2_16x16.gif (Transp = any colour)
    'GameTab3_16x16.gif (Transp = any colour)
    'GameTab4_16x16.gif (Transp = any colour)
    'GameTab4a_16x16.gif (Transp = any colour)
    'MusicTab_16x16.gif (Transp = any colour)

    'RecentMenuGame_13x13.gif (Transp = White #FFFFFF)
    'RecentMenuMusic_13x13.gif (Transp = White #FFFFFF)

    'no musician photo image
    Call LoadGBImageFromFile(imgNoPhoto, g_strGameBasePathWBS & "Gfx\NoPhoto_100x100.gif")

    'picture up and down images
    Call LoadGBImageFromFile(imgSSUp, g_strGameBasePathWBS & "Gfx\PictureUp_15x10.gif")
    Call LoadGBImageFromFile(imgSSDown, g_strGameBasePathWBS & "Gfx\PictureDown_15x10.gif")

    'listview icons
    Call LoadGBImageFromFile(imgListGameHasBoth, g_strGameBasePathWBS & "Gfx\ListGameHasBoth_16x16.gif")
    Call LoadGBImageFromFile(imgListGameHasGame, g_strGameBasePathWBS & "Gfx\ListGameHasGame_16x16.gif")
    Call LoadGBImageFromFile(imgListGameHasMusic, g_strGameBasePathWBS & "Gfx\ListGameHasMusic_16x16.gif")
    Call LoadGBImageFromFile(imgListGameHasNone, g_strGameBasePathWBS & "Gfx\ListGameHasNone_16x16.gif")
    Call LoadGBImageFromFile(imgListMusic, g_strGameBasePathWBS & "Gfx\ListMusic_16x16.gif")
    With ilstMain16
        .ListImages.Add 1, , imgListGameHasBoth.Picture
        .ListImages.Add 2, , imgListGameHasGame.Picture
        .ListImages.Add 3, , imgListGameHasMusic.Picture
        .ListImages.Add 4, , imgListGameHasNone.Picture
        .ListImages.Add 5, , imgListMusic.Picture
    End With

    'classic game image
    If LoadGBImageFromFile(imgCup, g_strGameBasePathWBS & "Gfx\ClassicGame_16x11.gif") = True Then
        imgCup2.Picture = imgCup.Picture
    End If

    'review rating bars
    For lngIdx = 0 To 100 Step 2
        Call LoadGBImageFromFile(imgRRStore(GetReviewRatingImgIdx(lngIdx)), g_strGameBasePathWBS & "Gfx\reviewratings\" & lngIdx & ".gif")
    Next lngIdx

    'Precursor/sequel/related
    Call LoadGBImageFromFile(imgNoPrequel, g_strGameBasePathWBS & "Gfx\NoPrecursor_13x11.gif")
    Call LoadGBImageFromFile(imgNoSequel, g_strGameBasePathWBS & "Gfx\NoSequel_13x11.gif")
    Call LoadGBImageFromFile(imgNoRelated, g_strGameBasePathWBS & "Gfx\NoRelated_13x11.gif")
    Call LoadGBImageFromFile(imgNoCloneOf, g_strGameBasePathWBS & "Gfx\NoCloneOf_13x11.gif")
    Call LoadGBImageFromFile(imgRelatedStore, g_strGameBasePathWBS & "Gfx\Related_13x11.gif")
    Call LoadGBImageFromFile(imgSequelStore, g_strGameBasePathWBS & "Gfx\Sequel_13x11.gif")
    Call LoadGBImageFromFile(imgPrequelStore, g_strGameBasePathWBS & "Gfx\Precursor_13x11.gif")
    Call LoadGBImageFromFile(imgCloneOfStore, g_strGameBasePathWBS & "Gfx\CloneOf_13x11.gif")

    'adult/not adult images
    Call LoadGBImageFromFile(imgNoAdultStore, g_strGameBasePathWBS & "Gfx\NotAdult_13x11.gif")
    Call LoadGBImageFromFile(imgAdultStore, g_strGameBasePathWBS & "Gfx\Adult_13x11.gif")

    'game rating images
    Call LoadGBImageFromFile(imgFace(4), g_strGameBasePathWBS & "Gfx\VeryGoodRating_13x11.gif")
    Call LoadGBImageFromFile(imgFace(3), g_strGameBasePathWBS & "Gfx\QuiteGoodRating_13x11.gif")
    Call LoadGBImageFromFile(imgFace(2), g_strGameBasePathWBS & "Gfx\AverageRating_13x11.gif")
    Call LoadGBImageFromFile(imgFace(1), g_strGameBasePathWBS & "Gfx\PoorRating_13x11.gif")
    Call LoadGBImageFromFile(imgFace(0), g_strGameBasePathWBS & "Gfx\TerribleRating_13x11.gif")
    Call LoadGBImageFromFile(imgNoRating, g_strGameBasePathWBS & "Gfx\UnknownRating_13x11.gif")

    'details tab caption images
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\GameTab1_16x16.gif") = True Then tabDetails.TabPicture(0) = imgLoad.Picture
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\GameTab2_16x16.gif") = True Then tabDetails.TabPicture(1) = imgLoad.Picture
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\GameTab3_16x16.gif") = True Then tabDetails.TabPicture(2) = imgLoad.Picture
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\GameTab4_16x16.gif") = True Then
        tabDetails.TabPicture(3) = imgLoad.Picture
        imgTab3N.Picture = imgLoad.Picture
        imgTab3Y.Picture = imgLoad.Picture
    End If
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\GameTab4a_16x16.gif") = True Then imgTab3Y.Picture = imgLoad.Picture
    If LoadGBImageFromFile(imgLoad, g_strGameBasePathWBS & "Gfx\MusicTab_16x16.gif") = True Then tabDetails.TabPicture(4) = imgLoad.Picture

    'recent menu images
    Call LoadGBImageFromFile(imgRecent(1), g_strGameBasePathWBS & "Gfx\RecentMenuGame_13x13.gif")
    Call LoadGBImageFromFile(imgRecent(0), g_strGameBasePathWBS & "Gfx\RecentMenuMusic_13x13.gif")

    'quick link icons
    strPath = g_strGameBasePathWBS & "Gfx\quicklinks"
    strPathWBS = strPath & "\"
    With imgListQL.ListImages
        lngIdx = .Count + 1
        'in-built quicklinks
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamelink_cloneof.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxCloneOf = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamelink_prequel.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxPrequel = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamelink_related.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxRelated = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamelink_sequel.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxSequel = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamefield_weblink.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxWeblink = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_gamefield_version_weblink.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxVersionWeblink = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_play_game.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxPlayGame = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_play_music.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxPlayMusic = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_extra_default.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxDefaultExtra = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_extra_gemus.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxGemusExtra = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_extra_music.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxMusicExtra = lngIdx
            lngIdx = lngIdx + 1
        End If
        If LoadGBImageFromFile(imgLoad, strPathWBS & "_extra_url.gif") Then
            .Add lngIdx, , imgLoad.Picture
            m_lngQLIconIdxURLExtra = lngIdx
            lngIdx = lngIdx + 1
        End If
        'named extras, and specified file extensions
        If g_xobjFunc.DirExists(strPath) Then
            strFile = Dir(strPathWBS & "*.gif", vbNormal Or vbReadOnly)
            Do Until strFile = ""
                lngDotPos = InStrRev(strFile, ".")
                If Left$(strFile, 1) <> "_" And lngDotPos > 1 Then
                    If LoadGBImageFromFile(imgLoad, strPathWBS & strFile) Then
                    
                        .Add lngIdx, , imgLoad.Picture

                        'load lowercase version into named array
                        strLower = LCase$(Left$(strFile, lngDotPos - 1))
                        lngArrayIdx = FindQLNameIconByName(strLower)
                        If lngArrayIdx < 0 Then
                            ReDim Preserve ma_udtQLNameIcons(m_lngTotalQLNameIcons)
                            With ma_udtQLNameIcons(m_lngTotalQLNameIcons)
                                .strExtension = strLower
                                .lngIconIdx = lngIdx
                            End With
                            m_lngTotalQLNameIcons = m_lngTotalQLNameIcons + 1
                        Else
                            ma_udtQLNameIcons(lngArrayIdx).lngIconIdx = lngIdx
                        End If

                        'split into file extensions and load into file extension array
                        a_strExtensions = Split(strLower, ".")
                        lngTotalExtensions = UBound(a_strExtensions) + 1
                        For lngExtensionIdx = 0 To lngTotalExtensions - 1
                            strLower = a_strExtensions(lngExtensionIdx)
                            lngArrayIdx = FindQLIconByExtension(strLower)
                            If lngArrayIdx < 0 Then
                                ReDim Preserve ma_udtQLIcons(m_lngTotalQLIcons)
                                With ma_udtQLIcons(m_lngTotalQLIcons)
                                    .strExtension = strLower
                                    .lngIconIdx = lngIdx
                                End With
                                m_lngTotalQLIcons = m_lngTotalQLIcons + 1
                            Else
                                ma_udtQLIcons(lngArrayIdx).lngIconIdx = lngIdx
                            End If
                        Next lngExtensionIdx

                        lngIdx = lngIdx + 1

                    End If
                End If
                strFile = Dir()
            Loop
        End If
    End With
    
End Sub
Private Function GetResourceFileAsString(ByVal lngResID_in As Long) As String

    Dim a_bytFile() As Byte
    Dim lngIdx As Long
    Dim lngLastByte As Long

    GetResourceFileAsString = ""
    a_bytFile = LoadResData(lngResID_in, "CUSTOM")
    lngLastByte = UBound(a_bytFile)
    For lngIdx = 0 To lngLastByte
        GetResourceFileAsString = GetResourceFileAsString & Chr$(a_bytFile(lngIdx))
    Next lngIdx

End Function
Private Sub CreateFile(strFilename_in As String, strContents_in As String)

    Dim f As Integer

    f = FreeFile()
    Open strFilename_in For Output As #f
    Print #f, strContents_in
    Close #f

End Sub
Private Sub AddToFavourites()

    '----------------------------------------------------
    'adds the selected game or sid to the favourites view
    '----------------------------------------------------
    
    Dim lngFilterIdx As Long        'loop index
    Dim blnAlreadyFav As Boolean    'flag specifying if the item is already in the favourites view
    Dim lngItemCount As Long        'for getting number of items in the favourites view
        
    Dim lngViewIdx As Long  'the selected view
    
    'only add if not already in favourites view!
    If lstView.ListIndex <> gc_lngFAVOURITESINDEX Then
        
        'gotta have a selected item!
        lngViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If Not (lvMain(lngViewIdx).SelectedItem Is Nothing) Then
            
            'add to favourites view
            If g_eWhatIsSelected = e_SelType_IsGame Then
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Fav FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
            Else
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT SFav AS Fav FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
            End If
            With g_rsGames
                If Not .EOF Then
                    .MoveFirst
                    If !Fav = True Then
                        blnAlreadyFav = True
                    Else
                        .Edit
                        !Fav = True
                        .Update
                        .Edit
                        blnAlreadyFav = False
                    End If
                End If
                .Close
            End With
            Set g_rsGames = Nothing
            
            If blnAlreadyFav = False Then
                'uncache the favourites view
                ga_blnViewsLoaded(gc_lngFAVOURITESINDEX) = False
                
                'remove from current view if a filter specifies "no favourites"
                With ga_udtViewData(lstView.ListIndex)
                    For lngFilterIdx = 1 To ga_udtViewData(lstView.ListIndex).lngFilterCount
                        With .a_udtFilters(lngFilterIdx)
                            If .eClauseType = e_ClauseType_IsDatabaseField And .eOperator = e_Op_Equal And .strFieldName = "Fav" And .strClauseData = "0" Then
                                'no favourites allowed in this view so remove it!
                                Call RemoveItemFromList(lngViewIdx, lstView.ListIndex, lvMain(lngViewIdx).SelectedItem.Index, True)
                                Exit For
                            End If
                        End With
                    Next lngFilterIdx
                End With
                
                'update the view selector and title text
                lngItemCount = 0
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Fav FROM Games WHERE Fav = True", dbOpenSnapshot)
                With g_rsGames
                    If Not .EOF Then
                        .MoveLast
                        .MoveFirst
                        lngItemCount = .RecordCount
                        .Close
                    End If
                End With
                Set g_rsGames = Nothing
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT SFav FROM Music WHERE SFav = True", dbOpenSnapshot)
                With g_rsGames
                    If Not .EOF Then
                        .MoveLast
                        .MoveFirst
                        lngItemCount = lngItemCount + .RecordCount
                        .Close
                    End If
                End With
                Set g_rsGames = Nothing
                With lstView
                    If lngItemCount = 0 Then
                        .List(gc_lngFAVOURITESINDEX) = ga_udtViewData(gc_lngFAVOURITESINDEX).strTitle & "  (" & g_strViewNone & ")"
                    Else
                        .List(gc_lngFAVOURITESINDEX) = ga_udtViewData(gc_lngFAVOURITESINDEX).strTitle & "  (" & lngItemCount & ")"
                    End If
                    lvMain(gc_lngFAVOURITESINDEX).ColumnHeaders(1).Text = .List(gc_lngFAVOURITESINDEX)
                End With

                'make the listitem bold in all loaded views
                lvMain(lngViewIdx).SelectedItem.Bold = True 'incase custom view
                Call SetItemBoldInAllViews(lvMain(lngViewIdx).SelectedItem.Key, True)

            End If

            'show "Added" message
            MsgBox lvMain(lngViewIdx).SelectedItem.Text & " " & g_xobjFunc.GetINIMessage("Messages", "AddedToFavs", "has been added to your Favourites!"), vbInformation, gc_strAPPNAME

        End If
    End If

End Sub
Public Sub AddToRecentList(strKey_in As String, strName_in As String)

    Dim lngMenuIdx As Long              'loop index
    Dim blnIsGame As Boolean            'testing flag
    Dim lngFoundIdx As Long             'position of item if already found in recent list
    Dim udtTempItem As RecentMenuType   'temp swap storage

    'abort if we're not allowed to add
    If ma_udtRecent(0).strKey = strKey_in Then
        'already at top of the recent list!
        Exit Sub
    End If
    blnIsGame = (Left$(strKey_in, 1) = "G")
    If (g_lngRecentStore = gc_lngRECENT_GAMES) And (blnIsGame = False) Then
        Exit Sub
    ElseIf (g_lngRecentStore = gc_lngRECENT_SIDS) And (blnIsGame = True) Then
        Exit Sub
    End If

    'is the item already in the list?
    lngFoundIdx = 0
    For lngMenuIdx = 0 To 9
        If ma_udtRecent(lngMenuIdx).strKey = strKey_in Then
            lngFoundIdx = lngMenuIdx
        End If
    Next lngMenuIdx

    If lngFoundIdx > 0 Then
        'item found in list already - move to top
        udtTempItem = ma_udtRecent(lngFoundIdx)
        For lngMenuIdx = lngFoundIdx To 1 Step -1
            ma_udtRecent(lngMenuIdx) = ma_udtRecent(lngMenuIdx - 1)
        Next lngMenuIdx
        ma_udtRecent(0) = udtTempItem
    Else
        'add item to top of list
        For lngMenuIdx = 9 To 1 Step -1
            ma_udtRecent(lngMenuIdx) = ma_udtRecent(lngMenuIdx - 1)
        Next lngMenuIdx
        With ma_udtRecent(0)
            .strName = Replace(strName_in, "&", "&&")
            .strKey = strKey_in
            .blnIsGame = blnIsGame
        End With
    End If

    'update the menu
    Call UpdateRecentMenu

End Sub

Private Function AllowVerifyFiles() As Boolean

    If ((g_lngTotalGamePaths > 0) And (g_blnValidGamePathExists = True)) Or (g_lngTotalScrnShotPaths > 0) Or (g_udtRegPathSid.ePathType <> e_PathLocation_Invalid) Then
        AllowVerifyFiles = True
    Else
        AllowVerifyFiles = False
    End If
    
End Function

Private Sub CopyCustomView(udtSourceView_in As ViewDataType, udtDestView_out As ViewDataType, blnCopyPrefs_in As Boolean)

    '-----------------------------------------------------------------------------
    'copies one view into another view
    '
    'Inputs:    udtSourceView_in:   source view to copy from
    '           udtDestView_out:    dest view to copy into
    '           blnCopyPrefs_in:    flag specifying whether to copy the view prefs
    '-----------------------------------------------------------------------------
    
    Dim lngFilterIdx As Long    'loop index
    Dim lngIdx As Long
    
    'copy across the view info
    With udtDestView_out
        'view data
        .strTitle = udtSourceView_in.strTitle
        .eFilterMode = udtSourceView_in.eFilterMode
        .eIncludeMusicTable = udtSourceView_in.eIncludeMusicTable
        .lngFilterCount = udtSourceView_in.lngFilterCount
        ReDim .a_udtFilters(.lngFilterCount)
        'view filters
        For lngFilterIdx = 1 To .lngFilterCount
            With .a_udtFilters(lngFilterIdx)
                .eClauseType = udtSourceView_in.a_udtFilters(lngFilterIdx).eClauseType
                .eOperator = udtSourceView_in.a_udtFilters(lngFilterIdx).eOperator
                .strClauseData = udtSourceView_in.a_udtFilters(lngFilterIdx).strClauseData
                .strFieldName = udtSourceView_in.a_udtFilters(lngFilterIdx).strFieldName
                .strFieldTable = udtSourceView_in.a_udtFilters(lngFilterIdx).strFieldTable
                .strMusicFieldName = udtSourceView_in.a_udtFilters(lngFilterIdx).strMusicFieldName
                .strMusicFieldTable = udtSourceView_in.a_udtFilters(lngFilterIdx).strMusicFieldTable
            End With
        Next lngFilterIdx
        'view prefs
        If blnCopyPrefs_in Then
            .lngCol1Width = udtSourceView_in.lngCol1Width
            .lngListViewType = udtSourceView_in.lngListViewType
            .lngSortColumn = udtSourceView_in.lngSortColumn
            .lngSortOrder = udtSourceView_in.lngSortOrder
            .strSelectedItem = udtSourceView_in.strSelectedItem
            .lngExtraColumnsCount = udtSourceView_in.lngExtraColumnsCount
            .lngMaxExtraColumns = udtSourceView_in.lngMaxExtraColumns
            .strExtraColumns = udtSourceView_in.strExtraColumns
            If udtSourceView_in.lngExtraColumnsCount > 0 Then
                ReDim Preserve .a_udtExtraColumns(udtSourceView_in.lngExtraColumnsCount - 1)
                For lngIdx = 0 To udtSourceView_in.lngExtraColumnsCount - 1
                    .a_udtExtraColumns(lngIdx) = udtSourceView_in.a_udtExtraColumns(lngIdx)
                Next lngIdx
            End If
        End If
    End With

End Sub

Private Sub DeleteCustomViewFromSelector(ByVal lngIdxToDelete_in As Long)

    '--------------------------------------------------
    'Deletes the passed in item from the view selector
    '
    'Inputs:    lngIdxToDelete_in:  The index to delete
    '--------------------------------------------------
    
    Dim lngViewIdx As Long      'loop index
        
    With lstView
        'move the views up one from the one after the one to delete
        For lngViewIdx = (lngIdxToDelete_in + 1) To .ListCount - 1
            .List(lngViewIdx - 1) = .List(lngViewIdx)
        Next lngViewIdx
        'remove the last view
        .RemoveItem .ListCount - 1
    End With

End Sub

Private Sub CreateObjects()

    '-----------------------------------------------
    'Creates all the global and module level objects
    '-----------------------------------------------

    'timing
    Set g_xobjClock = New GBFuncs.clsClock

    'SQL Engine
    Set m_xobjSQLEngine = New GBSql.clsQuery

    'GEMUS
    Set g_xobjGEMUS = New Gemus.clsGemus

End Sub

Private Sub FindInList(ByVal lngListViewIdx_in As Long, ByVal lngStartIdx_in As Long)

    '-------------------------------------------------------------------------------
    'finds m_strFindInList in the lngListViewIdx_in listview starting from index lngStartIdx_in
    '-------------------------------------------------------------------------------

    Dim lngCurrentIdx As Long
    Dim lngItemIdx As Long
    Dim blnFound As Boolean
    
    With lvMain(lngListViewIdx_in)
        If .ListItems.Count = 0 Then
            mnuEditFindNext.Enabled = False
            Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            DoEvents
            blnFound = False
            lngCurrentIdx = lngStartIdx_in
            If lngCurrentIdx = .ListItems.Count Then
                For lngItemIdx = 1 To .ListItems.Count
                    If InStr(1, .ListItems(lngItemIdx).Text, m_strFindInList, vbTextCompare) > 0 Then
                        blnFound = True
                        Exit For
                    End If
                    If .ColumnHeaders.Count = 2 Then
                        If InStr(1, .ListItems(lngItemIdx).SubItems(1), m_strFindInList, vbTextCompare) > 0 Then
                            blnFound = True
                            Exit For
                        End If
                    End If
                    If m_lngFindCommentInList = 1 Then
                        If MatchAKA(m_strFindInList, .ListItems(lngItemIdx).Key) = True Then
                            blnFound = True
                            Exit For
                        End If
                    End If
                Next lngItemIdx
            Else
                For lngItemIdx = lngCurrentIdx + 1 To .ListItems.Count
                    If InStr(1, .ListItems(lngItemIdx).Text, m_strFindInList, vbTextCompare) > 0 Then
                        blnFound = True
                        Exit For
                    End If
                    If .ColumnHeaders.Count = 2 Then
                        If InStr(1, .ListItems(lngItemIdx).SubItems(1), m_strFindInList, vbTextCompare) > 0 Then
                            blnFound = True
                            Exit For
                        End If
                    End If
                    If m_lngFindCommentInList = 1 Then
                        If MatchAKA(m_strFindInList, .ListItems(lngItemIdx).Key) = True Then
                            blnFound = True
                            Exit For
                        End If
                    End If
                Next lngItemIdx
                If blnFound = False Then
                    For lngItemIdx = 1 To lngCurrentIdx
                        If InStr(1, .ListItems(lngItemIdx).Text, m_strFindInList, vbTextCompare) > 0 Then
                            blnFound = True
                            Exit For
                        End If
                        If .ColumnHeaders.Count = 2 Then
                            If InStr(1, .ListItems(lngItemIdx).SubItems(1), m_strFindInList, vbTextCompare) > 0 Then
                                blnFound = True
                                Exit For
                            End If
                        End If
                        If m_lngFindCommentInList = 1 Then
                            If MatchAKA(m_strFindInList, .ListItems(lngItemIdx).Key) = True Then
                                blnFound = True
                                Exit For
                            End If
                        End If
                    Next lngItemIdx
                End If
            End If

            If blnFound = False Then
                MsgBox g_xobjFunc.GetINIMessage("Find Window", "NoMatches", "No matches found!"), vbInformation, gc_strAPPNAME
                mnuEditFindNext.Enabled = False
            Else
                Call SetItemInList(.ListItems(lngItemIdx).Key, False)
                mnuEditFindNext.Enabled = True
            End If
            Screen.MousePointer = vbDefault
            DoEvents
        End If
    End With

End Sub
Private Sub FindByID(ByVal blnGame_in As Boolean)

    Dim LItem As ListItem
    Dim strKey As String

    If blnGame_in Then
        strKey = "G" & Trim$(m_lngFindIDGame)
    Else
        strKey = "M" & Trim$(m_lngFindIDMusic)
    End If

    Set LItem = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If LItem Is Nothing Then
        Call SetItemInList(strKey, True, IIf((Left$(strKey, 1) = "G"), 0, 1))
    Else
        Call SetItemInList(strKey, True, IIf((Left$(strKey, 1) = "G"), 0, 1), LItem.Key)
    End If
    Set LItem = Nothing

End Sub
Private Sub GetCustomView(ByVal blnNewView_in As Boolean)
    
    '------------------------------------------------------
    'Allows the user to create a new custom view or
    'edit an existing custom view.  The Default views
    'cannot be edited or removed.  The QuickFilter view
    'cannot be removed.
    '
    'Inputs:    blnNewView_in:      Flag specifying whether
    '                               to create a new view
    '                               or edit the currently
    '                               selected view
    '------------------------------------------------------

    Dim lngViewIdx As Long      'the index of the currently selected view
    Dim lngFilterIdx As Long    'loop index
    Dim lngViewMax As Long      'storage for UBound() of views array
    Dim blnSavePrefs As Boolean 'flag to say whether saving view prefs
    Dim lngIdx As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Make sure we CAN edit the selected view if the user selected to edit
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'only if editing a view
    If Not blnNewView_in Then
        If lstView.ListIndex <= gc_lngDEFAULTVIEWSMAX Then
            'disallow editing of Default views
            Call MsgBox(g_xobjFunc.GetINIMessage("Error Messages", "Error_CannotEditView", "You cannot edit a Default View!"), vbInformation, gc_strAPPNAME)
            Exit Sub
        ElseIf (lstView.ListIndex = (lstView.ListCount - 1)) And (m_blnLastViewIsQuick) Then
            'Trying to edit a Quick Filter, so redirect to GetQuickFilter()
            Call GetQuickFilter
            Exit Sub
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'set the Custom View data structure accordingly
    '''''''''''''''''''''''''''''''''''''''''''''''

    'store the current view index
    lngViewIdx = lstView.ListIndex
        
    If blnNewView_in Then
        'get the window title
        g_strCustomViewWindowTitle = g_xobjFunc.GetINIMessage("CustomViewEditor", "WindowTitleNew", "New Custom View")
        'initialise the view structure
        With g_udtCustomView
            .strTitle = g_xobjFunc.GetINIMessage("CustomViewEditor", "NewViewTitle", "New View")
            .eIncludeMusicTable = e_Table_GamesOnly
            .eFilterMode = e_FilterMode_AND
            .lngFilterCount = 0
            ReDim .a_udtFilters(0)
        End With
    Else
        'get the window title
        g_strCustomViewWindowTitle = g_xobjFunc.GetINIMessage("CustomViewEditor", "WindowTitleEdit", "Edit Custom View")
        'copy the info across from the currently selected view
        With ga_udtViewData(lngViewIdx)
            g_udtCustomView.strTitle = .strTitle
            g_udtCustomView.eIncludeMusicTable = .eIncludeMusicTable
            g_udtCustomView.eFilterMode = .eFilterMode
            g_udtCustomView.lngFilterCount = .lngFilterCount
            ReDim g_udtCustomView.a_udtFilters(.lngFilterCount)
            For lngFilterIdx = 0 To .lngFilterCount
                With .a_udtFilters(lngFilterIdx)
                    g_udtCustomView.a_udtFilters(lngFilterIdx).eClauseType = .eClauseType
                    g_udtCustomView.a_udtFilters(lngFilterIdx).eOperator = .eOperator
                    g_udtCustomView.a_udtFilters(lngFilterIdx).strClauseData = .strClauseData
                    g_udtCustomView.a_udtFilters(lngFilterIdx).strFieldName = .strFieldName
                    g_udtCustomView.a_udtFilters(lngFilterIdx).strFieldTable = .strFieldTable
                    g_udtCustomView.a_udtFilters(lngFilterIdx).strMusicFieldName = .strMusicFieldName
                    g_udtCustomView.a_udtFilters(lngFilterIdx).strMusicFieldTable = .strMusicFieldTable
                End With
            Next lngFilterIdx
        End With
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    'Show the editor and wait for user to close it
    ''''''''''''''''''''''''''''''''''''''''''''''

    frmCV.Show vbModal, Me
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Save the custom view if the user clicked ok
    ''''''''''''''''''''''''''''''''''''''''''''
    
    If g_blnSaveCustomView Then
        
        If blnNewView_in Then
        
            '''''''''''''
            'ADD NEW VIEW
            '''''''''''''
            
            'copy view prefs from the current view
            With ga_udtViewData(lngViewIdx)
                g_udtCustomView.lngCol1Width = .lngCol1Width
                g_udtCustomView.lngListViewType = .lngListViewType
                g_udtCustomView.lngSortColumn = .lngSortColumn
                g_udtCustomView.lngSortOrder = .lngSortOrder
                g_udtCustomView.strSelectedItem = .strSelectedItem
                g_udtCustomView.lngExtraColumnsCount = .lngExtraColumnsCount
                g_udtCustomView.lngMaxExtraColumns = .lngMaxExtraColumns
                g_udtCustomView.strExtraColumns = .strExtraColumns
                If .lngExtraColumnsCount > 0 Then
                    ReDim Preserve g_udtCustomView.a_udtExtraColumns(.lngExtraColumnsCount - 1)
                    For lngIdx = 0 To .lngExtraColumnsCount - 1
                        g_udtCustomView.a_udtExtraColumns(lngIdx) = .a_udtExtraColumns(lngIdx)
                    Next lngIdx
                End If
            End With
            
            'add space in the global views array
            lngViewMax = UBound(ga_udtViewData) + 1
            ReDim Preserve ga_udtViewData(lngViewMax)
            
            'find index to add view into
            If m_blnLastViewIsQuick Then
                'the current last view is a quick filter view, so move it to the new last space
                Call CopyCustomView(ga_udtViewData(lngViewMax - 1), ga_udtViewData(lngViewMax), True)
                lngViewIdx = lngViewMax - 1
            Else
                lngViewIdx = lngViewMax
            End If
        
            'insert the view in the View Selector
            lstView.AddItem g_udtCustomView.strTitle & "  (?)", lngViewIdx
        
            'we need to save the prefs
            blnSavePrefs = True

        Else
        
            ''''''''''''''''''''''''''
            'SAVE CURRENT CHANGED VIEW
            ''''''''''''''''''''''''''
            
            'update the title in the View Selector
            lstView.List(lstView.ListIndex) = g_udtCustomView.strTitle & "  (?)"
            
            'we don't need to save the prefs
            blnSavePrefs = False
            
        End If
        
        'copy into the view's array and save to database
        Call CopyCustomView(g_udtCustomView, ga_udtViewData(lngViewIdx), blnSavePrefs)
        Call SaveCustomView(ga_udtViewData(lngViewIdx))
        
        'refresh the view
        m_blnSaveViewPrefs = blnSavePrefs
        With lstView
            .ListIndex = -1
            .ListIndex = lngViewIdx
        End With
        m_blnSaveViewPrefs = True
         
    End If
    
    
    
End Sub




Private Sub GetQuickFilter()

    'get the filter for the quick-filter view
    frmQuickFilter.Show vbModal, Me
    
    
    
    'test if the user clicked ok or cancel
    If g_blnEasterEgg = True Then
        frmEggs.Show vbModal, Me
    ElseIf g_udtQuickFilter.strClauseData(0) <> "" Then
        Call SetQuickFilter(False)
    End If

End Sub

Public Sub LoadDetailsTabPositions()

    Dim lngIdx As Long
    Dim lngPosIdx As Long
    Dim a_blnVisible() As Boolean
    Dim blnIsAllHidden As Boolean
    
    'read field posiions from the ini file
    ReDim ga_lngDetailsTabPositions(g_lngTotalDetailsTabPositions - 1)
    ReDim a_blnVisible(g_lngTotalDetailsTabPositions - 1)
    For lngIdx = 0 To g_lngTotalDetailsTabPositions - 1
        ga_lngDetailsTabPositions(lngIdx) = g_xobjFunc.INI_GetSetting( _
                "UI", "DetailsTabField" & lngIdx & "Pos", Trim$(lngIdx), g_strSettingsFile)
        a_blnVisible(lngIdx) = False
    Next lngIdx

    'validate tab 0 (game info)
    For lngIdx = 0 To 5
        If ga_lngDetailsTabPositions(lngIdx) < 0 Or ga_lngDetailsTabPositions(lngIdx) > 5 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx
    For lngIdx = 6 To 11
        If ga_lngDetailsTabPositions(lngIdx) < 6 Or ga_lngDetailsTabPositions(lngIdx) > 11 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx

    'validate tab 1 (version info)
    For lngIdx = 12 To 17
        If ga_lngDetailsTabPositions(lngIdx) < 12 Or ga_lngDetailsTabPositions(lngIdx) > 17 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx
    For lngIdx = 18 To 23
        If ga_lngDetailsTabPositions(lngIdx) < 18 Or ga_lngDetailsTabPositions(lngIdx) > 23 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx

    'validate tab 2 (personal info)
    For lngIdx = 24 To 28
        If ga_lngDetailsTabPositions(lngIdx) < 24 Or ga_lngDetailsTabPositions(lngIdx) > 28 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx

    'validate tab 4 (music)
    For lngIdx = 29 To 32
        If ga_lngDetailsTabPositions(lngIdx) < 29 Or ga_lngDetailsTabPositions(lngIdx) > 32 Then
            ga_lngDetailsTabPositions(lngIdx) = -1
        End If
    Next lngIdx

    'set the label positions on the details tabs
    For lngIdx = 0 To g_lngTotalDetailsTabPositions - 1
        lngPosIdx = ga_lngDetailsTabPositions(lngIdx)
        If lngPosIdx > -1 Then
            a_blnVisible(lngPosIdx) = True
            With ga_udtDetailsTabPositionsLookup(lngIdx)
                lblPrompt(lngPosIdx).Visible = True
                lblPrompt(lngPosIdx).Left = .sngLabelTwipsX
                lblPrompt(lngPosIdx).Top = .sngTwipsY
                If lngPosIdx > 28 Then
                    'music tab
                    'lblPrompt(lngPosIdx + 10).Visible = True
                    'lblPrompt(lngPosIdx + 10).Left = .sngLabelTwipsX + 30  'hack due to no picturebox
                    'lblPrompt(lngPosIdx + 10).Top = .sngTwipsY + 720 'hack due to no picturebox
                End If
                Select Case lngPosIdx
                    Case 0
                        lbl_Info_Year.Visible = True
                        lbl_Info_Publisher.Visible = True
                        lbl_Info_Year.Left = .sngFieldTwipsX
                        lbl_Info_Year.Top = .sngTwipsY
                        lbl_Info_Publisher.Top = .sngTwipsY
                        lbl_Info_Publisher.Left = lbl_Info_Year.Left + lbl_Info_Year.Width + g_sngCopyrightGap
                    Case 1
                        lbl_Info_Developer.Visible = True
                        lbl_Info_Developer.Left = .sngFieldTwipsX
                        lbl_Info_Developer.Top = .sngTwipsY
                    Case 2
                        lbl_Info_Programmer.Visible = True
                        lbl_Info_Programmer.Left = .sngFieldTwipsX
                        lbl_Info_Programmer.Top = .sngTwipsY
                    Case 3
                        lbl_Info_Musician.Visible = True
                        lbl_Info_Musician.Left = .sngFieldTwipsX
                        lbl_Info_Musician.Top = .sngTwipsY
                    Case 4
                        lbl_Info_Artist.Visible = True
                        lbl_Info_Artist.Left = .sngFieldTwipsX
                        lbl_Info_Artist.Top = .sngTwipsY
                    Case 5
                        lbl_Info_License.Visible = True
                        lbl_Info_License.Left = .sngFieldTwipsX
                        lbl_Info_License.Top = .sngTwipsY
                    Case 6
                        lbl_Info_Genre.Visible = True
                        lbl_Info_Genre.Left = .sngFieldTwipsX
                        lbl_Info_Genre.Top = .sngTwipsY
                    Case 7
                        lbl_Info_Players.Visible = True
                        lbl_Info_Players.Left = .sngFieldTwipsX
                        lbl_Info_Players.Top = .sngTwipsY
                    Case 8
                        lbl_Info_Language.Visible = True
                        lbl_Info_Language.Left = .sngFieldTwipsX
                        lbl_Info_Language.Top = .sngTwipsY
                    Case 9
                        lbl_Info_Control.Visible = True
                        lbl_Info_Control.Left = .sngFieldTwipsX
                        lbl_Info_Control.Top = .sngTwipsY
                    Case 10
                        lbl_Info_Rarity.Visible = True
                        lbl_Info_Rarity.Left = .sngFieldTwipsX
                        lbl_Info_Rarity.Top = .sngTwipsY
                    Case 11
                        lbl_Info_WebLink.Visible = True
                        lbl_Info_WebLink.Left = .sngFieldTwipsX
                        lbl_Info_WebLink.Top = .sngTwipsY
                    Case 12
                        lbl_Info_Cracker.Visible = True
                        lbl_Info_Cracker.Left = .sngFieldTwipsX
                        lbl_Info_Cracker.Top = .sngTwipsY
                    Case 13
                        lbl_Info_Trainers.Visible = True
                        lbl_Info_Trainers.Left = .sngFieldTwipsX
                        lbl_Info_Trainers.Top = .sngTwipsY
                    Case 14
                        lbl_Info_Playable.Visible = True
                        lbl_Info_Playable.Left = .sngFieldTwipsX
                        lbl_Info_Playable.Top = .sngTwipsY
                    Case 15
                        lbl_Info_GameLength.Visible = True
                        lbl_Info_GameLength.Left = .sngFieldTwipsX
                        lbl_Info_GameLength.Top = .sngTwipsY
                    Case 16
                        lbl_Info_NTSC.Visible = True
                        lbl_Info_NTSC.Left = .sngFieldTwipsX
                        lbl_Info_NTSC.Top = .sngTwipsY
                    Case 17
                        lbl_Info_VWebLink.Visible = True
                        lbl_Info_VWebLink.Left = .sngFieldTwipsX
                        lbl_Info_VWebLink.Top = .sngTwipsY
                    Case 18
                        lbl_Info_Original.Visible = True
                        lbl_Info_Original.Left = .sngFieldTwipsX
                        lbl_Info_Original.Top = .sngTwipsY
                    Case 19
                        lbl_Info_LoadingScreen.Visible = True
                        lbl_Info_LoadingScreen.Left = .sngFieldTwipsX
                        lbl_Info_LoadingScreen.Top = .sngTwipsY
                    Case 20
                        lbl_Info_TitleScreen.Visible = True
                        lbl_Info_TitleScreen.Left = .sngFieldTwipsX
                        lbl_Info_TitleScreen.Top = .sngTwipsY
                    Case 21
                        lbl_Info_Docs.Visible = True
                        lbl_Info_Docs.Left = .sngFieldTwipsX
                        lbl_Info_Docs.Top = .sngTwipsY
                    Case 22
                        lbl_Info_HSS.Visible = True
                        lbl_Info_HSS.Left = .sngFieldTwipsX
                        lbl_Info_HSS.Top = .sngTwipsY
                    Case 23
                        lbl_Info_TDE.Visible = True
                        lbl_Info_TDE.Left = .sngFieldTwipsX
                        lbl_Info_TDE.Top = .sngTwipsY
                    Case 24
                        lbl_Info_HighScore.Visible = True
                        lbl_Info_HighScore.Left = .sngFieldTwipsX
                        lbl_Info_HighScore.Top = .sngTwipsY
                    Case 25
                        lbl_Info_Difficulty.Visible = True
                        lbl_Info_Difficulty.Left = .sngFieldTwipsX
                        lbl_Info_Difficulty.Top = .sngTwipsY
                    Case 26
                        lbl_Info_TP.Visible = True
                        lbl_Info_TP.Left = .sngFieldTwipsX
                        lbl_Info_TP.Top = .sngTwipsY
                    Case 27
                        lbl_Info_LP.Visible = True
                        lbl_Info_LP.Left = .sngFieldTwipsX
                        lbl_Info_LP.Top = .sngTwipsY
                    Case 28
                        lbl_Info_Rating.Visible = True
                        lbl_Info_Rating.Left = .sngFieldTwipsX
                        lbl_Info_Rating.Top = .sngTwipsY
                    Case 29
                        lbl_SFInfo_Musician.Visible = True
                        'lbl_SInfo_Musician.Visible = True
                        lbl_SFInfo_Musician.Left = .sngFieldTwipsX
                        lbl_SFInfo_Musician.Top = .sngTwipsY
                        'lbl_SInfo_Musician.Left = .sngFieldTwipsX + 30  'hack due to no picturebox
                        'lbl_SInfo_Musician.Top = .sngTwipsY + 720 'hack due to no picturebox
                    Case 30
                        lbl_SFInfo_Nick.Visible = True
                        'lbl_SInfo_Nick.Visible = True
                        lbl_SFInfo_Nick.Left = .sngFieldTwipsX
                        lbl_SFInfo_Nick.Top = .sngTwipsY
                        'lbl_SInfo_Nick.Left = .sngFieldTwipsX + 30  'hack due to no picturebox
                        'lbl_SInfo_Nick.Top = .sngTwipsY + 720 'hack due to no picturebox
                    Case 31
                        lbl_SFInfo_Group.Visible = True
                        'lbl_SInfo_Group.Visible = True
                        lbl_SFInfo_Group.Left = .sngFieldTwipsX
                        lbl_SFInfo_Group.Top = .sngTwipsY
                        'lbl_SInfo_Group.Left = .sngFieldTwipsX + 30  'hack due to no picturebox
                        'lbl_SInfo_Group.Top = .sngTwipsY + 720 'hack due to no picturebox
                    Case 32
                        lbl_SFInfo_Photo.Visible = True
                        'lbl_SInfo_Photo.Visible = True
                        lbl_SFInfo_Photo.Left = .sngFieldTwipsX
                        lbl_SFInfo_Photo.Top = .sngTwipsY
                        'lbl_SInfo_Photo.Left = .sngFieldTwipsX + 30  'hack due to no picturebox
                        'lbl_SInfo_Photo.Top = .sngTwipsY + 720 'hack due to no picturebox
                End Select
            End With
        End If
    Next lngIdx

    'hide the non-showing fields
    For lngPosIdx = 0 To g_lngTotalDetailsTabPositions - 1
        If a_blnVisible(lngPosIdx) = False Then
            lblPrompt(lngPosIdx).Visible = False
            If lngPosIdx > 28 Then
                'music tab
                'lblPrompt(lngPosIdx + 10).Visible = False
            End If
            Select Case lngPosIdx
                Case 0
                    lbl_Info_Year.Visible = False
                    lbl_Info_Publisher.Visible = False
                Case 1
                    lbl_Info_Developer.Visible = False
                Case 2
                    lbl_Info_Programmer.Visible = False
                Case 3
                    lbl_Info_Musician.Visible = False
                Case 4
                    lbl_Info_Artist.Visible = False
                Case 5
                    lbl_Info_License.Visible = False
                Case 6
                    lbl_Info_Genre.Visible = False
                Case 7
                    lbl_Info_Players.Visible = False
                Case 8
                    lbl_Info_Language.Visible = False
                Case 9
                    lbl_Info_Control.Visible = False
                Case 10
                    lbl_Info_Rarity.Visible = False
                Case 11
                    lbl_Info_WebLink.Visible = False
                Case 12
                    lbl_Info_Cracker.Visible = False
                Case 13
                    lbl_Info_Trainers.Visible = False
                Case 14
                    lbl_Info_Playable.Visible = False
                Case 15
                    lbl_Info_GameLength.Visible = False
                Case 16
                    lbl_Info_NTSC.Visible = False
                Case 17
                    lbl_Info_VWebLink.Visible = False
                Case 18
                    lbl_Info_Original.Visible = False
                Case 19
                    lbl_Info_LoadingScreen.Visible = False
                Case 20
                    lbl_Info_TitleScreen.Visible = False
                Case 21
                    lbl_Info_Docs.Visible = False
                Case 22
                    lbl_Info_HSS.Visible = False
                Case 23
                    lbl_Info_TDE.Visible = False
                Case 24
                    lbl_Info_HighScore.Visible = False
                Case 25
                    lbl_Info_Difficulty.Visible = False
                Case 26
                    lbl_Info_TP.Visible = False
                Case 27
                    lbl_Info_LP.Visible = False
                Case 28
                    lbl_Info_Rating.Visible = False
                Case 29
                    lbl_SFInfo_Musician.Visible = False
                    'lbl_SInfo_Musician.Visible = False
                Case 30
                    lbl_SFInfo_Nick.Visible = False
                    'lbl_SInfo_Nick.Visible = False
                Case 31
                    lbl_SFInfo_Group.Visible = False
                    'lbl_SInfo_Group.Visible = False
                Case 32
                    lbl_SFInfo_Photo.Visible = False
                    'lbl_SInfo_Photo.Visible = False
            End Select
        End If
    Next lngPosIdx

    'hide tabs with no fields (does not hide the music tab for the music table entries)
    blnIsAllHidden = True
    For lngIdx = 0 To 11
        If lblPrompt(lngIdx).Visible = True Then
            blnIsAllHidden = False
            Exit For
        End If
    Next lngIdx
    tabDetails.TabVisible(0) = Not blnIsAllHidden

    blnIsAllHidden = True
    For lngIdx = 12 To 23
        If lblPrompt(lngIdx).Visible = True Then
            blnIsAllHidden = False
            Exit For
        End If
    Next lngIdx
    tabDetails.TabVisible(1) = Not blnIsAllHidden

    blnIsAllHidden = True
    For lngIdx = 24 To 28
        If lblPrompt(lngIdx).Visible = True Then
            blnIsAllHidden = False
            Exit For
        End If
    Next lngIdx
    tabDetails.TabVisible(2) = Not blnIsAllHidden

    blnIsAllHidden = True
    For lngIdx = 29 To 32
        If lblPrompt(lngIdx).Visible = True Then
            blnIsAllHidden = False
            Exit For
        End If
    Next lngIdx
    tabDetails.TabVisible(4) = Not blnIsAllHidden

End Sub

Private Function LoadImageExtrasIntoArray(a_lngExtraIDs_inout() As Long, a_strImageExtras_inout() As String, lngCurrentImageExtrasUbound_inout As Long, lngGameId_in As Long)

    On Error GoTo Err_Handler

    'grab the extra's details
    Dim lngExtrasPathIdx As Long
    Dim strPath As String
    Dim strPathFile As String
    
    LoadImageExtrasIntoArray = 0
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT EX_Id, Name, Path FROM Extras WHERE EA=True AND Type=0 AND GA_Id = " & lngGameId_in & " AND (Path LIKE ""*.png"" OR Path LIKE ""*.gif"" OR Path LIKE ""*.jpg"" OR Path LIKE ""*.bmp"") ORDER BY DisplayOrder", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                strPath = !Path
                If FindExtrasFile(lngExtrasPathIdx, strPath) = "" Then
                    lngCurrentImageExtrasUbound_inout = lngCurrentImageExtrasUbound_inout + 1
                    ReDim Preserve a_strImageExtras_inout(lngCurrentImageExtrasUbound_inout)
                    ReDim Preserve a_lngExtraIDs_inout(lngCurrentImageExtrasUbound_inout)
                    a_strImageExtras_inout(lngCurrentImageExtrasUbound_inout) = ga_udtRegPathExtras(lngExtrasPathIdx).strWithBackSlash & strPath
                    a_lngExtraIDs_inout(lngCurrentImageExtrasUbound_inout) = !EX_Id
                    LoadImageExtrasIntoArray = LoadImageExtrasIntoArray + 1
                End If
                .MoveNext
            Loop
            .Close
        End If
    End With
    Set g_rsGames = Nothing

    Exit Function

Err_Handler:

End Function

Private Function LoadPictureFromFile(strFilename_in As String, picDest_in As PictureBox) As Boolean

    On Error GoTo BadFile

    Dim udtSnug As SnugFitType
    Dim strType As String

    strType = UCase$(Right$(strFilename_in, 3))

    Select Case strType
        Case "PNG"
            picFull.Cls
            picFull.Picture = LoadPicture
            picFull.AutoSize = False
            png_load strFilename_in
            picFull.Width = png_width * Screen.TwipsPerPixelX
            picFull.Height = png_height * Screen.TwipsPerPixelY
            png_draw picFull.hDC, 0, 0
            picFull.Picture = picFull.Image
        Case Else
            picFull.AutoSize = True
            picFull.Picture = LoadPicture(strFilename_in)
    End Select
    
    'calculate the new size and blit to picPreview
    With udtSnug
        .lngDestH = picDest_in.ScaleHeight
        .lngDestW = picDest_in.ScaleWidth
        .lngSourceH = picFull.ScaleHeight
        .lngSourceW = picFull.ScaleWidth
        Call GetSnugFitImageDimensions(udtSnug)
        picDest_in.Picture = LoadPicture
        picDest_in.PaintPicture picFull.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picFull.ScaleWidth, picFull.ScaleHeight
    End With
    LoadPictureFromFile = True
    
    Exit Function

BadFile:
    
    LoadPictureFromFile = False
    If m_lngRegScrnShots = mc_SCRNSHOTSBIG Then
        With udtSnug
            .lngDestH = picDest_in.ScaleHeight
            .lngDestW = picDest_in.ScaleWidth
            .lngSourceH = picNoSS.ScaleHeight
            .lngSourceW = picNoSS.ScaleWidth
            Call GetSnugFitImageDimensions(udtSnug)
            picDest_in.Picture = LoadPicture
            picDest_in.PaintPicture picNoSS.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picNoSS.ScaleWidth, picNoSS.ScaleHeight
        End With
    Else
        With udtSnug
            .lngDestH = picDest_in.ScaleHeight
            .lngDestW = picDest_in.ScaleWidth
            .lngSourceH = picNoSSSmall.ScaleHeight
            .lngSourceW = picNoSSSmall.ScaleWidth
            Call GetSnugFitImageDimensions(udtSnug)
            picDest_in.Picture = LoadPicture
            picDest_in.PaintPicture picNoSSSmall.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picNoSSSmall.ScaleWidth, picNoSSSmall.ScaleHeight
        End With
    End If

End Function

Public Sub LoadSidPlayerPopupMenu()

    '-------------------------------------------------
    'loads the popup menu for the sid player selection
    '-------------------------------------------------

    Dim lngMenuIdx As Long      'loop index

    'delete all current menus (if any)
    For lngMenuIdx = 1 To pmnuSidPlayer.Count - 1
        Unload pmnuSidPlayer(lngMenuIdx)
    Next lngMenuIdx

    If g_lngTotalSidPlayers > 0 Then
        'load any new ones
        For lngMenuIdx = 1 To (g_lngTotalSidPlayers - 1)
            Load pmnuSidPlayer(lngMenuIdx)
        Next lngMenuIdx
        'set their properties
        For lngMenuIdx = 0 To (g_lngTotalSidPlayers - 1)
            With pmnuSidPlayer(lngMenuIdx)
                .Caption = ga_udtSidPlayers(lngMenuIdx).strName
                .Enabled = True
                .Visible = True
            End With
        Next lngMenuIdx
    Else
        'no players available
        pmnuSidPlayer(0).Enabled = False
        pmnuSidPlayer(0).Caption = cmdSid.Caption
    End If

End Sub
Public Sub LoadEmulatorPopupMenu()

    '------------------------------------------------
    'loads the popup menu for the emulators selection
    '------------------------------------------------

    Dim lngMenuIdx As Long      'loop index

    'delete all current menus (if any)
    For lngMenuIdx = 1 To pmnuEmulator.Count - 1
        Unload pmnuEmulator(lngMenuIdx)
    Next lngMenuIdx

    If g_xobjGEMUS.TotalEmulators > 0 Then
        'load any new ones
        For lngMenuIdx = 1 To (g_xobjGEMUS.TotalEmulators - 1)
            Load pmnuEmulator(lngMenuIdx)
        Next lngMenuIdx
        'set their properties
        For lngMenuIdx = 0 To (g_xobjGEMUS.TotalEmulators - 1)
            With pmnuEmulator(lngMenuIdx)
                .Caption = g_xobjGEMUS.EmulatorName(lngMenuIdx)
                .Enabled = True
                .Visible = True
            End With
        Next lngMenuIdx
    Else
        'no emulators available
        pmnuEmulator(0).Enabled = False
        pmnuEmulator(0).Caption = cmdPlay.Caption
    End If

End Sub
Public Sub LoadGemusMenu()

    '------------------------------------------------
    'loads the Gemus menu for the emulators selection
    '------------------------------------------------

    Dim lngMenuIdx As Long      'loop index

    'delete all current menus (if any)
    For lngMenuIdx = 9 To mnuGemusEmu.Count - 1
        Unload mnuGemusEmu(lngMenuIdx)
    Next lngMenuIdx

    If g_xobjGEMUS.TotalEmulators > 0 Then
        mnuGemusBar3.Visible = True
        'load any new ones
        For lngMenuIdx = 9 To (g_xobjGEMUS.TotalEmulators - 1)
            Load mnuGemusEmu(lngMenuIdx)
        Next lngMenuIdx
        'set their properties
        For lngMenuIdx = 0 To (g_xobjGEMUS.TotalEmulators - 1)
            With mnuGemusEmu(lngMenuIdx)
                .Caption = (lngMenuIdx + 1) & ".  " & g_xobjGEMUS.EmulatorName(lngMenuIdx) & "..."
                .Enabled = True
                .Visible = True
            End With
        Next lngMenuIdx
        'hide and disable surplus menus
        For lngMenuIdx = g_xobjGEMUS.TotalEmulators To mnuGemusEmu.Count - 1
            With mnuGemusEmu(lngMenuIdx)
                .Caption = ""
                .Enabled = False
                .Visible = False
            End With
        Next lngMenuIdx
    Else
        'no emulators available
        For lngMenuIdx = 0 To (mnuGemusEmu.Count - 1)
            With mnuGemusEmu(lngMenuIdx)
                .Enabled = False
                .Visible = False
                .Caption = ""
            End With
        Next lngMenuIdx
        mnuGemusBar3.Visible = False
    End If

    mnuGemusExport.Enabled = (g_xobjGEMUS.TotalEmulators > 0)

End Sub

Private Function MatchAKA(strFindText_in As String, ByVal strItemKey_in As String) As Boolean

    'returns true if the game has matching text in the game comment field (matching to the find text string)
    'else returns false

    Dim blnMatch As Boolean

    If Left$(strItemKey_in, 1) = "G" Then
        blnMatch = False
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Comment FROM Games WHERE GA_Id = " & Right$(strItemKey_in, Len(strItemKey_in) - 1), dbOpenSnapshot)
        With g_rsGames
            If Not .EOF Then
                .MoveFirst
                blnMatch = (InStr(1, !Comment, strFindText_in, vbTextCompare) > 0)
                .Close
            End If
        End With
        Set g_rsGames = Nothing
        MatchAKA = blnMatch
    Else
        MatchAKA = False
    End If

End Function


Private Sub ReadInRecentMenuList()

    Dim lngMenuIdx As Long  'loop index
    Dim lngClearIdx As Long 'loop index

    'read in the items
    For lngMenuIdx = 0 To 9
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngMenuIdx & "Name", ma_udtRecent(lngMenuIdx).strName, True, ""
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngMenuIdx & "Key", ma_udtRecent(lngMenuIdx).strKey, True, ""
        If ma_udtRecent(lngMenuIdx).strKey = "" Then
            ma_udtRecent(lngMenuIdx).strName = ""
            ma_udtRecent(lngMenuIdx).blnIsGame = False
            Exit For
        Else
            ma_udtRecent(lngMenuIdx).blnIsGame = (Left$(ma_udtRecent(lngMenuIdx).strKey, 1) = "G")
            ma_udtRecent(lngMenuIdx).strName = Replace(ma_udtRecent(lngMenuIdx).strName, " & ", " && ")
        End If
    Next lngMenuIdx

    'clear surplus items
    For lngClearIdx = (lngMenuIdx + 1) To 9
        With ma_udtRecent(lngClearIdx)
            .strName = ""
            .strKey = ""
            .blnIsGame = False
        End With
    Next lngClearIdx

    'update the menu
    Call UpdateRecentMenu

End Sub
Private Sub RemoveExtras(ByVal lngGameId_in As Long, blnClearFlag_in As Boolean, blnRemoveExtras_in As Boolean)

    'cleans up the extras for a game

    On Error GoTo Err_Handler

    If blnClearFlag_in Then
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Extras FROM Games WHERE GA_Id = " & lngGameId_in, dbOpenDynaset)
        With g_rsGames
            If Not .EOF Then
                .MoveFirst
                .Edit
                !Extras = False
                .Update
            End If
            .Close
        End With
        Set g_rsGames = Nothing
    End If

    If blnRemoveExtras_in Then
        g_dbGamebase.Execute "DELETE * FROM Extras WHERE GA_Id = " & lngGameId_in
    End If

    Exit Sub

Err_Handler:

End Sub

Private Sub RemoveRecentItem(ByVal lngIdxToRemove As Long)

    Dim lngIdx As Long

    For lngIdx = lngIdxToRemove To 8
        ma_udtRecent(lngIdx) = ma_udtRecent(lngIdx + 1)
    Next lngIdx
    ma_udtRecent(9).blnIsGame = False
    ma_udtRecent(9).strKey = ""
    ma_udtRecent(9).strName = ""

End Sub

Private Sub RepositionDTIconsX()

    m_sngTabIconsX = ScaleWidth - (120 + imgAdult.Width)

    imgAdult.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgAdult.Width)

    imgRatingFace.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgRatingFace.Width)

    imgRelated.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgRelated.Width)

    imgCloneOf.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgCloneOf.Width)

    imgSequel.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgSequel.Width)

    imgPrequel.Left = m_sngTabIconsX
    m_sngTabIconsX = m_sngTabIconsX - ((m_sngTabIconsGap * 10) + imgRRStore(51).Width)

    imgRRStore(51).Left = m_sngTabIconsX

End Sub

Public Sub RunExtra(ByVal lngExtraID_in As Long)

    Dim lngGameID As Long
    Dim lngType As Long
    Dim strName As String
    Dim strPath As String
    Dim strData As String
    Dim strError As String
    Dim lngExtrasPathIdx As Long
    Dim strExtrasFile As String
    Dim strExt As String
    Dim lngGemusType As Long
    Dim lngPlayersFrom As Long
    Dim lngPlayersTo As Long
    Dim strRunnableFileBefore As String
    Dim strRunnableFile As String
    Dim strOldSBText As String

    On Error GoTo Err_Handler

    'grab the extra's details
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Extras WHERE EX_Id = " & lngExtraID_in, dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            'cache details
            lngGameID = !GA_Id
            lngType = !Type
            strName = !Name
            strPath = !Path
            strData = !Data
            strRunnableFile = !FileToRun
            strRunnableFileBefore = !FileToRun
            .Close
        Else
            'no extra!
            .Close
            Set g_rsGames = Nothing
            Call RemoveExtras(g_udtGame.lngID, True, False)
            MsgBox "The Extra was not found in the database!", vbInformation, gc_strAPPNAME
            Exit Sub
        End If
    End With
    Set g_rsGames = Nothing

    'check it's for the current game
    If lngGameID <> g_udtGame.lngID Then
        MsgBox "Game ID mismatch in database!", vbInformation, gc_strAPPNAME
        Exit Sub
    End If

    'find the full extra's path and filename
    If lngType < 3 Then
        strError = FindExtrasFile(lngExtrasPathIdx, strPath)
        If strError <> "" Then
            MsgBox strError & vbCrLf & vbCrLf & strPath, vbInformation, gc_strAPPNAME
            Exit Sub
        End If
        strExtrasFile = ga_udtRegPathExtras(lngExtrasPathIdx).strWithBackSlash & strPath
    Else
        'url
        strExtrasFile = strPath
    End If

    '---------------------------------------------------------------
    'now we have the full path and file of the extra we want to run!
    '---------------------------------------------------------------
    
    'check the type of extra
    If lngType = 0 Then

        '======
        'NORMAL
        '======

        strExt = UCase$(Right$(strExtrasFile, 4))
        If strExt = ".EXE" Or strExt = ".COM" Or strExt = ".BAT" Then
            Call Shell(strExtrasFile)
        Else
            Call g_xobjFunc.ShellExecuteFile(strExtrasFile, "", False)
        End If

    ElseIf lngType = 1 Then

        '=========
        'USE GEMUS
        '=========

        'get details for running the game
        lngGemusType = GetExtrasPropsFromDataString(strData)
        With g_udtExtrasProps
            If lngGemusType = 1 Then
                lngPlayersFrom = g_udtGame.lngPlayersFrom
                lngPlayersTo = g_udtGame.lngPlayersTo
            Else
                lngPlayersFrom = .lngGF_Players
                lngPlayersTo = .lngGF_Players
            End If

            'run the game!
            Call g_xobjGEMUS.SetRezipNotify(g_blnZipNotify)
            Call g_xobjGEMUS.SetRezip(g_blnReZip, g_blnRezipAsk, ge_RezipAdditional)

            strOldSBText = sbStatus.Panels(4).Text
            sbStatus.Panels(4).Text = "Launching..."
            MousePointer = vbHourglass
            cmdExtra.Enabled = False
            DoEvents

            strError = g_xobjGEMUS.RunEmulator(-1, strExtrasFile _
            , .lngGF_ImageIndex, strRunnableFile, g_udtPathGBGame.strRealPath, True, lngPlayersFrom _
            , lngPlayersTo, .lngGF_Control, .strGF_Comment, .strGF_VComment, .lngGF_PalNTSC, .blnGF_TDE, .strGF_Notes)

            MousePointer = vbDefault
            sbStatus.Panels(4).Text = strOldSBText
            cmdExtra.Enabled = True
            DoEvents

            'was there an error?
            If strError = "" Then
                'save default filetorun if changed
                If UCase$(strRunnableFile) <> UCase$(strRunnableFileBefore) Then
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT FileToRun FROM Extras WHERE EX_Id = " & lngExtraID_in, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !FileToRun = strRunnableFile
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                End If
            Else
                MsgBox strError, vbInformation, gc_strAPPNAME
            End If
        End With

    ElseIf lngType = 2 Then
    
        '================
        'USE MUSIC PLAYER
        '================
        
        Call RunMusicExtra(strExtrasFile, True)
    
    ElseIf lngType = 3 Then

        '===
        'URL
        '===

        Call g_xobjFunc.ShellExecuteFile(strExtrasFile, "", False)

    Else

        '============
        'INVALID TYPE
        '============

        MsgBox "Unknown Extras Type!", vbExclamation, gc_strAPPNAME

    End If
    
    Exit Sub

Err_Handler:
    MsgBox Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME

End Sub

Private Sub RunMusicExtra(strFullPathFile_in As String, ByVal blnSetFocusToList_in As Boolean)

    '------------------------------------------------------------------------
    'play the selected music extra with the first valid player in the list of players
    '------------------------------------------------------------------------

    Dim lngPlayerIdx As Long    'loop index
    Dim dblProcessID As Double  'process ID of shelled player
    Dim lngTotalPlayed As Long  'for stats
    Dim lngHandle As Long

    On Error GoTo Err_Handler

    If g_lngTotalSidPlayers > 0 Then
        If g_xobjFunc.FileExists(strFullPathFile_in, False, "") Then
            'find first valid player
            lngPlayerIdx = GetAppropriatePlayer(strFullPathFile_in)
            If lngPlayerIdx > -1 Then
                With ga_udtSidPlayers(lngPlayerIdx)
                    'now shell the player
                    If Not .blnUseShortFilenames Then
                        dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, strFullPathFile_in, False), vbNormalFocus)
                    Else
                        dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, strFullPathFile_in, True), vbNormalFocus)
                    End If
                    'store the process handle
                    lngHandle = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                    Call AddToSidPlayerHandles(lngHandle)
                    'modify stats
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed, False, 0
                    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed + 1, False
                End With
            Else
                Call MsgBox(g_xobjFunc.GetINIMessage("Error_Messages", "Error_NoMusicPlayersForType", "There were no Music Players found that support this Music File.  You can set the supported file types in Options!"), vbInformation, gc_strAPPNAME)
            End If
        End If
    End If
    
    'set focus back to the listview?
    If blnSetFocusToList_in Then
        lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SetFocus
    End If

    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbExclamation, gc_strAPPNAME

End Sub

Private Sub SaveRecentMenuList()

    Dim lngMenuIdx As Long  'loop index
    Dim lngClearIdx As Long 'loop index

    'save the items
    For lngMenuIdx = 0 To 9
        If ma_udtRecent(lngMenuIdx).strKey <> "" Then
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngMenuIdx & "Name", ma_udtRecent(lngMenuIdx).strName, True
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngMenuIdx & "Key", ma_udtRecent(lngMenuIdx).strKey, True
        Else
            Exit For
        End If
    Next lngMenuIdx

    'clear surplus items
    For lngClearIdx = (lngMenuIdx + 1) To 9
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngClearIdx & "Name", "", True
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\RecentItems", lngClearIdx & "Key", "", True
    Next lngClearIdx

End Sub
Private Sub RefreshExtraColumns(ByVal lngListViewIdx_in As Long, LItem As ListItem)

    Dim lngIdx As Long

    With ga_udtViewData(lstView.ListIndex)
        For lngIdx = 0 To .lngExtraColumnsCount - 1
            With .a_udtExtraColumns(lngIdx)
                If lvMain(lngListViewIdx_in).ColumnHeaders.Count > lngIdx + 1 Then
                    If g_eWhatIsSelected = e_SelType_IsGame Then
                        If .lngTypeIndex = 0 And .lngFieldIndex = 0 Then
                            'publisher
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strPublisher
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 1 Then
                            'year
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strYear
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 2 Then
                            'developer
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strDeveloper
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 3 Then
                            'programmer
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strProgrammer
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 4 Then
                            'musician
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strMusician
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 5 Then
                            'artist
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strArtist
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 6 Then
                            'license
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strLicense
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 7 Then
                            'genre
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strGenre
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 8 Then
                            'players
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strPlayers
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 9 Then
                            'control
                            LItem.SubItems(lngIdx + 1) = ga_strControlLookupTable(g_udtGame.lngControl)
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 10 Then
                            'language
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strLanguage
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 11 Then
                            'rarity
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strRarity
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 12 Then
                            'more info name
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strWebLinkName
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 13 Then
                            'more info url
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strWebLinkURL
                        ElseIf .lngTypeIndex = 0 And .lngFieldIndex = 14 Then
                            'comment
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strComment
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 0 Then
                            'game filename
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtGame.strGameFilename = g_strNone), "", g_udtGame.strGameFilename)
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 1 Then
                            'cracker
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strCracker
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 2 Then
                            'trainers
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strTrainers
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 3 Then
                            'playable
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strPlayable
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 4 Then
                            'game length
                            LItem.SubItems(lngIdx + 1) = g_udtGame.lngGameLength & " " & ga_strLengthTypeLookupTable(g_udtGame.lngGameLengthType)
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 5 Then
                            'pal/ntsc
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strNTSC
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 6 Then
                            'download name
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strVWebLinkName
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 7 Then
                            'download url
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strVWebLinkURL
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 8 Then
                            'original
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strOriginal
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 9 Then
                            'loading screen
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strLoadingScreen
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 10 Then
                            'title screen
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strTitleScreen
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 11 Then
                            'included docs
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strDocs
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 12 Then
                            'highscore saver
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strHighScoreSaver
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 13 Then
                            'true drive emulation
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strTrueDriveEmulation
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 14 Then
                            'version comment
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strVersionComment
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 16 Then
                            'music filename
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtSid.strFilename = g_strNone), "", g_udtSid.strFilename)
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 17 Then
                            'picture filename
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtGame.strScreenshotFilename = g_strNone), "", g_udtGame.strScreenshotFilename)
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 18 Then
                            'musician nick
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strNick
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 19 Then
                            'musician group
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strGroup
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 20 Then
                            'musician photo
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strPhoto
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 0 Then
                            'highscore
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strHighScore
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 1 Then
                            'difficulty
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strDifficulty
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 2 Then
                            'times played
                            LItem.SubItems(lngIdx + 1) = g_udtGame.strTimesPlayed
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 3 Then
                            'date last played
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtGame.strDateLastPlayed = g_strNever), "", Format$(g_udtGame.strDateLastPlayed, "dd-mm-yyyy"))
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 4 Then
                            'rating
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtGame.blnClassic = True), ga_strRatingLookupTable(6), g_udtGame.strRating)
                        ElseIf .lngTypeIndex = 2 And .lngFieldIndex = 5 Then
                            'review rating
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtGame.lngReviewRating = -1), "(" & g_strUnknown & ")", g_udtGame.lngReviewRating & "%")
                        End If
                    ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
                        If .lngTypeIndex = 0 And .lngFieldIndex = 4 Then
                            'musician
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strMusician
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 16 Then
                            'music filename
                            LItem.SubItems(lngIdx + 1) = IIf((g_udtSid.strFilename = g_strNone), "", g_udtSid.strFilename)
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 18 Then
                            'nick
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strNick
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 19 Then
                            'grp
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strGroup
                        ElseIf .lngTypeIndex = 1 And .lngFieldIndex = 20 Then
                            'photo
                            LItem.SubItems(lngIdx + 1) = g_udtSid.strPhoto
                        End If
                    End If
                Else
                    Exit For
                End If
            End With
        Next lngIdx
    End With

End Sub
Private Sub SetFocusToListview(ByVal lngListViewIdx_in As Long)

    On Error Resume Next

    If lvMain(lngListViewIdx_in).Enabled = True Then
        lvMain(lngListViewIdx_in).SetFocus
    End If

End Sub

Public Sub SetQuickFilter(ByVal blnSetItem_in As Boolean)

    Dim lngViewIdx As Long          'index of Quick Filter view in global views array
    Dim lngFilterIdx As Long        'loop index
    Dim strRegIdx As String         'registry value index for filter data
    Dim lngListViewIdx As Long      'current list view control index
    Dim strKeyName As String
    Dim lngIdx As Long

    ''''''''''''''''''''''''''
    'NEW QUICK FILTER SELECTED
    ''''''''''''''''''''''''''

    If Left$(g_udtQuickFilter.strTitle, 1) <> "[" And Right$(g_udtQuickFilter.strTitle, 1) <> "]" Then
        g_udtQuickFilter.strTitle = "[" & g_udtQuickFilter.strTitle & "]"
    End If

    'Get the index in ga_udtViewData() to store the new view
    If m_blnLastViewIsQuick Then
        'quick filter view already loaded!
        lngViewIdx = UBound(ga_udtViewData)
    Else
        'load storage for a new view
        lngViewIdx = UBound(ga_udtViewData) + 1
        ReDim Preserve ga_udtViewData(lngViewIdx)
        'signal that the last view is now a quick filter view
        m_blnLastViewIsQuick = True
    End If

    With ga_udtViewData(lngViewIdx)
    
        ''''''''''''''''''''''''''''''''''
        'store the QUERY data for the view
        ''''''''''''''''''''''''''''''''''
        
        'save to global views array
        .lngFilterCount = g_udtQuickFilter.lngNumFilters
        ReDim .a_udtFilters(.lngFilterCount)
        For lngFilterIdx = 1 To .lngFilterCount
            With .a_udtFilters(lngFilterIdx)
                .strClauseData = g_udtQuickFilter.strClauseData(lngFilterIdx - 1)
                .strFieldName = g_udtQuickFilter.strFieldName(lngFilterIdx - 1)
                .strFieldTable = g_udtQuickFilter.strFieldTable(lngFilterIdx - 1)
                .strMusicFieldName = g_udtQuickFilter.strMusicFieldName(lngFilterIdx - 1)
                .strMusicFieldTable = g_udtQuickFilter.strMusicFieldTable(lngFilterIdx - 1)
                .eOperator = g_udtQuickFilter.eOperator(lngFilterIdx - 1)
                .eClauseType = g_udtQuickFilter.eClauseType(lngFilterIdx - 1)
            End With
        Next lngFilterIdx
        .eFilterMode = g_udtQuickFilter.eFilterMode
        .eIncludeMusicTable = g_udtQuickFilter.eTablesToInclude
        .lngID = -1
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If blnSetItem_in Then
            .strSelectedItem = lvMain(lngListViewIdx).SelectedItem.Key
        Else
            .strSelectedItem = "INVALID"
        End If
        .strTitle = g_udtQuickFilter.strTitle

        'save to registry
        strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter"
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "NumFilters", .lngFilterCount, False
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "Title", .strTitle, True
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "TablesToInclude", .eIncludeMusicTable, False
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterMode", g_udtQuickFilter.eFilterMode, False
        For lngFilterIdx = 1 To .lngFilterCount
            strRegIdx = IIf((lngFilterIdx = 1), "", Trim$(lngFilterIdx))
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "ClauseType" & strRegIdx, g_udtQuickFilter.eClauseType(lngFilterIdx - 1), False
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "Operator" & strRegIdx, g_udtQuickFilter.eOperator(lngFilterIdx - 1), False
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterField" & strRegIdx, g_udtQuickFilter.strFieldName(lngFilterIdx - 1), True
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterMusicField" & strRegIdx, g_udtQuickFilter.strMusicFieldName(lngFilterIdx - 1), True
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterTable" & strRegIdx, g_udtQuickFilter.strFieldTable(lngFilterIdx - 1), True
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterMusicTable" & strRegIdx, g_udtQuickFilter.strMusicFieldTable(lngFilterIdx - 1), True
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", strKeyName, "FilterText" & strRegIdx, g_udtQuickFilter.strClauseData(lngFilterIdx - 1), True
        Next lngFilterIdx
    
        '''''''''''''''''''''''''''''''''
        'copy across view prefs if needed
        '''''''''''''''''''''''''''''''''

        If lngViewIdx > lstView.ListIndex Then
            'the current view is not the Quick Filter view, so copy the view prefs across
            .lngCol1Width = ga_udtViewData(lstView.ListIndex).lngCol1Width
            .lngListViewType = ga_udtViewData(lstView.ListIndex).lngListViewType
            .lngSortColumn = ga_udtViewData(lstView.ListIndex).lngSortColumn
            .lngSortOrder = ga_udtViewData(lstView.ListIndex).lngSortOrder
            .strSelectedItem = lvMain(lngListViewIdx).SelectedItem.Key
            .lngExtraColumnsCount = ga_udtViewData(lstView.ListIndex).lngExtraColumnsCount
            .lngMaxExtraColumns = ga_udtViewData(lstView.ListIndex).lngMaxExtraColumns
            .strExtraColumns = ga_udtViewData(lstView.ListIndex).strExtraColumns
            If .lngExtraColumnsCount > 0 Then
                ReDim Preserve .a_udtExtraColumns(.lngExtraColumnsCount - 1)
                For lngIdx = 0 To .lngExtraColumnsCount - 1
                    .a_udtExtraColumns(lngIdx) = ga_udtViewData(lstView.ListIndex).a_udtExtraColumns(lngIdx)
                Next lngIdx
            End If
        End If

    End With
    
    ''''''''''''''''''''''''
    'ADD AND SELECT THE VIEW
    ''''''''''''''''''''''''
    If lngViewIdx > lstView.ListCount - 1 Then
        lstView.AddItem g_udtQuickFilter.strTitle & "  (?)"
    Else
        lstView.List(lngViewIdx) = g_udtQuickFilter.strTitle & "  (?)"
    End If
    lstView.ListIndex = gc_lngNOLISTINDEX
    lstView.ListIndex = lngViewIdx

End Sub

Private Sub LoadPaths()

    '-------------------------------------------------
    'reads in the paths and exes to use in the program
    '-------------------------------------------------

    Dim strSidPath As String            'sid path retrieval
    Dim strPhotoPath As String          'photo path retrieval
    Dim lngNumFiles As Long             'for emulator files retrieval
    Dim strFiles() As String            'for emulator files retrieval
    Dim strPaths() As String            'paths retrieval array
    Dim strNames() As String
    Dim strExes() As String             'exes retrieval array
    Dim strTypes As String              'for music supported filetypes
    Dim a_strTypes() As String          'for music supported filetypes
    Dim strCLP() As String              'command line parameters array
    Dim strScript() As String           'script string array
    Dim blnUseShort() As Boolean        'short filename flags
    Dim lngPathIdx As Long              'loop index
    Dim lngFileIdx As Long              'loop index
    Dim strPath As String
    Dim lngTotalTypes As Long
    Dim strSection As String
    Dim strSections() As String
    Dim lngPlayerIdx As Long
    Dim strGBGamePath As String

    'GBGame path
    strGBGamePath = g_xobjFunc.INI_GetSetting("Paths", "GBGame", "", g_strGameBaseDotINI)
    If strGBGamePath = "" Then
        strGBGamePath = GetDefaultGBGameFolder()
    End If
    On Error Resume Next
    Call MkDir(strGBGamePath)
    On Error GoTo 0
    Call GetPathInfo(strGBGamePath, g_udtPathGBGame)
    If g_udtPathGBGame.ePathType = e_PathLocation_Invalid Then
        'not valid, try to use default path
        If Not g_xobjFunc.DirExists(GetDefaultGBGameFolder()) Then
            If Not CreateDefaultGBGameFolder() Then
                'could not create default folder
                MsgBox "Invalid temp GBGame path in """ & g_strGameBaseDotINI & """:" & vbCrLf & vbCrLf & _
                    "[Paths]" & vbCrLf & "GBGame=" & strGBGamePath, vbCritical Or vbOKOnly, gc_strAPPNAME
                g_dbGamebase.Close
                Unload frmSplash
                Set frmSplash = Nothing
                End
            End If
        End If
        'using default as specified does not exist
        Call GetPathInfo(GetDefaultGBGameFolder(), g_udtPathGBGame)
        MsgBox "Invalid temp GBGame path in """ & g_strGameBaseDotINI & """:" & vbCrLf & vbCrLf & _
            "[Paths]" & vbCrLf & "GBGame=" & strGBGamePath & vbCrLf & vbCrLf & _
            "Using this instead:" & vbCrLf & vbCrLf & """" & GetDefaultGBGameFolder() & """", vbExclamation Or vbOKOnly, gc_strAPPNAME
    End If

    'Music Path
    strSidPath = g_xobjFunc.INI_GetSetting("Misc", "Music", "", g_strPathFile)
    Call GetPathInfo(strSidPath, g_udtRegPathSid)

    'Photos Path
    strPhotoPath = g_xobjFunc.INI_GetSetting("Misc", "Photos", "", g_strPathFile)
    Call GetPathInfo(strPhotoPath, g_udtRegPathPhotos)
    
    'game paths
    lngPathIdx = 1
    g_lngTotalGamePaths = 0
    ReDim strPaths(0)
    Do
        strPath = g_xobjFunc.INI_GetSetting("Games", Trim$(lngPathIdx), "", g_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        g_lngTotalGamePaths = lngPathIdx - 1
        ReDim ga_udtRegPathGame(g_lngTotalGamePaths - 1)
        For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ga_udtRegPathGame(lngPathIdx))
        Next lngPathIdx
    End If
    g_blnValidGamePathExists = False
    For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
        If ga_udtRegPathGame(lngPathIdx).ePathType <> e_PathLocation_Invalid Then
            g_blnValidGamePathExists = True
            Exit For
        End If
    Next lngPathIdx
        
    'picture paths
    lngPathIdx = 1
    g_lngTotalScrnShotPaths = 0
    ReDim strPaths(0)
    Do
        strPath = g_xobjFunc.INI_GetSetting("Pictures", Trim$(lngPathIdx), "", g_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        g_lngTotalScrnShotPaths = lngPathIdx - 1
        ReDim ga_udtRegPathScrnShot(g_lngTotalScrnShotPaths - 1)
        For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ga_udtRegPathScrnShot(lngPathIdx))
        Next lngPathIdx
    End If

    'extras paths
    lngPathIdx = 1
    g_lngTotalExtrasPaths = 0
    ReDim strPaths(0)
    Do
        strPath = g_xobjFunc.INI_GetSetting("Extras", Trim$(lngPathIdx), "", g_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        g_lngTotalExtrasPaths = lngPathIdx - 1
        ReDim ga_udtRegPathExtras(g_lngTotalExtrasPaths - 1)
        For lngPathIdx = 0 To (g_lngTotalExtrasPaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ga_udtRegPathExtras(lngPathIdx))
        Next lngPathIdx
    End If

    'emulators
    Call LoadGemusMenu
    Call LoadEmulatorPopupMenu

    'music players

    'get list of music player names (sections)
    g_lngTotalSidPlayers = 0
    ReDim strSections(0)
    lngPathIdx = 1
    Do
        strSection = g_xobjFunc.INI_GetSetting("Players", Trim$(lngPathIdx), "", g_strMusicFile)
        If strSection <> "" Then
            ReDim Preserve strSections(lngPathIdx - 1)
            strSections(lngPathIdx - 1) = strSection
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    'read in their details
    If lngPathIdx > 1 Then
        g_lngTotalSidPlayers = lngPathIdx - 1
        ReDim ga_udtSidPlayers(g_lngTotalSidPlayers - 1)
        ReDim strNames(g_lngTotalSidPlayers - 1)
        ReDim strExes(g_lngTotalSidPlayers - 1)
        ReDim blnUseShort(g_lngTotalSidPlayers - 1)
        ReDim strCLP(g_lngTotalSidPlayers - 1)
        lngPathIdx = 0
        With g_xobjFunc
            For lngPlayerIdx = 0 To (g_lngTotalSidPlayers - 1)
                blnUseShort(lngPathIdx) = (.INI_GetSetting(strSections(lngPlayerIdx), "UseShortFilenames", "0", g_strMusicFile) = "1")
                strNames(lngPathIdx) = strSections(lngPlayerIdx)
                If strNames(lngPathIdx) = "" Then
                    strNames(lngPathIdx) = "- ??? -"
                End If
                strExes(lngPathIdx) = .INI_GetSetting(strSections(lngPlayerIdx), "PlayerPathFile", "", g_strMusicFile)
                strCLP(lngPathIdx) = .INI_GetSetting(strSections(lngPlayerIdx), "CommandLineParams", "%musicfile%", g_strMusicFile)
                strTypes = .INI_GetSetting(strSections(lngPlayerIdx), "SupportedTypes", "", g_strMusicFile)
                Call SetSidFileTypes(strTypes, a_strTypes(), lngTotalTypes)
                Call GetSidPlayerInfo(strCLP(lngPathIdx), blnUseShort(lngPathIdx), strNames(lngPathIdx), strExes(lngPathIdx), ga_udtSidPlayers(lngPathIdx), a_strTypes(), lngTotalTypes)
                lngPathIdx = lngPathIdx + 1
            Next lngPlayerIdx
        End With
    End If
    Call LoadSidPlayerPopupMenu

End Sub
Private Sub LoadQuickFilterView()

    '-----------------------------------------------------------
    'loads the quick filter view from the registry if it exists
    'and sets a flag to sepcify that the last loaded custom view
    'is a Quick Filter view.
    '-----------------------------------------------------------

    Dim lngViewIdx As Long                  'Quick Filter view index in the global views array
    Dim lngNumFilters As Long               'number of filters in the quickview
    Dim strRegIdx As String                 'registry index for filter value lookup
    Dim lngColumnIdx As Long
    Dim strKeyName As String
    Dim strExtraColumns As String
    Dim a_strColumnWidths() As Long

    strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter"
    
    'is the quick view there?
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "NumFilters", lngNumFilters, False, 0
    m_blnLastViewIsQuick = (lngNumFilters > 0)
    If m_blnLastViewIsQuick Then
        
        '''''''''''''''''''''
        'quick filter exists!
        '''''''''''''''''''''
        
        'add extra storage to the views array
        lngViewIdx = UBound(ga_udtViewData) + 1
        ReDim Preserve ga_udtViewData(lngViewIdx)
        
        'fill values from registry
        With ga_udtViewData(lngViewIdx)
            'general
            .lngID = -1
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterMode", .eFilterMode, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "Title", .strTitle, True, ""
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "TablesToInclude", .eIncludeMusicTable, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "Col1Width", .lngCol1Width, False, gc_lngDEFAULTCOLUMNWIDTH
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "ListViewType", .lngListViewType, False, lvwReport
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SortColumn", .lngSortColumn, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SortOrder", .lngSortOrder, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SelectedItem", .strSelectedItem, True, "INVALID"
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "ExtraColumns", strExtraColumns, True, ""
            
             'parse the extra columns
             .strExtraColumns = strExtraColumns
            .lngExtraColumnsCount = ParseExtraColumns(.a_udtExtraColumns, strExtraColumns)
            .lngMaxExtraColumns = UBound(.a_udtExtraColumns)

            'now the number of columns is known, get the extra columns widths
            For lngColumnIdx = 0 To .lngExtraColumnsCount - 1
                g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM _
                    strKeyName, _
                    "Col" & (lngColumnIdx + 2) & "Width", _
                    .a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                    False, _
                    gc_lngDEFAULTEXTRACOLUMNWIDTH
            Next lngColumnIdx

            'filters
            .lngFilterCount = lngNumFilters
            ReDim .a_udtFilters(.lngFilterCount)
            For lngViewIdx = 1 To .lngFilterCount
                strRegIdx = IIf((lngViewIdx = 1), "", Trim$(lngViewIdx))
                With .a_udtFilters(lngViewIdx)
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterField" & strRegIdx, .strFieldName, True, ""
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterTable" & strRegIdx, .strFieldTable, True, ""
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterMusicField" & strRegIdx, .strMusicFieldName, True, ""
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterMusicTable" & strRegIdx, .strMusicFieldTable, True, ""
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "FilterText" & strRegIdx, .strClauseData, True, ""
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "Operator" & strRegIdx, .eOperator, False, e_Op_Equal
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "ClauseType" & strRegIdx, .eClauseType, False, e_ClauseType_FieldContainsText
                End With
            Next lngViewIdx
        End With
        
    End If

End Sub
Private Function LoadSidFileInfo(blnPSID_out As Boolean) As Boolean

    '----------------------------------------------------------
    'loads the Sid Information for a game in the Games table
    'Only loads the info if it hasn't done it already
    '
    'Returns:       True if the Sid-File exists
    '               False otherwise
    '----------------------------------------------------------

    'get the Sid file info
    With g_udtSid
        If g_udtRegPathSid.ePathType <> e_PathLocation_Invalid Then
            blnPSID_out = g_xobjFunc.GetSidInfo(g_udtRegPathSid.strWithBackSlash & .strFilename, .strSFName, .strSFAuthor, .strSFCopyright, .lngSFTotalSongs, .lngSFDefaultSong)
            If blnPSID_out = False Then
                .blnIsPSID = False
                If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & .strFilename, False, "") = True Then
                    LoadSidFileInfo = True
                Else
                    LoadSidFileInfo = False
                End If
            Else
                .blnIsPSID = True
                LoadSidFileInfo = True
            End If
        Else
            .blnIsPSID = False
            LoadSidFileInfo = False
        End If
    End With

End Function


Private Sub MakeGB64SQLFilterString(a_udtFilters_in() As CVFilterType, lngFilterCount_in As Long, a_strBuffers_out() As String)

    '-------------------------------------------------------
    'Copies the selected array of SQL filters into a fixed
    'length structure and then transfers it into a
    'buffer string to send to GBSql.clsQuery in order to
    'make the SQL string for the required view.
    '
    'Inputs:    a_udtFilters_in():  The array of filters to
    '                               transfer into the buffer
    '           lngFilterCount_in:  The number of filters
    '
    'Outputs:   a_strBuffers_out(): Array of buffers
    '                               containing the filter
    '                               data.
    '-------------------------------------------------------

    Dim udtFixed As CVFilterFixedType               'the fixed length structure
    Dim lngFilterIdx As Long                        'loop index
    Dim strBuffer As CVFilterFixedBufferType        'buffer for fixed length data

    If lngFilterCount_in = 0 Then
        'no filters - return an empty string
        ReDim a_strBuffers_out(0)

    Else
        'Transfer filters into Fixed length array
        ReDim a_strBuffers_out(lngFilterCount_in)
        For lngFilterIdx = 1 To lngFilterCount_in   'filters start at index 1
            'copy structure
            With udtFixed
                .eClauseType = a_udtFilters_in(lngFilterIdx).eClauseType
                .eOperator = a_udtFilters_in(lngFilterIdx).eOperator
                .strClauseData = a_udtFilters_in(lngFilterIdx).strClauseData
                .strFieldName = a_udtFilters_in(lngFilterIdx).strFieldName
                .strFieldTable = a_udtFilters_in(lngFilterIdx).strFieldTable
                .strMusicFieldName = a_udtFilters_in(lngFilterIdx).strMusicFieldName
                .strMusicFieldTable = a_udtFilters_in(lngFilterIdx).strMusicFieldTable
            End With
            'transfer to buffer
            LSet strBuffer = udtFixed
            a_strBuffers_out(lngFilterIdx) = strBuffer.strBuffer
        Next lngFilterIdx

    End If

End Sub

Private Sub RemoveFromFavourites()

    '---------------------------------------------------------
    'removes the selected game or sid from the favourites view
    '---------------------------------------------------------

    Dim lngViewIdx As Long  'the selected view
    Dim strRemovedKey As String
    Dim blnClassic As Boolean
    
    'only remove if in favourites view!
    If lstView.ListIndex = gc_lngFAVOURITESINDEX Then
        'gotta have a selected item!
        lngViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If Not (lvMain(lngViewIdx).SelectedItem Is Nothing) Then
            'remove from favourites view
            If g_eWhatIsSelected = e_SelType_IsGame Then
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Fav FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                strRemovedKey = "G" & g_udtGame.lngID
                blnClassic = g_udtGame.blnClassic
            Else
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT SFav AS Fav FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
                strRemovedKey = "M" & g_udtSid.lngID
                blnClassic = False
            End If
            With g_rsGames
                If Not .EOF Then
                    .MoveFirst
                    .Edit
                    !Fav = False
                    .Update
                End If
                .Close
            End With
            Set g_rsGames = Nothing
            'remove item from the list
            Call RemoveItemFromList(lngViewIdx, lstView.ListIndex, lvMain(lngViewIdx).SelectedItem.Index, True)
            'remove bold property if game is not a classic
            If Left$(strRemovedKey, 1) = "G" Then
                If blnClassic = False Then
                    Call SetItemBoldInAllViews(strRemovedKey, False)
                End If
            Else
                Call SetItemBoldInAllViews(strRemovedKey, False)
            End If
        End If
    End If

End Sub

Private Sub RemoveItemFromList(ByVal lngListViewIdx_in As Long, ByVal lngViewIdx_in As Long, ByVal lngItemIdxToRemove_in As Long, ByVal blnClickNewItem_in As Boolean)
    
    '-------------------------------------------------------------------
    'Removes an item (game or sid) from the selected view's listview
    'and optionally selects and clicks another listitem.  Also amends
    'the number of items reported in the View Selector listbox.
    '
    'Inputs:    lngListViewIdx_in:      The index of the listview.
    '           lngItemIdxToRemove_in:  The index of the item to remove.
    '           blnClickNewItem_in:     Flag specifying whether to click
    '                                   the item.
    '-------------------------------------------------------------------
    
    Dim lngNewIndex As Long     'index of item to select after removal of item
    
    With lvMain(lngListViewIdx_in)
        
        'remove the item
        .ListItems.Remove (lngItemIdxToRemove_in)
    
        'set the new item or amend controls if listview now empty
        If .ListItems.Count = 0 Then
            'only disable the controls if no listitems left in the listview
            'because before we removed the item there WAS one item in the list
            Call SetControlsForView(False, lngListViewIdx_in)
            'amend the number of items in the View Selector and title
            lstView.List(lngViewIdx_in) = ga_udtViewData(lngViewIdx_in).strTitle & "  (" & g_strViewNone & ")"
            .ColumnHeaders(1).Text = lstView.List(lngViewIdx_in)
        Else
            'set the selected item to the nearest item to the one removed
            If lngItemIdxToRemove_in > .ListItems.Count Then
                lngNewIndex = lngItemIdxToRemove_in - 1
            Else
                lngNewIndex = lngItemIdxToRemove_in
            End If
            'amend the number of items in the View Selector and title
            lstView.List(lngViewIdx_in) = ga_udtViewData(lngViewIdx_in).strTitle & "  (" & .ListItems.Count & ")"
            .ColumnHeaders(1).Text = lstView.List(lngViewIdx_in)
            'click the item
            If blnClickNewItem_in Then
                Call SetItemInList(.ListItems(lngNewIndex).Key, False)
            End If
        End If
        
    End With
    
End Sub

Private Sub SaveCustomView(udtView_in As ViewDataType)

    '------------------------------------------------------------------
    'saves the passed in custom view to the database.  If the passed in
    'view doesn't exist in the database then a new record is created
    '
    'Inputs:    udtView_in:     The custom view to save
    '------------------------------------------------------------------
    
    Dim lngFilterIdx As Long    'loop index
    Dim strKeyName As String
    Dim lngColumnIdx As Long
    Dim lngOrdinal As Long

    '''''''''''''''''''
    'Save the View Data
    '''''''''''''''''''

    'get the next free ordinal
    lngOrdinal = 0
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Ordinal FROM ViewData ORDER BY Ordinal DESC", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            lngOrdinal = !Ordinal + 1
        End If
    End With
    Set g_rsGames = Nothing

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM ViewData WHERE ID = " & udtView_in.lngID, dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            'view exists so update
            .MoveFirst
            .Edit
        Else
            'view doesn't exist so create record
            .AddNew
            !Ordinal = lngOrdinal
        End If
        !FilterCount = udtView_in.lngFilterCount
        !FilterMode = udtView_in.eFilterMode
        !IncludeMusicTable = udtView_in.eIncludeMusicTable
        !ListViewType = udtView_in.lngListViewType
        !SelectedItem = udtView_in.strSelectedItem
        !SortColumn = udtView_in.lngSortColumn
        !SortOrder = udtView_in.lngSortOrder
        !Title = udtView_in.strTitle
        !ExtraColumns = udtView_in.strExtraColumns
        'store ID for filters (below)
        udtView_in.lngID = !ID
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    'column widths to the registry
    strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\CustomViews"
    g_xobjRegistry.SaveSetting _
        "HKEY_CURRENT_USER", _
        strKeyName, _
        "Col1Width_" & udtView_in.lngID, _
        udtView_in.lngCol1Width, _
        False

    For lngColumnIdx = 0 To udtView_in.lngExtraColumnsCount - 1
        g_xobjRegistry.SaveSetting _
            "HKEY_CURRENT_USER", _
            strKeyName, _
            "Col" & (lngColumnIdx + 2) & "Width_" & udtView_in.lngID, _
            udtView_in.a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
            False
    Next lngColumnIdx

    ''''''''''''''''''''''
    'Save the View Filters
    ''''''''''''''''''''''
    
    'delete any existing filters first
    g_dbGamebase.Execute ("DELETE * FROM ViewFilters WHERE ID = " & udtView_in.lngID)
    
    'add the view's filters
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM ViewFilters WHERE ID = " & udtView_in.lngID, dbOpenDynaset)
    With g_rsGames
        For lngFilterIdx = 1 To udtView_in.lngFilterCount
            .AddNew
            !ID = udtView_in.lngID
            !ClauseType = udtView_in.a_udtFilters(lngFilterIdx).eClauseType
            !Operator = udtView_in.a_udtFilters(lngFilterIdx).eOperator
            !ClauseData = udtView_in.a_udtFilters(lngFilterIdx).strClauseData
            !FieldName = udtView_in.a_udtFilters(lngFilterIdx).strFieldName
            !FieldTable = udtView_in.a_udtFilters(lngFilterIdx).strFieldTable
            !MusicFieldName = udtView_in.a_udtFilters(lngFilterIdx).strMusicFieldName
            !MusicFieldTable = udtView_in.a_udtFilters(lngFilterIdx).strMusicFieldTable
            .Update
        Next lngFilterIdx
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Private Sub SetControlsForView(ByVal blnViewHasItems_in As Boolean, ByVal lngListViewIdx_in As Long)

    '----------------------------------------------------------------
    'sets the controls enabled/disabled state depending on whether
    'there are any items in the current view's listview
    '
    'Inputs:    lngListViewIdx_in:      The currently selected view's
    '                                   listview control index.
    '           blnViewHasItems_in:     Whether the view is empty
    '                                   or not.
    '----------------------------------------------------------------

    'listview
    lvMain(lngListViewIdx_in).Enabled = blnViewHasItems_in

    'menus
    mnuEditFav.Enabled = blnViewHasItems_in
    mnuEditSidFile.Enabled = blnViewHasItems_in
    mnuEditFind.Enabled = blnViewHasItems_in
    mnuFileRemove.Enabled = blnViewHasItems_in
    mnuEditFindNext.Enabled = (blnViewHasItems_in And (m_strFindInList <> ""))
    mnuEditFindID.Enabled = blnViewHasItems_in
    mnuViewExportToTextFile.Enabled = blnViewHasItems_in
    mnuViewExportToRomLister.Enabled = blnViewHasItems_in

    'edit(1) and remove(2) custom view menus
    If lstView.ListIndex > gc_lngDEFAULTVIEWSMAX Then
        mnuViewEditView.Enabled = True
        mnuViewRemoveView.Enabled = True
    Else
        mnuViewEditView.Enabled = False
        mnuViewRemoveView.Enabled = False
    End If

    If blnViewHasItems_in = False Then
        'screenshots
        g_udtGame.strScreenshotFilename = g_strNone
        Call ClearScreenshots
        imgPhoto.Visible = False
        If Not g_blnMamePicStyle Then
            Call LoadScreenshots(g_udtGame.strScreenshotFilename)
        Else
            Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
        End If
        Call ShowScreenshots
        'tabs
        tabSid.Visible = False
        tabSid.Enabled = False
        tabDetails.Visible = False
        tabDetails.Enabled = False
        'button frame
        cmdPlay.Enabled = False
        cmdSid.Enabled = False
        cmdExtra.Enabled = False
        sbStatus.Panels(1).Text = ""
        sbStatus.Panels(2).Text = ""
        sbStatus.Panels(7).Text = ""
        'menus
        mnuEditGame.Enabled = False
        mnuEditVersion.Enabled = False
        mnuEditPersonal.Enabled = False
        mnuEditVersion.Enabled = False
        mnuEditPersonal.Enabled = False
        mnuEditLinkSet.Enabled = False
        mnuEditMark.Enabled = False
        mnuEditLinkClear.Enabled = False
        mnuEditGameFile.Enabled = False
        mnuEditScrnShots.Enabled = False
        mnuEditExtras.Enabled = False
    End If

    'random game select
    cmdRandom.Enabled = blnViewHasItems_in

End Sub

Private Sub SetInitialModularValues()

    Dim strNextFile As String       'for Dir()
    Dim lngIdx As Long

    'help menu
    m_strDBHelpFile = ""
    m_strDBHelpEmail = ""
    m_strDBHelpURL = ""

    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainWidth", m_sngRegWidth, False, gc_lngINVALIDWINDOWSIZE
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainHeight", m_sngRegHeight, False, gc_lngINVALIDWINDOWSIZE
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainTop", m_sngRegTop, False, gc_lngINVALIDWINDOWPOS
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainLeft", m_sngRegLeft, False, gc_lngINVALIDWINDOWPOS
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainMaximized", m_blnRegMaximized, False, False
        'show details flag
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "ShowDetails", m_blnRegDetails, False, True
        'show screenshots type
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "ScreenshotsType", m_lngRegScrnShots, False, 2
        'sidebar
        m_blnIsDraggingSidebarSplitter = False
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "Sidebar", m_lngRegSidebar, False, 1 'default on left
        m_sngSplitterWidth = picSplitter.Width
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "SidebarWidth", m_sngSidebarWidth, False, mc_sngDEFAULTSIDEBARWIDTH
        If m_sngSidebarWidth < mc_sngMINSIDEBARWIDTH Then m_sngSidebarWidth = mc_sngMINSIDEBARWIDTH
        'picSplitter.ZOrder = 0
        'play sound on classic game flag
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "SoundsOn", m_blnRegSounds, False, True
    End With
    
    'screenshots and photo state variables
    m_blnLastScrnShotWasNone = False
    m_blnLastPhotoWasNone = False
    m_ScrnShotNavButtonWidth = imgSSUp.Width / Screen.TwipsPerPixelX
    m_ScrnShotNavButtonHeight = imgSSUp.Height / Screen.TwipsPerPixelY
    picSSUp.Width = imgSSUp.Width
    picSSDown.Width = imgSSDown.Width
    picSSUp.Height = imgSSUp.Height
    picSSDown.Height = imgSSDown.Height
    m_blnScrnShotBugWorkaround = True 'm_blnScrnShotBugWorkaround = False
    
    'for non stop music
    m_strLastSidFilePlayed = ""
    m_blnRegNonStopMusic = False
    m_blnSettingView = False

    'drag/drop
    m_blnDragging = False
    
    'notes changing
    m_blnNotesChanged = False
    m_blnNotesByCode = True
    
    'status bar offset
    With sbStatus
        m_sngStatusBarOffset = .Panels(7).Width + .Panels(6).Width + .Panels(5).Width + 360
    End With
    
    'find
    m_strFindInList = ""
    m_lngFindIDGame = 0
    m_lngFindIDMusic = 0
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "FindCommentInList", m_lngFindCommentInList, False, 0
    If (m_lngFindCommentInList < 0) Or (m_lngFindCommentInList > 1) Then
        m_lngFindCommentInList = 0
    End If

    'flag to specify if we want to call SetView() from lstView_Click()
    m_blnSetTheView = True
    
    'set this to false only when deleting custom views as we don't want to save their prefs!
    m_blnSaveViewPrefs = True

    'marked game
    m_udtMarkedGame.lngID = 0
    m_udtMarkedGame.strName = ""

    'gfx on the details tabs
    m_sngTabIconsTop = (tabDetails.TabHeight / 2) - (5 * Screen.TwipsPerPixelY)
    m_sngTabIconsGap = Screen.TwipsPerPixelX
    m_sngTabIconsCupGap = 8 * Screen.TwipsPerPixelX
    m_sngTabIconsStart = tabDetails.Width - (13 * Screen.TwipsPerPixelX)
    m_sngHalfTabWidth = tabDetails.Width / 2
    m_sngTabIconsCupTop = lblName.Top + ((lblName.Height / 2) - ((Screen.TwipsPerPixelY * 11) / 2)) + (Screen.TwipsPerPixelY / 2)
    m_sngTabIconsCupWidth = 16 * Screen.TwipsPerPixelX

    'read in classic game sounds
    m_strLastSelectedKey = ""
    m_lngMaxSounds = 0 'first sound stored in position (1)
    strNextFile = Dir(g_strAppPathWithBackslash & g_strGameBaseName & "\Sounds\*.wav")
    Do While strNextFile <> ""
        m_lngMaxSounds = m_lngMaxSounds + 1
        ReDim Preserve ma_strSounds(m_lngMaxSounds)
        ma_strSounds(m_lngMaxSounds) = strNextFile
        strNextFile = Dir
    Loop

    '================== sidebar quick links ==================

    'load listview icons
    m_lngQLIconIdxPlayGame = 1      'play game
    m_lngQLIconIdxPlayMusic = 2     'play music
    m_lngQLIconIdxGemusExtra = 3    'gemus extra
    m_lngQLIconIdxMusicExtra = 4    'music extra
    m_lngQLIconIdxDefaultExtra = 5  'default extra
    m_lngQLIconIdxURLExtra = 10     'url
    m_lngQLIconIdxWeblink = 10      'weblink
    m_lngQLIconIdxWeblink = 10      'version weblink
    m_lngQLIconIdxPrequel = 11      'prequel game link
    m_lngQLIconIdxSequel = 12       'sequel game link
    m_lngQLIconIdxRelated = 13      'related game link
    m_lngQLIconIdxCloneOf = 14      'cloneof game link

    m_lngTotalQLNameIcons = 0
    ReDim ma_udtQLNameIcons(0)
    m_lngTotalQLIcons = 41
    ReDim Preserve ma_udtQLIcons(m_lngTotalQLIcons - 1)
    
    'audio files
    ma_udtQLIcons(0).strExtension = "aiff"
    ma_udtQLIcons(1).strExtension = "wav"
    ma_udtQLIcons(2).strExtension = "flac"
    ma_udtQLIcons(3).strExtension = "m4a"
    ma_udtQLIcons(4).strExtension = "ape"
    ma_udtQLIcons(5).strExtension = "wma"
    ma_udtQLIcons(6).strExtension = "mp2"
    ma_udtQLIcons(7).strExtension = "mp3"
    ma_udtQLIcons(8).strExtension = "aac"
    ma_udtQLIcons(9).strExtension = "mid"
    ma_udtQLIcons(10).strExtension = "m3u"
    ma_udtQLIcons(11).strExtension = "pls"
    ma_udtQLIcons(12).strExtension = "ogg"
    For lngIdx = 0 To 12
        ma_udtQLIcons(lngIdx).lngIconIdx = 6
    Next lngIdx
    
    'video files
    ma_udtQLIcons(13).strExtension = "avi"
    ma_udtQLIcons(14).strExtension = "mpg"
    ma_udtQLIcons(15).strExtension = "mpeg"
    ma_udtQLIcons(16).strExtension = "mpe"
    ma_udtQLIcons(17).strExtension = "mov"
    ma_udtQLIcons(18).strExtension = "mkv"
    ma_udtQLIcons(19).strExtension = "mp4"
    ma_udtQLIcons(20).strExtension = "vob"
    ma_udtQLIcons(21).strExtension = "rm"
    ma_udtQLIcons(22).strExtension = "flv"
    ma_udtQLIcons(23).strExtension = "swf"
    ma_udtQLIcons(24).strExtension = "wmv"
    For lngIdx = 13 To 24
        ma_udtQLIcons(lngIdx).lngIconIdx = 7
    Next lngIdx
    
    'image files
    ma_udtQLIcons(25).strExtension = "bmp"
    ma_udtQLIcons(26).strExtension = "gif"
    ma_udtQLIcons(27).strExtension = "ico"
    ma_udtQLIcons(28).strExtension = "jpg"
    ma_udtQLIcons(29).strExtension = "jpeg"
    ma_udtQLIcons(30).strExtension = "mng"
    ma_udtQLIcons(31).strExtension = "png"
    For lngIdx = 25 To 31
        ma_udtQLIcons(lngIdx).lngIconIdx = 8
    Next lngIdx
    
    'document files
    ma_udtQLIcons(32).strExtension = "txt"
    ma_udtQLIcons(33).strExtension = "doc"
    ma_udtQLIcons(34).strExtension = "pdf"
    ma_udtQLIcons(35).strExtension = "rtf"
    ma_udtQLIcons(36).strExtension = "csv"
    ma_udtQLIcons(37).strExtension = "docx"
    ma_udtQLIcons(38).strExtension = "html"
    ma_udtQLIcons(39).strExtension = "htm"
    ma_udtQLIcons(40).strExtension = "xhtml"
    For lngIdx = 32 To 40
        ma_udtQLIcons(lngIdx).lngIconIdx = 9
    Next lngIdx

End Sub
Private Function GetQLIconIdx(strPath As String)

    Dim lngIdx As Long
    Dim strExt As String

    strPath = LCase$(strPath)
    strExt = Right$(strPath, Len(strPath) - InStrRev(strPath, "."))
    
    For lngIdx = 0 To m_lngTotalQLIcons - 1
        If strExt = ma_udtQLIcons(lngIdx).strExtension Then
            GetQLIconIdx = ma_udtQLIcons(lngIdx).lngIconIdx
            Exit Function
        End If
    Next lngIdx
    
    GetQLIconIdx = m_lngQLIconIdxDefaultExtra 'default extra

End Function
Private Function GetQLNameIconIdx(strName_in As String, ByVal lngDefaultIconIdx_out As Long)

    Dim lngIdx As Long
    Dim strLower As String

    strLower = LCase$(strName_in)

    For lngIdx = 0 To m_lngTotalQLNameIcons - 1
        If strLower = ma_udtQLNameIcons(lngIdx).strExtension Then
            GetQLNameIconIdx = ma_udtQLNameIcons(lngIdx).lngIconIdx
            Exit Function
        End If
    Next lngIdx
    
    GetQLNameIconIdx = lngDefaultIconIdx_out

End Function
Private Sub SetSortCustomViewsEnabledState()

    Dim lngUboundCustomViews As Long

    lngUboundCustomViews = UBound(ga_udtViewData)
    If m_blnLastViewIsQuick Then lngUboundCustomViews = lngUboundCustomViews - 1
    pmnuView(6).Enabled = (lngUboundCustomViews > (gc_lngDEFAULTVIEWSMAX + 1)) 'has to be 2 custom views to be able to sort them

End Sub

Public Sub ShowQuickLinks()

    On Error GoTo Err_Handler

    Dim lngExtraIdx As Long
    Dim itmSelectedItem As ListItem
    Dim lngID As Long
    Dim lngIconIdx As Long

    Call SendMessageAny(lvLinks.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
    If lstView.ListIndex = gc_lngNOLISTINDEX Then
        Exit Sub
    End If
    
    'grab the selected item and ID
    Set itmSelectedItem = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If itmSelectedItem Is Nothing Then
        Exit Sub
    End If
    With itmSelectedItem
        lngID = Val(Right$(.Key, Len(.Key) - 1))
    End With

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'test if an item from the Games table or Music table was clicked
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If itmSelectedItem.Icon <> 5 Then

        'game links
        With g_udtGame
            If .lngPrequel > 0 Then
                lvLinks.ListItems.Add , "A" & .lngPrequel, .strPrequel, m_lngQLIconIdxPrequel, m_lngQLIconIdxPrequel
            End If
            If .lngSequel > 0 Then
                lvLinks.ListItems.Add , "B" & .lngSequel, .strSequel, m_lngQLIconIdxSequel, m_lngQLIconIdxSequel
            End If
            If .lngRelated > 0 Then
                lvLinks.ListItems.Add , "C" & .lngRelated, .strRelated, m_lngQLIconIdxRelated, m_lngQLIconIdxRelated
            End If
            If .lngCloneOf > 0 Then
                lvLinks.ListItems.Add , "D" & .lngCloneOf, .strCloneOf, m_lngQLIconIdxCloneOf, m_lngQLIconIdxCloneOf
            End If
        End With

        'weblinks
        With g_udtGame
            If .blnHasWebLink Then
                lvLinks.ListItems.Add , "W" & .lngID, .strWebLinkName, m_lngQLIconIdxURLExtra, m_lngQLIconIdxWeblink
            End If
            If .blnHasVWebLink Then
                lvLinks.ListItems.Add , "V" & .lngID, .strVWebLinkName, m_lngQLIconIdxURLExtra, m_lngQLIconIdxVersionWeblink
            End If
        End With

        'game file
        With g_udtGame
            If .blnGameFileExists Then
                lvLinks.ListItems.Add , "G" & .lngID, m_strQLPlayGame, m_lngQLIconIdxPlayGame, m_lngQLIconIdxPlayGame
            End If
        End With
    
        'music file
        With g_udtSid
            If .blnFileExists Then
                lvLinks.ListItems.Add , "M" & .lngID, m_strQLPlayMusic, m_lngQLIconIdxPlayMusic, m_lngQLIconIdxPlayMusic
            End If
        End With
  
        'extras
        If g_lngTotalExtrasPaths > 0 Then
            If g_udtGame.blnHasExtras = True Then
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT EX_Id, Name, Path, EA, Type FROM Extras WHERE GA_Id = " & g_udtGame.lngID & " ORDER BY DisplayOrder", dbOpenSnapshot)
                With g_rsGames
                    If Not .EOF Then
                        'add to the popup menu
                        .MoveFirst
                        Do Until .EOF
                            If !EA Then
                                If !Type = 3 Then
                                    lngIconIdx = GetQLNameIconIdx(!Name, m_lngQLIconIdxURLExtra)
                                ElseIf !Type = 2 Then
                                    lngIconIdx = GetQLNameIconIdx(!Name, m_lngQLIconIdxMusicExtra)
                                ElseIf !Type = 1 Then
                                    lngIconIdx = GetQLNameIconIdx(!Name, m_lngQLIconIdxGemusExtra)
                                Else
                                    lngIconIdx = GetQLNameIconIdx(!Name, GetQLIconIdx(!Path))
                                End If
                                lvLinks.ListItems.Add , "E" & !EX_Id, !Name, lngIconIdx, lngIconIdx
                            End If
                            .MoveNext
                        Loop
                        .Close
                    End If
                End With
                Set g_rsGames = Nothing
            End If
        End If
    Else

        'music file
        With g_udtSid
            If .blnFileExists Then
                lvLinks.ListItems.Add , "S" & .lngID, m_strQLPlayMusic, m_lngQLIconIdxPlayMusic, m_lngQLIconIdxPlayMusic
            End If
        End With

    End If

    Exit Sub

Err_Handler:

End Sub
Private Sub SetItemBoldInAllViews(strItemKey_in As String, ByVal blnSetAsBold_in As Boolean)

    '------------------------------------------------------------
    'Sets an item to either bold or not bold in all loaded views.
    'Omits the favourites view because they are always bold.
    '------------------------------------------------------------

    Dim lngListViewIdx As Long      'loop index

    For lngListViewIdx = 0 To UBound(ga_blnViewsLoaded)
        'only look for items in loaded views
        If ga_blnViewsLoaded(lngListViewIdx) = True Then
            'not favourites view
            If lngListViewIdx <> gc_lngFAVOURITESINDEX Then
                If ItemInListExists(strItemKey_in, lngListViewIdx) Then
                    lvMain(lngListViewIdx).ListItems(strItemKey_in).Bold = blnSetAsBold_in
                End If
            End If
        End If
    Next lngListViewIdx

End Sub

Private Sub RunGame(ByVal lngEmuIdx_in As Long)

    Dim strError As String
    Dim lngGPIdx As Long
    Dim strExt As String
    Dim lngDotPos As Long
    Dim strFileToRun As String
    Dim eExtractResult As Gemus.GB64ExtractResultEnum
    Dim Extractor As Gemus.clsExtract
    Dim strExtractedFilename As String
    Dim strFileList As String
    Dim strDefaultFile As String
    Dim strStatus As String
    Dim strArchiveCaption As String
    Dim strArchivePrompt As String
    Dim strStatusArchiveContents As String
    Dim strImageName As String
    Dim lngTotalPlayed As Long
    Dim strOldSBText As String

    'find actual game path\file
    If g_udtGame.blnGameFileExists = True Then
        Call ReleaseSidPlayers(True)
        MousePointer = vbHourglass
        DoEvents
        strError = FindGameFile(lngGPIdx, g_udtGame.strGameFilename)
        MousePointer = vbDefault
        DoEvents
        If strError <> "" Then
            'not found or aborted
            If strError <> "~~ABORTED~~" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
            End If
        Else
            'game found!
            strDefaultFile = g_udtGame.strFileToRun
            strFileToRun = ga_udtRegPathGame(lngGPIdx).strWithBackSlash & g_udtGame.strGameFilename
            With g_udtGame
                'get memo text if not already loaded
                If .blnLoaded_3 = False Then
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MemoText FROM Games WHERE GA_Id = " & .lngID, dbOpenSnapshot)
                    .strMemo = g_rsGames!MemoText
                    g_rsGames.Close
                    Set g_rsGames = Nothing
                End If
                'extract and run game
                Call g_xobjGEMUS.SetRezipNotify(g_blnZipNotify)
                Call g_xobjGEMUS.SetRezip(g_blnReZip, g_blnRezipAsk, ge_RezipAdditional)
                
                strOldSBText = sbStatus.Panels(4).Text
                sbStatus.Panels(4).Text = "Launching..."
                MousePointer = vbHourglass
                cmdPlay.Enabled = False
                DoEvents
                strError = g_xobjGEMUS.RunEmulator(lngEmuIdx_in, strFileToRun _
                , .lngGameFilenameIndex, strDefaultFile, g_udtPathGBGame.strRealPath, True, .lngPlayersFrom _
                , .lngPlayersTo, .lngControl, .strComment, .strVersionComment, .lngNTSC, IIf(.lngTrueDriveEmulation = 1, True, False), .strPairs)
                MousePointer = vbDefault
                sbStatus.Panels(4).Text = strOldSBText
                cmdPlay.Enabled = True
                DoEvents
            End With
            
            If strError = "" Then
                'save default filetorun if changed
                If UCase$(strDefaultFile) <> UCase$(g_udtGame.strFileToRun) Then
                    g_udtGame.strFileToRun = strDefaultFile
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT FileToRun FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !FileToRun = strDefaultFile
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                End If
                'modify stats
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT TimesPlayed, DateLastPlayed FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .MoveFirst
                        .Edit
                        !TimesPlayed = !TimesPlayed + 1
                        !DateLastPlayed = Format$(Now, "dd-mm-yyyy")
                        .Update
                        g_udtGame.strDateLastPlayed = GetFormattedDate(!DateLastPlayed)
                        g_udtGame.strTimesPlayed = Trim$(!TimesPlayed)
                        lbl_Info_LP = g_udtGame.strDateLastPlayed
                        lbl_Info_TP = g_udtGame.strTimesPlayed
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalGamesPlayed", lngTotalPlayed, False, 0
                g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalGamesPlayed", lngTotalPlayed + 1, False
                'add to the recent list?
                If (Not m_blnFormIsLoading) And (g_lngRecentUpdate = gc_lngRECENT_PLAYED) And (g_lngRecentStore <> gc_lngRECENT_SIDS) Then
                    Call AddToRecentList("G" & g_udtGame.lngID, g_udtGame.strName)
                End If
            Else
                MsgBox strError, vbInformation, gc_strAPPNAME
            End If
        End If
    Else
        MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInGamePaths", "The file was not found in any of your Game Paths!"), vbExclamation, gc_strAPPNAME
    End If

End Sub

Private Sub SetupHelpMenu()

    Dim strText As String
    Dim strItem As String
    Dim strINIFile As String

    strINIFile = g_strGameBasePathWBS & "Credits.ini"

    With g_xobjFunc
        If .FileExists(strINIFile, False, "") = False Then Exit Sub
        
        'help file
        strText = Trim$(.INI_GetSetting("Help Menu", "HelpFileText", "", strINIFile))
        If strText <> "" Then
            strItem = Trim$(.INI_GetSetting("Help Menu", "HelpFile", "", strINIFile))
            If strItem <> "" Then
                If .FileExists(g_strGameBasePathWBS & strItem, False, "") = True Then
                    'help file ok
                    If Right$(strText, 3) <> "..." Then strText = strText & "..."
                    With mnuHelpDB(0)
                        .Caption = strText
                        .Visible = True
                        .Enabled = True
                    End With
                    m_strDBHelpFile = g_strGameBasePathWBS & strItem
                End If
            End If
        End If
        
        'email address
        strText = Trim$(.INI_GetSetting("Help Menu", "HelpEmailText", "", strINIFile))
        If strText <> "" Then
            strItem = Trim$(.INI_GetSetting("Help Menu", "HelpEmail", "", strINIFile))
            If strItem <> "" Then
                If (InStr(strItem, "@") > 0) And (InStr(strItem, ".") > 0) Then
                    'email ok
                    With mnuHelpDB(2)
                        .Caption = strText
                        .Visible = True
                        .Enabled = True
                    End With
                    m_strDBHelpEmail = strItem
                End If
            End If
        End If
        
        'website address
        strText = Trim$(.INI_GetSetting("Help Menu", "HelpWebSiteText", "", strINIFile))
        If strText <> "" Then
            strItem = Trim$(.INI_GetSetting("Help Menu", "HelpWebSite", "", strINIFile))
            If strItem <> "" Then
                'website ok
                With mnuHelpDB(1)
                    .Caption = strText
                    .Visible = True
                    .Enabled = True
                End With
                m_strDBHelpURL = strItem
            End If
        End If
        
    End With
    
    If (m_strDBHelpURL <> "") Or (m_strDBHelpEmail <> "") Or (m_strDBHelpFile <> "") Then
        mnuHelpBar1.Visible = False
        mnuHelpBar3.Visible = True
    End If

End Sub

Private Sub ShowEditGameInfo(ByVal blnAllGames_in As Boolean)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            With g_udtEditGameInfo
                .blnEditAll = blnAllGames_in
                .blnSimPlay = g_udtGame.blnPlayersSim
                .lngControlID = g_udtGame.lngControl
                .lngGenreID = g_udtGame.lngGenreID
                .lngLanguageID = g_udtGame.lngLanguageID
                .lngMusicianID = g_udtSid.lngMU_Id
                .lngPlayersFrom = g_udtGame.lngPlayersFrom
                .lngPlayersTo = g_udtGame.lngPlayersTo
                .lngProgrammerID = g_udtGame.lngProgrammerID
                .lngPublisherID = g_udtGame.lngPublisherID
                .lngYearID = g_udtGame.lngYearID
                .lngArtistID = g_udtGame.lngArtistID
                .lngDeveloperID = g_udtGame.lngDeveloperID
                .lngLicenseID = g_udtGame.lngLicenseID
                .lngRarityID = g_udtGame.lngRarityID
                .strComment = g_udtGame.strComment
                .strWebLinkName = g_udtGame.strWebLinkName
                If .strWebLinkName = "(" & g_strNone & ")" Then
                    .strWebLinkName = ""
                End If
                .strWebLinkURL = g_udtGame.strWebLinkURL
                .lngReviewRating = g_udtGame.lngReviewRating
                .blnTableEdited = False
                .blnUpdateYear = False
                .blnUpdatePublisher = False
                .blnUpdateProgrammer = False
                .blnUpdateMusician = False
                .blnUpdateLanguage = False
                .blnUpdateArtist = False
                .blnUpdateDeveloper = False
                .blnUpdateLicense = False
                .blnUpdateRarity = False
                .blnUpdateWebLink = False
                .blnUpdateGenre = False
                .blnUpdateControl = False
                .blnUpdatePlayers = False
                .blnUpdateComment = False
                .blnUpdateReviewRating = False
                .blnOKClicked = False
            End With
            Screen.MousePointer = vbHourglass
            DoEvents
            frmGameInfoProps.tabProps.Tab = IIf((blnAllGames_in), 0, 1)
            frmGameInfoProps.Show vbModal, Me
            If g_udtEditGameInfo.blnOKClicked Then
                'SAVE CHANGES
                If g_udtEditGameInfo.blnEditAll = False Then
                    'just selected game - save to database
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !PlayersSim = g_udtEditGameInfo.blnSimPlay
                            !Control = g_udtEditGameInfo.lngControlID
                            !GE_Id = g_udtEditGameInfo.lngGenreID
                            !LA_Id = g_udtEditGameInfo.lngLanguageID
                            !MU_Id = g_udtEditGameInfo.lngMusicianID
                            !PlayersFrom = g_udtEditGameInfo.lngPlayersFrom
                            !PlayersTo = g_udtEditGameInfo.lngPlayersTo
                            !PR_Id = g_udtEditGameInfo.lngProgrammerID
                            !PU_Id = g_udtEditGameInfo.lngPublisherID
                            !YE_ID = g_udtEditGameInfo.lngYearID
                            !AR_Id = g_udtEditGameInfo.lngArtistID
                            !DE_Id = g_udtEditGameInfo.lngDeveloperID
                            !LI_Id = g_udtEditGameInfo.lngLicenseID
                            !RA_Id = g_udtEditGameInfo.lngRarityID
                            !WebLink_Name = g_udtEditGameInfo.strWebLinkName
                            !WebLink_URL = g_udtEditGameInfo.strWebLinkURL
                            !Comment = g_udtEditGameInfo.strComment
                            !ReviewRating = g_udtEditGameInfo.lngReviewRating
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    're-click game
                    Call SetItemInList(liSelected.Key, False)
                    Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                    Call UnCacheViews
                Else
                    'all games in view
                    Load frmBatchEdit
                    frmBatchEdit.m_blnUpdateGameInfo = True
                    frmBatchEdit.Show vbModal, Me
                    Call UnCacheViews
                    Call ClickLstView
                End If
            ElseIf g_udtEditGameInfo.blnTableEdited = True Then
                'just reclick game
                Call SetItemInList(liSelected.Key, False)
                Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                Call UnCacheViews
            End If
            Screen.MousePointer = vbDefault
        End If
    End If

End Sub
Private Sub ShowEditVersionInfo(ByVal blnAllGames_in As Boolean)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            With g_udtEditVersionInfo
                .blnEditAll = blnAllGames_in
                .lngDocs = g_udtGame.lngDocs
                .lngHSS = g_udtGame.lngHSS
                .lngOriginal = g_udtGame.lngOriginal
                .lngPlayable = g_udtGame.lngPlayable
                .lngTitleScreen = g_udtGame.lngTitleScreen
                .lngLoadingScreen = g_udtGame.lngLoadingScreen
                .lngTDE = g_udtGame.lngTrueDriveEmulation
                .lngCrackerID = g_udtGame.lngCrackerID
                .lngGameLength = g_udtGame.lngGameLength
                .lngGameLengthType = g_udtGame.lngGameLengthType
                .lngNTSC = g_udtGame.lngNTSC
                .lngTrainers = g_udtGame.lngTrainers
                .strVComment = g_udtGame.strVersionComment
                .strWebLinkName = g_udtGame.strVWebLinkName
                If .strWebLinkName = "(" & g_strNone & ")" Then
                    .strWebLinkName = ""
                End If
                .strWebLinkURL = g_udtGame.strVWebLinkURL
                .blnUpdateCracker = False
                .blnUpdateDocs = False
                .blnUpdateHSS = False
                .blnUpdateLength = False
                .blnUpdateLoadingScreen = False
                .blnUpdateNTSC = False
                .blnUpdateTDE = False
                .blnUpdateTrainers = False
                .blnUpdateVComment = False
                .blnOKClicked = False
                .blnTableEdited = False
            End With
            Screen.MousePointer = vbHourglass
            DoEvents
            frmVersionProps.tabProps.Tab = IIf((blnAllGames_in), 0, 1)
            frmVersionProps.Show vbModal, Me
            If g_udtEditVersionInfo.blnOKClicked Then
                'SAVE CHANGES
                If g_udtEditVersionInfo.blnEditAll = False Then
                    'just selected game - save to database
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !V_TitleScreen = g_udtEditVersionInfo.lngTitleScreen
                            !V_Original = g_udtEditVersionInfo.lngOriginal
                            !V_Playable = g_udtEditVersionInfo.lngPlayable
                            !V_IncludedDocs = g_udtEditVersionInfo.lngDocs
                            !V_HighScoreSaver = g_udtEditVersionInfo.lngHSS
                            !V_LoadingScreen = g_udtEditVersionInfo.lngLoadingScreen
                            !V_TrueDriveEmu = g_udtEditVersionInfo.lngTDE
                            !CR_Id = g_udtEditVersionInfo.lngCrackerID
                            !V_Length = g_udtEditVersionInfo.lngGameLength
                            !V_LengthType = g_udtEditVersionInfo.lngGameLengthType
                            !V_PalNTSC = g_udtEditVersionInfo.lngNTSC
                            !V_Trainers = g_udtEditVersionInfo.lngTrainers
                            !V_Comment = g_udtEditVersionInfo.strVComment
                            !V_WebLink_Name = g_udtEditVersionInfo.strWebLinkName
                            !V_WebLink_URL = g_udtEditVersionInfo.strWebLinkURL
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    're-click game
                    Call SetItemInList(liSelected.Key, False)
                    Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                    Call UnCacheViews
                Else
                    'all games in view
                    Load frmBatchEdit
                    frmBatchEdit.m_blnUpdateVersionInfo = True
                    frmBatchEdit.Show vbModal, Me
                    Call UnCacheViews
                    Call ClickLstView
                End If
            ElseIf g_udtEditVersionInfo.blnTableEdited = True Then
                'just re-click game
                Call SetItemInList(liSelected.Key, False)
                Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                Call UnCacheViews
            End If
            Screen.MousePointer = vbDefault
        End If
    End If

End Sub
Private Sub ShowEditMusicInfo(ByVal blnAllTunes_in As Boolean)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsSid Then
            With g_udtEditMusicInfo
                .blnAdult = g_udtSid.blnAdult
                .blnEditAll = blnAllTunes_in
                .blnFav = g_udtSid.blnFav
                .blnOKClicked = False
                .blnTableEdited = False
                .blnUpdateAdult = False
                .blnUpdateFav = False
                .blnUpdateMusician = False
                .lngMusicianID = g_udtSid.lngMU_Id
            End With
            Screen.MousePointer = vbHourglass
            DoEvents
            frmMusicProps.tabProps.Tab = IIf((blnAllTunes_in), 0, 1)
            frmMusicProps.Show vbModal, Me
            If g_udtEditMusicInfo.blnOKClicked Then
                'SAVE CHANGES
                If g_udtEditMusicInfo.blnEditAll = False Then
                    'just selected game - save to database
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !MU_Id = g_udtEditMusicInfo.lngMusicianID
                            !SFav = g_udtEditMusicInfo.blnFav
                            !Adult = g_udtEditMusicInfo.blnAdult
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    're-click game
                    liSelected.Bold = g_udtEditMusicInfo.blnFav
                    Call SetItemInList(liSelected.Key, False)
                    Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                    Call UnCacheViews
                Else
                    'all games in view
                    Load frmBatchEdit
                    frmBatchEdit.m_blnUpdateMusicInfo = True
                    frmBatchEdit.Show vbModal, Me
                    Call UnCacheViews
                    Call ClickLstView
                End If
            ElseIf g_udtEditMusicInfo.blnTableEdited = True Then
                'just re-click game
                Call SetItemInList(liSelected.Key, False)
                Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                Call UnCacheViews
            End If
            Screen.MousePointer = vbDefault
        End If
    End If

End Sub
Private Sub ShowEditPersonalInfo(ByVal blnAllGames_in As Boolean)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            With g_udtEditPersonalInfo
                .blnEditAll = blnAllGames_in
                .blnClassic = g_udtGame.blnClassic
                .blnAdult = g_udtGame.blnAdult
                .blnFav = g_udtGame.blnFav
                .blnLP = False
                .blnOKClicked = False
                .blnTP = False
                .blnUpdateAdult = False
                .blnUpdateDiff = False
                .blnUpdateFav = False
                .blnUpdateHighScore = False
                .blnUpdateLP = False
                .blnUpdateNotes = False
                .blnUpdateRating = False
                .blnUpdateTP = False
                .lngDiffID = g_udtGame.lngDiffID
                .lngRating = g_udtGame.lngRating
                .strHighScore = g_udtGame.strHighScore
                If .strHighScore = g_strNotEntered Then
                    .strHighScore = ""
                End If
                If g_udtGame.blnLoaded_3 = False Then
                    'load the memotext
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MemoText FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenSnapshot)
                    g_udtGame.strMemo = g_rsGames!MemoText
                    g_rsGames.Close
                    Set g_rsGames = Nothing
                    'validate it's length
                    If Len(g_udtGame.strMemo) > gc_lngMAXMEMOTEXTLENGTH Then
                        g_udtGame.strMemo = Left$(g_udtGame.strMemo, gc_lngMAXMEMOTEXTLENGTH)
                    End If
                End If
                .strNotes = g_udtGame.strMemo
            End With
            Screen.MousePointer = vbHourglass
            DoEvents
            frmPersonalProps.tabProps.Tab = IIf((blnAllGames_in), 0, 1)
            frmPersonalProps.Show vbModal, Me
            If g_udtEditPersonalInfo.blnOKClicked Then
                'SAVE CHANGES
                If g_udtEditPersonalInfo.blnEditAll = False Then
                    'just selected game - save to database
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !Adult = g_udtEditPersonalInfo.blnAdult
                            !Classic = g_udtEditPersonalInfo.blnClassic
                            !Fav = g_udtEditPersonalInfo.blnFav
                            If g_udtEditPersonalInfo.blnLP = True Then
                                !DateLastPlayed = ""
                            End If
                            If g_udtEditPersonalInfo.blnTP = True Then
                                !TimesPlayed = 0
                            End If
                            !DI_Id = g_udtEditPersonalInfo.lngDiffID
                            !Rating = g_udtEditPersonalInfo.lngRating
                            !HighScore = g_udtEditPersonalInfo.strHighScore
                            !MemoText = g_udtEditPersonalInfo.strNotes
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    're-click game
                    liSelected.Bold = (g_udtEditPersonalInfo.blnClassic Or g_udtEditPersonalInfo.blnFav)
                    Call SetItemInList(liSelected.Key, False)
                    Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                    Call UnCacheViews
                Else
                    'all games in view
                    Load frmBatchEdit
                    frmBatchEdit.m_blnUpdatePersonalInfo = True
                    frmBatchEdit.Show vbModal, Me
                    Call UnCacheViews
                    Call ClickLstView
                End If
             End If
             Screen.MousePointer = vbDefault
        End If
    End If

End Sub

Private Sub ShowExtrasEditor()

    '------------------------
    'Shows the extras manager
    '------------------------

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long
    Dim strError As String
    Dim lngEXidx As Long
    Dim lngIdx As Long

    strError = GetValidExtrasPath(lngIdx)
    If strError = "" Then
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        Set liSelected = lvMain(lngListViewIdx).SelectedItem
        If Not (liSelected Is Nothing) Then
            If g_eWhatIsSelected = e_SelType_IsGame Then
                Screen.MousePointer = vbHourglass
                DoEvents
                frmExtrasEditor.Show vbModal, Me
                If g_blnExtrasEditorOKPressed = True Then
                    If g_blnApplyClearExtrasToAll = False Then
                        'just re-click game - already saved new extras data
                        Call SetItemInList(liSelected.Key, False)
                    Else
                        'batch delete extras for all games in list!
                        Load frmBatchEdit
                        frmBatchEdit.m_blnUpdateExtras = True
                        frmBatchEdit.Show vbModal, Me
                        Call UnCacheViews
                        Call ClickLstView
                    End If
                End If
            End If
        End If
        Set liSelected = Nothing
        Screen.MousePointer = vbDefault
    Else
        MsgBox strError, vbInformation, gc_strAPPNAME
    End If

End Sub

Private Sub ShowGameSelector(ByVal blnNewGame_in As Boolean)

    Dim strError As String
    Dim lngGPIdx As Long
    Dim liSelected As ListItem
    Dim blnShowIt As Boolean
    Dim lngListViewIdx As Long
    Dim lngCCode As Long
    Dim lngIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If (Not (liSelected Is Nothing)) And (blnNewGame_in = False) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            'determine that a gamepath exists
            strError = GetValidGamePath()
            If strError <> "" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
                Set liSelected = Nothing
                Exit Sub
            End If
            'get filename
            g_blnGameSelectorNewGame = blnNewGame_in
            g_lngGameSelectorVersion = g_udtGame.lngVer
            g_blnGameSelectorApplyAllVersion = False
            blnShowIt = True
            If g_udtGame.strGameFilename = g_strNone Then
                g_strGameSelectorFilename = ""
                g_strGameSelectorIndex = 0
                g_strGameSelectorDefaultFile = ""
            Else
                'find actual game path\file
                MousePointer = vbHourglass
                DoEvents
                strError = FindGameFile(lngGPIdx, g_udtGame.strGameFilename)
                MousePointer = vbDefault
                DoEvents
                If strError <> "" Then
                    'not found or aborted
                    If strError <> "~~ABORTED~~" Then
                        'not found
                        g_strGameSelectorFilename = ""
                        g_strGameSelectorIndex = 0
                        g_strGameSelectorDefaultFile = ""
                        MsgBox strError, vbExclamation, gc_strAPPNAME
                        If GetExistingGamePath() < 0 Then
                            blnShowIt = False
                        End If
                    Else
                        'aborted
                        blnShowIt = False
                    End If
                Else
                    'found file
                    g_strGameSelectorFilename = ga_udtRegPathGame(lngGPIdx).strWithBackSlash & g_udtGame.strGameFilename
                    g_strGameSelectorIndex = g_udtGame.lngGameFilenameIndex
                    g_strGameSelectorDefaultFile = g_udtGame.strFileToRun
                End If
            End If
            If blnShowIt = True Then
                'show the selector
                frmGameSelector.Show vbModal, Me
                If g_blnGameSelectorOKPressed = True Then
                    If blnNewGame_in = False Then
                        If g_blnApplyGameFileToAll = True Then
                            Load frmBatchEdit
                            frmBatchEdit.m_blnUpdateGameFile = True
                            frmBatchEdit.Show vbModal, Me
                            Call UnCacheViews
                            Call ClickLstView
                        Else
                            'save back to database
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Filename, FilenameIndex, FileToRun, FA, CCode, Version FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .MoveFirst
                                    .Edit
                                    !FA = (g_strGameSelectorFilename <> "")
                                    !Filename = g_strGameSelectorFilename
                                    !FileToRun = g_strGameSelectorDefaultFile
                                    !FilenameIndex = g_strGameSelectorIndex
                                    !Version = g_lngGameSelectorVersion
                                    If g_strGameSelectorFilename <> "" Then
                                        If g_udtSid.blnFileExists = True Then
                                             lngCCode = 1 'both exist
                                        Else
                                            lngCCode = 2 'only game exists
                                        End If
                                    Else
                                        If g_udtSid.blnFileExists = True Then
                                            lngCCode = 3 'only sid exists
                                        Else
                                            lngCCode = 4 'none exist
                                        End If
                                    End If
                                    !CCode = lngCCode
                                    .Update
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                            'amend the icon
                            liSelected.Icon = lngCCode
                            liSelected.SmallIcon = lngCCode
                            'amend the extra columns
                            With ga_udtViewData(lstView.ListIndex)
                                For lngIdx = 0 To .lngExtraColumnsCount - 1
                                    With .a_udtExtraColumns(lngIdx)
                                        If .lngTypeIndex = 1 And .lngFieldIndex = 0 Then
                                            If lvMain(lngListViewIdx).ColumnHeaders.Count > lngIdx + 1 Then
                                                liSelected.SubItems(lngIdx + 1) = g_strGameSelectorFilename
                                            End If
                                        End If
                                    End With
                                Next lngIdx
                            End With
                            Call UnCacheViews
                            'reclick the game/sid
                            Call SetItemInList(liSelected.Key, False)
                        End If
                    Else
                        'add the game
                    End If
                End If
            End If
        End If
    Else
        'adding a new game
        If blnNewGame_in = True Then
            'determine that a gamepath exists
            strError = GetValidGamePath()
            If strError <> "" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
                Set liSelected = Nothing
                Exit Sub
            End If
            g_blnGameSelectorApplyAllVersion = False
            g_blnGameSelectorNewGame = True
            g_strGameSelectorFilename = ""
            g_strGameSelectorIndex = 0
            g_lngGameSelectorVersion = 1
            g_strGameSelectorDefaultFile = ""
            blnShowIt = True
            frmGameSelector.Show vbModal, Me
        End If
    End If
    Set liSelected = Nothing

End Sub

Private Sub SetGameLink()

    '----------------------------------------------------------
    'Set's the selected link field to the currently marked game
    '----------------------------------------------------------

    Dim liSelected As ListItem  'list item must be selected
    Dim blnExists As Boolean    'the game to link exists
    Dim lngNewLink As Long
    Dim lngOldRelatedLink As Long

    Set liSelected = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            If m_udtMarkedGame.strName <> "" And m_udtMarkedGame.lngID > 0 Then
                If m_udtMarkedGame.lngID <> g_udtGame.lngID Then
                    'make sure the marked game exists
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Name FROM Games WHERE GA_Id = " & m_udtMarkedGame.lngID, dbOpenSnapshot)
                    If g_rsGames.EOF = False Then
                        blnExists = True
                        g_rsGames.MoveFirst
                        g_strMarkedGame = g_rsGames!Name
                    Else
                        blnExists = False
                    End If
                    g_rsGames.Close
                    Set g_rsGames = Nothing
                    If blnExists = True Then
                        g_blnGameLinkSetFlag = True
                        g_strGameLinkField = g_strDefaultGameLinkField
                        frmGameLinkSelector.Show vbModal, Me
                        If g_strGameLinkField <> "" Then
                            'update database links
                            lngNewLink = m_udtMarkedGame.lngID
                            If g_strGameLinkField = "CloneOf" Then

                                'set link to new original (cloneof)
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT CloneOf FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .Edit
                                        !CloneOf = g_udtGame.lngID
                                        .Update
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing

                            ElseIf g_strGameLinkField = "Prequel" Then

                                'update Precursor link(s)
                                'obtain existing prequel link from selected game
                                lngOldRelatedLink = 0
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        lngOldRelatedLink = !Prequel
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'any games with sequels pointing to this game are now void so remove them
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE Sequel = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .MoveFirst
                                        Do Until .EOF
                                            .Edit
                                            !Sequel = 0
                                            .Update
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'any games with Precursors pointing to the new Precursor are now void so remove them
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE Prequel = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .MoveFirst
                                        Do Until .EOF
                                            .Edit
                                            !Prequel = 0
                                            .Update
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'set link to new Precursor
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .Edit
                                        !Prequel = lngNewLink
                                        .Update
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'set the Precursor game to point to the current game as it's sequel
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .Edit
                                        !Sequel = g_udtGame.lngID
                                        .Update
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                
                                If (lngOldRelatedLink > 0 And lngNewLink <> lngOldRelatedLink) Then
                                    'set the prequel game to point to the old linked prequel as it's prequel (insertion)
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Prequel = lngOldRelatedLink
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                    'set the old linked game to point to the prequel as it's sequel (insertion)
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & lngOldRelatedLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Sequel = lngNewLink
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                End If

                            ElseIf g_strGameLinkField = "Sequel" Then

                                'update Sequel link(s)
                                
                                'obtain existing sequel link from selected game
                                lngOldRelatedLink = 0
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        lngOldRelatedLink = !Sequel
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'any games with Precursors pointing to this game are now void so remove them
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE Prequel = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .MoveFirst
                                        Do Until .EOF
                                            .Edit
                                            !Prequel = 0
                                            .Update
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'any games with sequels pointing to the new sequel are now void so remove them
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE Sequel = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .MoveFirst
                                        Do Until .EOF
                                            .Edit
                                            !Sequel = 0
                                            .Update
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'set link to new Sequel
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .Edit
                                        !Sequel = lngNewLink
                                        .Update
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                'set the sequel game to point to the current game as it's Precursor
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        .Edit
                                        !Prequel = g_udtGame.lngID
                                        .Update
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                
                                If (lngOldRelatedLink > 0 And lngNewLink <> lngOldRelatedLink) Then
                                    'set the sequel game to point to the old linked sequel as it's sequel (insertion)
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Sequel = lngOldRelatedLink
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                    'set the old linked game to point to the sequel as it's Precursor (insertion)
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & lngOldRelatedLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Prequel = lngNewLink
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                End If

                            Else

                                'update Related link(s)
                                
                                Dim lngLoopCount As Long
                                
                                'check if selected game (g_udtGame) is in the list with marked game (lngNewLink)
                                lngOldRelatedLink = 0
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                With g_rsGames
                                    If Not .EOF Then
                                        lngOldRelatedLink = !Related
                                    End If
                                    .Close
                                End With
                                Set g_rsGames = Nothing
                                lngLoopCount = 0
                                Do While ((lngOldRelatedLink <> g_udtGame.lngID) And (lngOldRelatedLink <> lngNewLink) And (lngOldRelatedLink > 0) And (lngLoopCount < 1000))
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngOldRelatedLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            lngOldRelatedLink = !Related
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                    lngLoopCount = lngLoopCount + 1
                                Loop

                                '--------------------
                                If lngOldRelatedLink = g_udtGame.lngID Then

                                    'get related game from current game
                                    lngOldRelatedLink = 0
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            lngOldRelatedLink = !Related
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                
                                    'remove current game from linked list by setting any games related to current game to its related link
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE Related = " & g_udtGame.lngID, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .MoveFirst
                                            Do Until .EOF
                                                .Edit
                                                !Related = lngOldRelatedLink
                                                .Update
                                                .MoveNext
                                            Loop
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                
                                    'point marked game to the current game
                                    lngOldRelatedLink = 0
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            lngOldRelatedLink = !Related
                                            !Related = g_udtGame.lngID
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing

                                    'make circular if no previous link
                                    If lngOldRelatedLink = 0 Then
                                        lngOldRelatedLink = lngNewLink
                                    End If
                                
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Related = lngOldRelatedLink
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                
                                Else
                                
                                    Dim lngMarkedStart As Long
                                    Dim lngSelectedStart As Long
                                
                                    'set start of list for marked game (lngNewLink)
                                    lngMarkedStart = 0
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            lngMarkedStart = !Related
                                        End If
                                        .Close
                                    End With
                                    If lngMarkedStart = 0 Then
                                        lngMarkedStart = lngNewLink
                                    End If

                                    'find link in list related to selected game (g_udtGame)
                                    lngSelectedStart = g_udtGame.lngID
                                    lngOldRelatedLink = 0
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            lngOldRelatedLink = !Related
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                    lngLoopCount = 0
                                    Do While ((lngOldRelatedLink <> g_udtGame.lngID) And (lngOldRelatedLink > 0) And (lngLoopCount < 1000))
                                        lngSelectedStart = lngOldRelatedLink
                                        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngOldRelatedLink, dbOpenDynaset)
                                        With g_rsGames
                                            If Not .EOF Then
                                                lngOldRelatedLink = !Related
                                            End If
                                            .Close
                                        End With
                                        Set g_rsGames = Nothing
                                        lngLoopCount = lngLoopCount + 1
                                    Loop
                                    If (lngOldRelatedLink = 0) Or (lngLoopCount >= 1000) Then
                                        lngSelectedStart = g_udtGame.lngID
                                    End If

                                    'point marked game to selected game
                                    lngOldRelatedLink = 0
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngNewLink, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            lngOldRelatedLink = !Related
                                            !Related = g_udtGame.lngID
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing

                                    'set start of selected list to start of marked list
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & lngSelectedStart, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .Edit
                                            !Related = lngMarkedStart
                                            .Update
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing

                                End If
                                '--------------------
                            End If
                            
                            'reclick the game/sid
                            Call SetItemInList(liSelected.Key, False)
                        End If
                    End If
                End If
            End If
        End If
        Set liSelected = Nothing
    End If

End Sub
Private Sub ClearGameLink()

    '-----------------------------------------------------------------------
    'Clears the selected link field
    '
    'Inputs:    strFieldName_in:    The field name to change:
    '                               Either "Prequel", "Sequel", or "Related"
    '-----------------------------------------------------------------------

    Dim liSelected As ListItem  'list item must be selected
    Dim lngCurrentLinked As Long
    Dim blnExists As Boolean

    Set liSelected = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            With g_udtGame
                If (.lngPrequel > 0) Or (.lngSequel > 0) Or (.lngRelated > 0) Or (.lngCloneOf > 0) Then
                    g_blnGameLinkSetFlag = False
                    If .lngPrequel > 0 Then
                        g_strGameLinkField = "Prequel"
                    ElseIf .lngSequel > 0 Then
                        g_strGameLinkField = "Sequel"
                    ElseIf .lngRelated > 0 Then
                        g_strGameLinkField = "Related"
                    Else
                        g_strGameLinkField = "CloneOf"
                    End If
                    frmGameLinkSelector.Show vbModal, Me
                    If g_strGameLinkField <> "" Then
                        'clear database links
                        If g_strGameLinkField = "CloneOf" Then
                            'remove cloneof link
                            lngCurrentLinked = .lngCloneOf
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT CloneOf FROM Games WHERE GA_Id = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .Edit
                                    !CloneOf = 0
                                    .Update
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                        ElseIf g_strGameLinkField = "Related" Then
                            'remove related link
                            lngCurrentLinked = .lngRelated
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .Edit
                                    !Related = 0
                                    .Update
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                            'If g_blnMaintainGameLinks = True Then   'changed to always maintain circular links
                                'find related games pointing to this game and
                                'point it to the related game we just cleared
                                '(keep circular links)
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id FROM Games WHERE GA_Id = " & lngCurrentLinked, dbOpenSnapshot)
                                blnExists = Not g_rsGames.EOF
                                g_rsGames.Close
                                Set g_rsGames = Nothing
                                If blnExists = True Then
                                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Related FROM Games WHERE Related = " & .lngID, dbOpenDynaset)
                                    With g_rsGames
                                        If Not .EOF Then
                                            .MoveFirst
                                            Do Until .EOF
                                                'don't point to self
                                                .Edit
                                                If !GA_Id <> lngCurrentLinked Then
                                                    !Related = lngCurrentLinked
                                                Else
                                                    !Related = 0
                                                End If
                                                .Update
                                                .MoveNext
                                            Loop
                                        End If
                                        .Close
                                    End With
                                    Set g_rsGames = Nothing
                                End If
                            'End If   'changed to always maintain circular links
                        ElseIf g_strGameLinkField = "Prequel" Then
                            'remove prequel link
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .Edit
                                    !Prequel = 0
                                    .Update
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                            'remove all games with _Sequel_ link to current game
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE Sequel = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .MoveFirst
                                    Do Until .EOF
                                        .Edit
                                        !Sequel = 0
                                        .Update
                                        .MoveNext
                                    Loop
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                        Else
                            'remove sequel link
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .Edit
                                    !Sequel = 0
                                    .Update
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                            'remove all games with _Prequel_ link to current game
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE Prequel = " & .lngID, dbOpenDynaset)
                            With g_rsGames
                                If Not .EOF Then
                                    .MoveFirst
                                    Do Until .EOF
                                        .Edit
                                        !Prequel = 0
                                        .Update
                                        .MoveNext
                                    Loop
                                End If
                                .Close
                            End With
                            Set g_rsGames = Nothing
                        End If
                        'reclick the game/sid
                        Call SetItemInList(liSelected.Key, False)
                    End If
                End If
            End With
        End If
        Set liSelected = Nothing
    End If

End Sub



Private Sub ShowMusicianProps(strFocus_in As String)

    'let the user change the musician properties (not unknown or none)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If (g_eWhatIsSelected = e_SelType_IsSid) Or (g_eWhatIsSelected = e_SelType_IsGame) Then
            If (g_udtSid.lngMU_Id = 38) Or (g_udtSid.lngMU_Id = 39) Then
                MsgBox g_xobjFunc.GetINIMessage("Messages", "CannotChangeMusicianProps", "You cannot change the properties for this Musician!"), vbInformation, gc_strAPPNAME
            Else
                'ok to change
                With g_udtMusicianProps
                    .blnOKClicked = False
                    .blnAdding = False
                    .strGroup = IIf((g_udtSid.strGroup = g_strNone), "", g_udtSid.strGroup)
                    .strMusician = g_udtSid.strMusician
                    .strNick = IIf((g_udtSid.strNick = g_strNone), "", g_udtSid.strNick)
                    .strPhoto = IIf((g_udtSid.strPhoto = g_strNone), "", g_udtSid.strPhoto)
                    .strFocus = strFocus_in
                    frmEditMusicianProps.Show vbModal, Me
                    If .blnOKClicked = True Then
                        'save details
                        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Musicians WHERE MU_Id = " & g_udtSid.lngMU_Id, dbOpenDynaset)
                        With g_rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                !Musician = g_udtMusicianProps.strMusician
                                !Grp = g_udtMusicianProps.strGroup
                                !Nick = g_udtMusicianProps.strNick
                                !Photo = g_udtMusicianProps.strPhoto
                                .Update
                            End If
                            .Close
                        End With
                        Set g_rsGames = Nothing

                        're-click game/tune
                        Call SetItemInList(liSelected.Key, False)
                        Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                        Call UnCacheViews
                    End If
                End With
            End If
        End If
        Set liSelected = Nothing
    End If

End Sub

Private Sub ShowRecordSelector(eRecType As RecordSelectorEnum)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long
    Dim strSaveSQL As String
    Dim lngOldId As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            With g_udtRecordSelector
                .blnOKClicked = False
                .eRecordSelectorType = eRecType
                If .eRecordSelectorType = e_RecordSelector_Cracker Then
                    .strSQL = "SELECT CR_Id as ID, Cracker as Name FROM Crackers ORDER BY Cracker"
                    .lngID = g_udtGame.lngCrackerID
                    strSaveSQL = "SELECT CR_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID
                ElseIf .eRecordSelectorType = e_RecordSelector_Language Then
                    .strSQL = "SELECT LA_Id as ID, Language as Name FROM Languages ORDER BY Language"
                    .lngID = g_udtGame.lngLanguageID
                    strSaveSQL = "SELECT LA_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID
                ElseIf .eRecordSelectorType = e_RecordSelector_Musician Then
                    .strSQL = "SELECT MU_Id as ID, Musician as Name FROM Musicians ORDER BY Musician"
                    .lngID = g_udtSid.lngMU_Id
                    strSaveSQL = "SELECT MU_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID
                ElseIf .eRecordSelectorType = e_RecordSelector_Programmer Then
                    .strSQL = "SELECT PR_Id as ID, Programmer as Name FROM Programmers ORDER BY Programmer"
                    .lngID = g_udtGame.lngProgrammerID
                    strSaveSQL = "SELECT PR_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID

                ElseIf .eRecordSelectorType = e_RecordSelector_Artist Then
                    .strSQL = "SELECT AR_Id as ID, Artist as Name FROM Artists ORDER BY Artist"
                    .lngID = g_udtGame.lngArtistID
                    strSaveSQL = "SELECT AR_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID

                ElseIf .eRecordSelectorType = e_RecordSelector_Developer Then
                    .strSQL = "SELECT DE_Id as ID, Developer as Name FROM Developers ORDER BY Developer"
                    .lngID = g_udtGame.lngDeveloperID
                    strSaveSQL = "SELECT DE_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID

                ElseIf .eRecordSelectorType = e_RecordSelector_License Then
                    .strSQL = "SELECT LI_Id as ID, License as Name FROM Licenses ORDER BY License"
                    .lngID = g_udtGame.lngLicenseID
                    strSaveSQL = "SELECT LI_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID

                ElseIf .eRecordSelectorType = e_RecordSelector_Rarity Then
                    .strSQL = "SELECT RA_Id as ID, Rarity as Name FROM Rarities ORDER BY Rarity"
                    .lngID = g_udtGame.lngRarityID
                    strSaveSQL = "SELECT RA_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID

                Else 'publisher
                    .strSQL = "SELECT PU_Id as ID, Publisher as Name FROM Publishers ORDER BY Publisher"
                    .lngID = g_udtGame.lngPublisherID
                    strSaveSQL = "SELECT PU_Id AS ID FROM Games WHERE GA_Id = " & g_udtGame.lngID
                End If
                lngOldId = .lngID
                Screen.MousePointer = vbHourglass
                DoEvents
                frmRecordSelector.Show vbModal, Me
                If .blnOKClicked Then
                    'save field
                    If .lngID <> lngOldId Then
                        Set g_rsGames = g_dbGamebase.OpenRecordset(strSaveSQL, dbOpenDynaset)
                        With g_rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                !ID = g_udtRecordSelector.lngID
                                .Update
                            End If
                            .Close
                        End With
                        Set g_rsGames = Nothing
                    End If
                    're-click game
                    Call SetItemInList(liSelected.Key, False)
                    Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
                    Call UnCacheViews
                End If
                Screen.MousePointer = vbDefault
            End With
        End If
    End If

End Sub

Private Sub ShowScrnShotSelector()

    '------------------------------------------------------------------------
    'Shows the screenshot selector and returns the chosen screenshot filename
    '------------------------------------------------------------------------

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long
    Dim strError As String
    Dim lngSSIdx As Long
    Dim lngIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            Screen.MousePointer = vbHourglass
            DoEvents
            'are we allowed to show the selector?
            strError = GetValidScrnShotPath(lngSSIdx)
            If strError <> "" Then
                'no screenshot path exists
                MsgBox strError, vbExclamation, gc_strAPPNAME
            Else

                'store a valid path to open if we can't find the file
                g_strScrnShotSelectorOpeningPath = ga_udtRegPathScrnShot(lngSSIdx).strRealPath
                
                'get the filename to send to the selector
                If g_udtGame.strScreenshotFilename = g_strNone Then
                    g_strScrnShotSelectorFilename = ""
                    g_strScrnShotSelectorDBFilename = ""
                Else
                    g_strScrnShotSelectorDBFilename = g_udtGame.strScreenshotFilename
                    strError = FindScrnShotFile(lngSSIdx, g_strScrnShotSelectorDBFilename)
                    If strError <> "" Then
                        MsgBox strError, vbExclamation, gc_strAPPNAME
                        g_strScrnShotSelectorFilename = ""
                    Else
                        g_strScrnShotSelectorFilename = ga_udtRegPathScrnShot(lngSSIdx).strWithBackSlash & g_udtGame.strScreenshotFilename
                        g_strScrnShotSelectorOpeningPath = ga_udtRegPathScrnShot(lngSSIdx).strRealPath
                    End If
                End If
            
                'show the selector
                frmScrnShotSelector.Show vbModal, Me
                Screen.MousePointer = vbDefault
                
                'just save the new filename if changed
                If g_blnScrnShotSelectorOKPressed = True Then
                    If g_blnApplyScrnShotFileToAll = True Then
                    
                        'change all games in current view to point to the returned scrnshot file
                        Load frmBatchEdit
                        frmBatchEdit.m_blnUpdatePictureFile = True
                        frmBatchEdit.Show vbModal, Me
                        Call UnCacheViews
                        Call ClickLstView
                    Else
                        
                        'just save the new filename if changed
                        If UCase$(g_strScrnShotSelectorFilename) <> UCase$(g_udtGame.strScreenshotFilename) Then
                            
                            'screenshot filename changed - save to database
                            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ScrnShotFilename FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                            With g_rsGames
                                .MoveFirst
                                .Edit
                                !ScrnShotFilename = g_strScrnShotSelectorFilename
                                .Update
                                .Close
                            End With
                            Set g_rsGames = Nothing

                            'amend the extra columns
                            With ga_udtViewData(lstView.ListIndex)
                                For lngIdx = 0 To .lngExtraColumnsCount - 1
                                    With .a_udtExtraColumns(lngIdx)
                                        If .lngTypeIndex = 1 And .lngFieldIndex = 17 Then
                                            If lvMain(lngListViewIdx).ColumnHeaders.Count > lngIdx + 1 Then
                                                liSelected.SubItems(lngIdx + 1) = IIf((g_strScrnShotSelectorFilename = g_strNone), "", g_strScrnShotSelectorFilename)
                                            End If
                                        End If
                                    End With
                                Next lngIdx
                            End With

                            'reclick game
                            Call SetItemInList(liSelected.Key, False)

                        End If
                        
                    End If
                End If
            
            End If
            
        End If
    End If
    Set liSelected = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub ShowSidSelector(ByVal blnNewSid_in As Boolean)

    '----------------------------------------------------
    'Shows the sid selector and returns a chosen sid file
    '----------------------------------------------------

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long
    Dim lngCCode As Long
    Dim a_strTypes() As String
    Dim lngIdx As Long

    If GetSingleMusicTypesArray(a_strTypes()) = 0 Then
        MsgBox g_xobjFunc.GetINIMessage("Error_Messages", "Error_NoMusicTypes", "No valid music player supported file-types have been set! Set them in Options!"), vbExclamation, gc_strAPPNAME
        Exit Sub
    End If

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If (g_eWhatIsSelected = e_SelType_IsGame) Or (g_eWhatIsSelected = e_SelType_IsSid) Then
            'are we allowed to see the selector?
            If g_xobjFunc.DirExists(g_udtRegPathSid.strRealPath) Then
                'new game/sid?
                g_blnSidSelectorNewItem = blnNewSid_in
                'set the sid filename for the selector
                If blnNewSid_in = False Then
                    g_strSidSelectorFilename = g_udtSid.strFilename
                Else
                    g_strSidSelectorFilename = g_strNone
                End If
                'show the selector
                frmSidSelector.Show vbModal, Me
                'set the new filename if changed
                If blnNewSid_in = False Then
                    If g_blnApplySidFileToAll = True Then
                        If UCase(g_strSidSelectorFilename) <> UCase(g_udtSid.strFilename) Then
                            Load frmBatchEdit
                            frmBatchEdit.m_blnUpdateSidFile = True
                            frmBatchEdit.Show vbModal, Me
                            Call UnCacheViews
                            Call ClickLstView
                        End If
                    Else
                        If UCase(g_strSidSelectorFilename) <> UCase(g_udtSid.strFilename) Then
                            'update the database and memory storage
                            If g_eWhatIsSelected = e_SelType_IsGame Then
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT SidFilename as Filename, SA, FA, CCode FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                                With g_rsGames
                                    .Edit
                                    !Filename = IIf((g_strSidSelectorFilename = g_strNone), "", g_strSidSelectorFilename)
                                    If g_strSidSelectorFilename <> g_strNone Then
                                        !SA = True
                                        If !FA = True Then
                                             lngCCode = 1 'both exist
                                        Else
                                            lngCCode = 3 'only sid exists
                                        End If
                                    Else
                                        !SA = False
                                        If !FA = True Then
                                            lngCCode = 2 'only game exists
                                        Else
                                            lngCCode = 4 'none exist
                                        End If
                                    End If
                                    !CCode = lngCCode
                                    .Update
                                    .Close
                                End With
                            Else
                                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Filename, SA FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
                                With g_rsGames
                                    .Edit
                                    !Filename = IIf((g_strSidSelectorFilename = g_strNone), "", g_strSidSelectorFilename)
                                    !SA = (g_strSidSelectorFilename <> g_strNone)
                                    .Update
                                    .Close
                                End With
                            End If
                            Set g_rsGames = Nothing
                            g_udtSid.strFilename = g_strSidSelectorFilename
                            If g_eWhatIsSelected = e_SelType_IsGame Then
                                'amend the icon
                                liSelected.Icon = lngCCode
                                liSelected.SmallIcon = lngCCode
                                Call UnCacheViews
                            End If
                            'amend the extra columns
                            With ga_udtViewData(lstView.ListIndex)
                                For lngIdx = 0 To .lngExtraColumnsCount - 1
                                    With .a_udtExtraColumns(lngIdx)
                                        If .lngTypeIndex = 1 And .lngFieldIndex = 16 Then
                                            If lvMain(lngListViewIdx).ColumnHeaders.Count > lngIdx + 1 Then
                                                liSelected.SubItems(lngIdx + 1) = IIf((g_strSidSelectorFilename = g_strNone), "", g_strSidSelectorFilename)
                                            End If
                                        End If
                                    End With
                                Next lngIdx
                            End With
                            'reclick the game/sid
                            Call SetItemInList(liSelected.Key, False)
                        End If
                    End If
                End If
            Else
                'no sid path found
                MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_MusicPathNotFound", "Your Music Path does not exist or was not found."), vbInformation, gc_strAPPNAME
            End If
        End If
    End If
    Set liSelected = Nothing
    
End Sub


Private Sub ClickLstView()

    Dim A As Long
    
    With lstView
        A = .ListIndex
        .ListIndex = gc_lngNOLISTINDEX
        .ListIndex = A
    End With

End Sub

Public Sub CreateLookupTables()

    '----------------------------------------------------
    'Creates the various lookup tables to speed things up
    '----------------------------------------------------
    
    Dim lngYearIdx As Long          'loop index
    Dim fldParentGenre As Field     'field data storage
    Dim fldPG_ID As Field           'Field data storage
    Dim lngIdx As Long              'loop index

    Const TAB0_COL1_LABEL_X As Single = 240
    Const TAB0_COL1_FIELD_X As Single = 1680
    Const TAB0_COL2_LABEL_X As Single = 5760
    Const TAB0_COL2_FIELD_X As Single = 7155
    
    Const TAB1_COL1_LABEL_X As Single = 960
    Const TAB1_COL1_FIELD_X As Single = 2430
    Const TAB1_COL2_LABEL_X As Single = 6960
    Const TAB1_COL2_FIELD_X As Single = 8355

    Const TAB4_COL1_LABEL_X As Single = 240
    Const TAB4_COL1_FIELD_X As Single = 1440

    Const TAB0_Y_START As Single = 0
    Const TAB0_Y_INC As Single = 240
    Const TAB1_Y_START As Single = 0
    Const TAB1_Y_INC As Single = 240
    Const TAB4_Y_START As Single = 0
    Const TAB4_Y_INC As Single = 240
    
    'sql engine
    ga_strViewModesLookupTable(0) = " AND "
    ga_strViewModesLookupTable(1) = " OR "
        
    'no/yes strings
    ga_strYesNoLookupTable(0) = g_xobjFunc.GetINIMessage("Database", "Field_No", "No")
    ga_strYesNoLookupTable(1) = g_xobjFunc.GetINIMessage("Database", "Field_Yes", "Yes")
    
    ga_strYesNoUnknownLookupTable(0) = g_xobjFunc.GetINIMessage("Database", "Field_No", "No")
    ga_strYesNoUnknownLookupTable(1) = g_xobjFunc.GetINIMessage("Database", "Field_Yes", "Yes")
    ga_strYesNoUnknownLookupTable(2) = g_xobjFunc.GetINIMessage("Database", "Field_Unknown", "(Unknown)")

    'PAL/NTSC
    ga_strPalNTSCLookupTable(0) = GetINIMessage("Database", "Field_PALNTSCType0", "PAL")
    ga_strPalNTSCLookupTable(1) = GetINIMessage("Database", "Field_PALNTSCType1", "PAL+NTSC")
    ga_strPalNTSCLookupTable(2) = GetINIMessage("Database", "Field_PALNTSCType2", "NTSC")
    ga_strPalNTSCLookupTable(3) = GetINIMessage("Database", "Field_PALNTSCType3", "PAL(+NTSC?)")

    'game length type
    ga_strLengthTypeLookupTable(0) = GetINIMessage("Database", "Field_LengthType0", "Block(s)")
    ga_strLengthTypeLookupTable(1) = GetINIMessage("Database", "Field_LengthType1", "Disk(s)")
    ga_strLengthTypeLookupTable(2) = GetINIMessage("Database", "Field_LengthType2", "Cartridge(s)")
    ga_strLengthTypeLookupTable(3) = GetINIMessage("Database", "Field_LengthType3", "Tape(s)")
    ga_strLengthTypeLookupTable(19) = GetINIMessage("Database", "Field_LengthType19", "Unknown")
    For lngIdx = 4 To 18
        ga_strLengthTypeLookupTable(lngIdx) = GetINIMessage("Database", "Field_LengthType" & Trim$(lngIdx), "Unused")
    Next lngIdx

    'years field
    ga_strYearLookupTable(1) = "????"
    ga_strYearLookupTable(2) = "19??"
    ga_strYearLookupTable(3) = "197?"
    ga_strYearLookupTable(4) = "198?"
    ga_strYearLookupTable(5) = "199?"
    ga_strYearLookupTable(6) = "20??"
    ga_strYearLookupTable(7) = "200?"
    ga_strYearLookupTable(8) = "201?"
    ga_strYearLookupTable(9) = "202?"
    For lngYearIdx = 10 To gc_lngYEARLOOKUPMAX
        ga_strYearLookupTable(lngYearIdx) = (gc_lngEARLIESTYEAR - 10) + lngYearIdx
    Next lngYearIdx

    With g_xobjFunc
        'players simulaneous field
        ga_strPlayerSimLookupTable(0) = ""
        ga_strPlayerSimLookupTable(1) = " (" & .GetINIMessage("Database", "Field_Players_Sumultaneous", "Simultaneous") & ")"
    
        'rating field
        ga_strRatingLookupTable(0) = .GetINIMessage("Database", "Field_Rating_Unknown", "?/5 - Unknown")
        ga_strRatingLookupTable(1) = .GetINIMessage("Database", "Field_Rating_Terrible", "1/5 - Terrible!")
        ga_strRatingLookupTable(2) = .GetINIMessage("Database", "Field_Rating_Poor", "2/5 - Poor")
        ga_strRatingLookupTable(3) = .GetINIMessage("Database", "Field_Rating_Average", "3/5 - Average")
        ga_strRatingLookupTable(4) = .GetINIMessage("Database", "Field_Rating_QuiteGood", "4/5 - Quite Good")
        ga_strRatingLookupTable(5) = .GetINIMessage("Database", "Field_Rating_VeryGood", "5/5 - Very Good")
        ga_strRatingLookupTable(6) = .GetINIMessage("Database", "Field_Rating_Classic", "Classic!")

        'control field
        ga_strControlLookupTable(0) = GetINIMessage("Database", "Field_Control_JoyPort2", "Joystick Port 2")
        ga_strControlLookupTable(1) = GetINIMessage("Database", "Field_Control_JoyPort1", "Joystick Port 1")
        ga_strControlLookupTable(2) = GetINIMessage("Database", "Field_Control_Keyboard", "Keyboard")
        ga_strControlLookupTable(3) = GetINIMessage("Database", "Field_Control_PaddlePort2", "Paddle Port 2")
        ga_strControlLookupTable(4) = GetINIMessage("Database", "Field_Control_PaddlePort1", "Paddle Port 1")
        ga_strControlLookupTable(5) = GetINIMessage("Database", "Field_Control_Mouse", "Mouse")
        ga_strControlLookupTable(6) = GetINIMessage("Database", "Field_Control_Lightpen", "Light Pen")
        ga_strControlLookupTable(7) = GetINIMessage("Database", "Field_Control_KoalaPad", "Koala Pad")
        ga_strControlLookupTable(8) = GetINIMessage("Database", "Field_Control_LightGun", "Light Gun")
    End With
    
    'genre/parent-genre field
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ParentGenre, PG_Id FROM PGenres ORDER BY PG_Id Desc", dbOpenSnapshot)
    With g_rsGames
        Set fldParentGenre = !ParentGenre
        Set fldPG_ID = !PG_Id
        .MoveFirst
        ReDim ga_strGenreParentLookupTable(fldPG_ID + 1)
        Do Until .EOF
            If fldParentGenre <> "" Then
                ga_strGenreParentLookupTable(fldPG_ID) = fldParentGenre & " - "
            Else
                ga_strGenreParentLookupTable(fldPG_ID) = fldParentGenre
            End If
            .MoveNext
        Loop
        Set fldParentGenre = Nothing
        Set fldPG_ID = Nothing
        .Close
    End With
    Set g_rsGames = Nothing

    'game name tweaks
    ga_udtGameNameLookupTable(0).strCommaText = ", The"
    ga_udtGameNameLookupTable(0).strReplaceText = "The "
    ga_udtGameNameLookupTable(0).lngLenCommaText = 5
    ga_udtGameNameLookupTable(0).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(1).strCommaText = ", Die"
    ga_udtGameNameLookupTable(1).strReplaceText = "Die "
    ga_udtGameNameLookupTable(1).lngLenCommaText = 5
    ga_udtGameNameLookupTable(1).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(2).strCommaText = ", Der"
    ga_udtGameNameLookupTable(2).strReplaceText = "Der "
    ga_udtGameNameLookupTable(2).lngLenCommaText = 5
    ga_udtGameNameLookupTable(2).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(3).strCommaText = ", Das"
    ga_udtGameNameLookupTable(3).strReplaceText = "Das "
    ga_udtGameNameLookupTable(3).lngLenCommaText = 5
    ga_udtGameNameLookupTable(3).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(4).strCommaText = ", Los"
    ga_udtGameNameLookupTable(4).strReplaceText = "Los "
    ga_udtGameNameLookupTable(4).lngLenCommaText = 5
    ga_udtGameNameLookupTable(4).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(5).strCommaText = ", Een"
    ga_udtGameNameLookupTable(5).strReplaceText = "Een "
    ga_udtGameNameLookupTable(5).lngLenCommaText = 5
    ga_udtGameNameLookupTable(5).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(6).strCommaText = ", De"
    ga_udtGameNameLookupTable(6).strReplaceText = "De "
    ga_udtGameNameLookupTable(6).lngLenCommaText = 4
    ga_udtGameNameLookupTable(6).lngLenReplaceText = 3
    
    ga_udtGameNameLookupTable(7).strCommaText = ", Le"
    ga_udtGameNameLookupTable(7).strReplaceText = "Le "
    ga_udtGameNameLookupTable(7).lngLenCommaText = 4
    ga_udtGameNameLookupTable(7).lngLenReplaceText = 3
    
    ga_udtGameNameLookupTable(8).strCommaText = ", La"
    ga_udtGameNameLookupTable(8).strReplaceText = "La "
    ga_udtGameNameLookupTable(8).lngLenCommaText = 4
    ga_udtGameNameLookupTable(8).lngLenReplaceText = 3
    
    ga_udtGameNameLookupTable(9).strCommaText = ", A"
    ga_udtGameNameLookupTable(9).strReplaceText = "A "
    ga_udtGameNameLookupTable(9).lngLenCommaText = 3
    ga_udtGameNameLookupTable(9).lngLenReplaceText = 2
    
    ga_udtGameNameLookupTable(10).strCommaText = ", A - "
    ga_udtGameNameLookupTable(10).strReplaceText = "A - "
    ga_udtGameNameLookupTable(10).lngLenCommaText = 6
    ga_udtGameNameLookupTable(10).lngLenReplaceText = 4
    
    ga_udtGameNameLookupTable(11).strCommaText = ", Dien"
    ga_udtGameNameLookupTable(11).strReplaceText = "Dien "
    ga_udtGameNameLookupTable(11).lngLenCommaText = 6
    ga_udtGameNameLookupTable(11).lngLenReplaceText = 5
    
    ga_udtGameNameLookupTable(12).strCommaText = ", The - "
    ga_udtGameNameLookupTable(12).strReplaceText = "The - "
    ga_udtGameNameLookupTable(12).lngLenCommaText = 8
    ga_udtGameNameLookupTable(12).lngLenReplaceText = 6

    '''''''''''''''''''''''''''''
    'Details Tabs Label Positions
    '''''''''''''''''''''''''''''

    g_lngTotalDetailsTabPositions = 33
    ReDim ga_udtDetailsTabPositionsLookup(g_lngTotalDetailsTabPositions)

    'tab 0 (game info)
    For lngIdx = 0 To 5
        With ga_udtDetailsTabPositionsLookup(lngIdx)
            .sngLabelTwipsX = TAB0_COL1_LABEL_X
            .sngFieldTwipsX = TAB0_COL1_FIELD_X
            .sngTwipsY = TAB0_Y_START + (TAB0_Y_INC * lngIdx)
        End With
    Next lngIdx
    For lngIdx = 6 To 11
        With ga_udtDetailsTabPositionsLookup(lngIdx)
            .sngLabelTwipsX = TAB0_COL2_LABEL_X
            .sngFieldTwipsX = TAB0_COL2_FIELD_X
            .sngTwipsY = TAB0_Y_START + (TAB0_Y_INC * (lngIdx - 6))
        End With
    Next lngIdx

    'tab 1 (version info)
    For lngIdx = 12 To 17
        With ga_udtDetailsTabPositionsLookup(lngIdx)
            .sngLabelTwipsX = TAB1_COL1_LABEL_X
            .sngFieldTwipsX = TAB1_COL1_FIELD_X
            .sngTwipsY = TAB1_Y_START + (TAB1_Y_INC * (lngIdx - 12))
        End With
    Next lngIdx
    For lngIdx = 18 To 23
        With ga_udtDetailsTabPositionsLookup(lngIdx)
            .sngLabelTwipsX = TAB1_COL2_LABEL_X
            .sngFieldTwipsX = TAB1_COL2_FIELD_X
            .sngTwipsY = TAB1_Y_START + (TAB1_Y_INC * (lngIdx - 18))
        End With
    Next lngIdx

    'tab 2 (personal info
    With ga_udtDetailsTabPositionsLookup(24)
        .sngLabelTwipsX = 1560
        .sngFieldTwipsX = 2760
        .sngTwipsY = 240
    End With
    With ga_udtDetailsTabPositionsLookup(25)
        .sngLabelTwipsX = 1560
        .sngFieldTwipsX = 2760
        .sngTwipsY = 600
    End With
    With ga_udtDetailsTabPositionsLookup(26)
        .sngLabelTwipsX = 6360
        .sngFieldTwipsX = 7560
        .sngTwipsY = 240
    End With
    With ga_udtDetailsTabPositionsLookup(27)
        .sngLabelTwipsX = 6360
        .sngFieldTwipsX = 7560
        .sngTwipsY = 600
    End With
    With ga_udtDetailsTabPositionsLookup(28)
        .sngLabelTwipsX = 4215
        .sngFieldTwipsX = 5160
        .sngTwipsY = 1100
    End With

    'tab 4 (version info) and music only tab
    For lngIdx = 29 To 32
        With ga_udtDetailsTabPositionsLookup(lngIdx)
            .sngLabelTwipsX = TAB4_COL1_LABEL_X
            .sngFieldTwipsX = TAB4_COL1_FIELD_X
            .sngTwipsY = TAB4_Y_START + (TAB4_Y_INC * (lngIdx - 29))
        End With
    Next lngIdx

End Sub

Private Function DeleteCustomView(ByVal lngIdxToDelete_in As Long) As Boolean

    '------------------------------------------------------------
    'Deletes a custom view from the database and from the global
    'custom views array.
    '
    'Inputs:        lngIdxToDelete_in   array index to delete
    '
    'Returns:       True if the Custom View was deleted ok
    '               False otherwise
    '
    'Remarks:       Will not delete default views.
    '               i.e. index must be over gc_lngDEFAULTVIEWSMAX
    '------------------------------------------------------------
    
    Dim lngViewID As Long       'Database View ID to delete
    Dim lngViewIdx As Long      'loop index
    Dim lngViewMax As Long      'storage for UBound() of views array
    
    'assertions
    lngViewMax = UBound(ga_udtViewData)
    If (lngIdxToDelete_in <= gc_lngDEFAULTVIEWSMAX) Or (lngIdxToDelete_in > lngViewMax) Then
        DeleteCustomView = False
        Exit Function
    End If

    'delete from the database or registry
    lngViewID = ga_udtViewData(lngIdxToDelete_in).lngID
    If (m_blnLastViewIsQuick) And (lngIdxToDelete_in = lngViewMax) Then
        'quick view - just set the number of filters in the quickview to 0
        'then the quick filter engine ignores it
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "NumFilters", 0, False
    Else
        'normal custom view
        g_dbGamebase.Execute "DELETE * FROM ViewData WHERE ID = " & lngViewID
        g_dbGamebase.Execute "DELETE * FROM ViewFilters WHERE ID = " & lngViewID
    End If

    'delete from the array by shuffling down over deleted item
    For lngViewIdx = (lngIdxToDelete_in + 1) To lngViewMax
        ga_udtViewData(lngViewIdx - 1) = ga_udtViewData(lngViewIdx)
    Next lngViewIdx
    ReDim Preserve ga_udtViewData(lngViewMax - 1)
    
    'change the Quick Filter View flag if we deleted the quick filter view
    If (m_blnLastViewIsQuick) And (lngIdxToDelete_in = lngViewMax) Then
        m_blnLastViewIsQuick = False
    End If
    
    'View Deleted!
    DeleteCustomView = True

End Function
Public Function FormatGameName(strName_in As String) As String
    
    'format a game name for display in the details tabs
    'from the name in the database/game-list

    Dim lngIdx As Long
    Dim blnDone As Boolean
    Dim lngPos As Long
    Dim lngLenName As Long
    Dim lngNextCharIdx As Long
    Dim lngAscChar As Long
    Dim blnSkip As Boolean

    blnDone = False
    lngLenName = Len(strName_in)
    For lngIdx = 0 To gc_lngGAMENAMELOOKUPMAX
        With ga_udtGameNameLookupTable(lngIdx)
            lngPos = InStr(1, strName_in, .strCommaText, vbTextCompare)
            If lngPos > 0 Then
                blnSkip = False
                lngNextCharIdx = lngPos + .lngLenCommaText
                If lngLenName >= lngNextCharIdx Then
                    lngAscChar = Asc(Mid$(strName_in, lngNextCharIdx, 1))
                    blnSkip = ((lngAscChar > 64) And (lngAscChar < 91)) Or ((lngAscChar > 96) And (lngAscChar < 123))
                End If
                If Not blnSkip Then
                    'commatext is either at the end of the string, or not immediately followed by a letter
                    FormatGameName = .strReplaceText & Left$(strName_in, lngPos - 1) & Right$(strName_in, lngLenName - lngNextCharIdx + 1)
                    blnDone = True
                    Exit For
                End If
            End If
        End With
    Next lngIdx

    If Not blnDone Then
        FormatGameName = strName_in
    End If

End Function
Private Function FormattedGameName(strName_in As String) As String

    'get the desired database game-name format for a game-name typed by the user

    Dim lngIdx As Long
    Dim blnFound As Boolean
    Dim lngPos As Long
    Dim lngLenName As Long

    blnFound = False
    lngLenName = Len(strName_in)
    For lngIdx = 0 To gc_lngGAMENAMELOOKUPMAX
        With ga_udtGameNameLookupTable(lngIdx)
            If InStr(1, strName_in, .strReplaceText, vbTextCompare) = 1 Then
                lngPos = InStr(.lngLenReplaceText, strName_in, " - ")
                If lngPos > 0 Then
                    'insert comma text before " - "
                    FormattedGameName = Left$(strName_in, lngPos - 1) & .strCommaText & Right$(strName_in, lngLenName - (lngPos - 1))
                    FormattedGameName = Right$(FormattedGameName, Len(FormattedGameName) - .lngLenReplaceText)
                Else
                    lngPos = InStr(.lngLenReplaceText, strName_in, " [")
                    If lngPos > 0 Then
                        'insert comma text before " ["
                        FormattedGameName = Left$(strName_in, lngPos - 1) & .strCommaText & Right$(strName_in, lngLenName - (lngPos - 1))
                        FormattedGameName = Right$(FormattedGameName, Len(FormattedGameName) - .lngLenReplaceText)
                    Else
                        'append comma text
                        FormattedGameName = Right$(strName_in, lngLenName - .lngLenReplaceText) & .strCommaText
                    End If
                End If
                blnFound = True
                Exit For
            End If
        End With
    Next lngIdx

    If Not blnFound Then
        FormattedGameName = strName_in
    End If

End Function

Private Sub GetINIMessages()

    '---------------------------------------------------------------------------------------------
    'read in all the messages from GBText.ini and Text.ini and set the controls on the main window
    '---------------------------------------------------------------------------------------------

    Dim lngIdx As Long      'loop index

    With g_xobjFunc
        'general
        g_strUnknown = .GetINIMessage("General", "Unknown", "Unknown")
        g_strNone = .GetINIMessage("General", "None", "None")
        g_strNotEntered = .GetINIMessage("General", "NotEntered", "Not Entered")
        g_strNoneSelected = .GetINIMessage("General", "NoneSelected", "None Selected")
        g_strNever = .GetINIMessage("General", "Never", "Never")
        g_strPleaseWait = .GetINIMessage("General", "PleaseWait", "Please Wait") & "..."
        g_strMarked = .GetINIMessage("General", "Marked", "Marked")

        'global buttons
        g_strButtonOKText = .GetINIMessage("General Buttons", "OK", "&OK")
        g_strButtonCancelText = .GetINIMessage("General Buttons", "Cancel", "&Cancel")
        g_strButtonCloseText = .GetINIMessage("General Buttons", "Close", "&Close")
        g_strButtonClearText = .GetINIMessage("General Buttons", "Clear", "C&lear")
        g_strButtonBrowseText = .GetINIMessage("General Buttons", "Browse", "&Browse")
        g_strButtonDefaultText = .GetINIMessage("General Buttons", "Default", "&Default")
        g_strButtonHelpText = .GetINIMessage("General Buttons", "Help", "&Help")

        'buttons
        cmdPlay.Caption = GetINIMessage("MainWindow", "Button_PlayGame", "&Play Game")
        cmdSid.Caption = GetINIMessage("MainWindow", "Button_PlayMusic", "Play &Music")
        cmdExtra.Caption = GetINIMessage("MainWindow", "Button_Extras", "E&xtras")
        cmdRandom.Caption = GetINIMessage("MainWindow", "Button_RandomSelect", "?")

        'tool-tip text
        cmdPlay.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Button_PlayGame", "Right-click to choose emulator")
        cmdSid.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Button_PlayMusic", "Right-click to choose player")
        cmdExtra.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Button_Extras", "View/Run Selected Extra")
        cmdRandom.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Button_RandomSelect", "Random Select")
        cmdCustViews.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Button_CustomViews", "Custom Views")
        imgSSUp.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Pictures_Prev", "Previous Picture")
        imgSSDown.ToolTipText = GetINIMessage("MainWindow", "MouseOver_Pictures_Next", "Next Picture")
    End With

    With sbStatus
        .Panels(1).ToolTipText = GetINIMessage("MainWindow", "MouseOver_StatusBar_GameFile", "Game File - Click to change")
        .Panels(2).ToolTipText = GetINIMessage("MainWindow", "MouseOver_StatusBar_MusicFile", "Music File - Click to change")
        .Panels(3).Text = GetINIMessage("MainWindow", "Status_Bar_Extras", "Extras") & "..."
        .Panels(3).ToolTipText = GetINIMessage("MainWindow", "MouseOver_StatusBar_Extras", "Game Extras - Click to change")
        .Panels(5).ToolTipText = GetINIMessage("MainWindow", "MouseOver_StatusBar_Sounds", "Sounds toggle - Click to change")
        .Panels(6).ToolTipText = GetINIMessage("MainWindow", "MouseOver_StatusBar_NonStopMusic", "Non Stop Music toggle - Click to change")
    End With

    With g_xobjFunc
        imgCup.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_ClassicGame", "Classic Game!")
        imgPrequel.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_Precursor", "Goto Precursor")
        imgRelated.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_Related", "Goto Related Game")
        imgSequel.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_Sequel", "Goto Sequel")
        imgCloneOf.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_CloneOf", "Clone of")
        m_strPrequelTTT = imgPrequel.ToolTipText
        m_strSequelTTT = imgSequel.ToolTipText
        m_strRelatedTTT = imgRelated.ToolTipText
        m_strCloneofTTT = imgCloneOf.ToolTipText
        imgAdult.ToolTipText = GetINIMessage("MainWindow", "MouseOver_TabIcons_AdultGame", "Adult Game")

        'views
        ga_strDefaultViewNames(0) = GetINIMessage("MainWindow", "Views_Default_AllGames", "All Games")
        ga_strDefaultViewNames(1) = GetINIMessage("MainWindow", "Views_Default_AllMusic", "All Music")
        ga_strDefaultViewNames(2) = GetINIMessage("MainWindow", "Views_Default_Favourites", "Favourite Games/Music")
        ga_strDefaultViewNames(3) = GetINIMessage("MainWindow", "Views_Default_NoClones", "All Games Excluding Clones")
        g_strViewNone = GetINIMessage("MainWindow", "Views_None", "None")

        'quick filter
        ga_QuickFilterNames(0) = GetINIMessage("Quick Filter", "Name", "Name")
        ga_QuickFilterNames(1) = GetINIMessage("Quick Filter", "Publisher", "Publisher")
        ga_QuickFilterNames(2) = GetINIMessage("Quick Filter", "Developer", "Developer")
        ga_QuickFilterNames(3) = GetINIMessage("Quick Filter", "Programmer", "Programmer")
        ga_QuickFilterNames(4) = GetINIMessage("Quick Filter", "Musician", "Musician")
        ga_QuickFilterNames(5) = GetINIMessage("Quick Filter", "Artist", "Artist")
        ga_QuickFilterNames(6) = GetINIMessage("Quick Filter", "License", "License")
        ga_QuickFilterNames(7) = GetINIMessage("Quick Filter", "Genre", "Genre")
        ga_QuickFilterNames(8) = GetINIMessage("Quick Filter", "Language", "Language")
        ga_QuickFilterNames(9) = GetINIMessage("Quick Filter", "Rarity", "Rarity")
        ga_QuickFilterNames(10) = GetINIMessage("Quick Filter", "PersonalComment", "Personal Comment")
        ga_QuickFilterNames(11) = GetINIMessage("Quick Filter", "VersionComment", "Version Comment")

        'sidebar quicklinks
        m_strQLPlayGame = GetINIMessage("Quick Links Window", "PlayGame", "Play Game")
        m_strQLPlayMusic = GetINIMessage("Quick Links Window", "PlayMusic", "Play Music")

    End With

    'details tabs titles
    With tabDetails
        .TabCaption(0) = " " & GetINIMessage("MainWindow", "Tab_Game_Title", "Game")
        .TabCaption(1) = " " & GetINIMessage("MainWindow", "Tab_Version_Title", "Version")
        .TabCaption(2) = " " & GetINIMessage("MainWindow", "Tab_Personal_Title", "Personal")
        .TabCaption(3) = " " & GetINIMessage("MainWindow", "Tab_Memo_Title", "Notes")
        .TabCaption(4) = " " & GetINIMessage("MainWindow", "Tab_Music_Title", "Music")
        tabSid.TabCaption(0) = .TabCaption(4)
    End With

    With g_xobjFunc
        'labels on the details tab "Game"
        lblPrompt(0) = GetINIMessage("MainWindow", "Tab_Game_Label_Copyright", "Copyright") & ":"
        lblPrompt(1) = GetINIMessage("MainWindow", "Tab_Game_Label_Developer", "Developer") & ":"
        lblPrompt(2) = GetINIMessage("MainWindow", "Tab_Game_Label_Programmer", "Programmer") & ":"
        lblPrompt(3) = GetINIMessage("MainWindow", "Tab_Game_Label_Musician", "Musician") & ":"
        lblPrompt(4) = GetINIMessage("MainWindow", "Tab_Game_Label_Artist", "Artist") & ":"
        lblPrompt(5) = GetINIMessage("MainWindow", "Tab_Game_Label_License", "License") & ":"
        lblPrompt(6) = GetINIMessage("MainWindow", "Tab_Game_Label_GameGenre", "Game Genre") & ":"
        lblPrompt(7) = GetINIMessage("MainWindow", "Tab_Game_Label_Players", "No. of Players") & ":"
        lblPrompt(8) = GetINIMessage("MainWindow", "Tab_Game_Label_Language", "Language") & ":"
        lblPrompt(9) = GetINIMessage("MainWindow", "Tab_Game_Label_Control", "Control") & ":"
        lblPrompt(10) = GetINIMessage("MainWindow", "Tab_Game_Label_Rarity", "Rarity") & ":"
        lblPrompt(11) = GetINIMessage("MainWindow", "Tab_Game_Label_WebLink", "More Info") & ":"

        'labels on the details tab "Version"
        lblPrompt(12) = GetINIMessage("MainWindow", "Tab_Version_Label_CrackingGroup", "Cracking Group") & ":"
        lblPrompt(13) = GetINIMessage("MainWindow", "Tab_Version_Label_Trainers", "No. of Trainers") & ":"
        lblPrompt(14) = GetINIMessage("MainWindow", "Tab_Version_Label_Playable", "Playable") & ":"
        lblPrompt(15) = GetINIMessage("MainWindow", "Tab_Version_Label_GameLength", "Game Length") & ":"
        lblPrompt(16) = GetINIMessage("MainWindow", "Tab_Version_Label_NTSC", "PAL/NTSC") & ":"
        lblPrompt(17) = GetINIMessage("MainWindow", "Tab_Version_Label_WebLink", "Download") & ":"
        lblPrompt(18) = GetINIMessage("MainWindow", "Tab_Version_Label_Original", "Original") & ":"
        lblPrompt(19) = GetINIMessage("MainWindow", "Tab_Version_Label_LoadingScreen", "Loading Screen") & ":"
        lblPrompt(20) = GetINIMessage("MainWindow", "Tab_Version_Label_TitleScreen", "Title Screen") & ":"
        lblPrompt(21) = GetINIMessage("MainWindow", "Tab_Version_Label_IncludedDocs", "Included Docs") & ":"
        lblPrompt(22) = GetINIMessage("MainWindow", "Tab_Version_Label_HighScoreSaver", "High Score Saver") & ":"
        lblPrompt(23) = GetINIMessage("MainWindow", "Tab_Version_Label_TrueDriveEmu", "True Drive Emul") & ":"
        
        'labels on the details tab "Personal"
        lblPrompt(24) = GetINIMessage("MainWindow", "Tab_Personal_Label_HighScore", "High Score") & ":"
        lblPrompt(25) = GetINIMessage("MainWindow", "Tab_Personal_Label_Difficulty", "Difficulty Level") & ":"
        lblPrompt(26) = GetINIMessage("MainWindow", "Tab_Personal_Label_TimesPlayed", "Times Played") & ":"
        lblPrompt(27) = GetINIMessage("MainWindow", "Tab_Personal_Label_LastPlayed", "Last Played") & ":"
        lblPrompt(28) = GetINIMessage("MainWindow", "Tab_Personal_Label_Rating", "Rating") & ":"

        'labels on the details tab "Memo"
        cmdMemoSave.Caption = .GetINIMessage("MainWindow", "Tab_Memo_Button_Save", "Save")
        cmdMemoClear.Caption = .GetINIMessage("MainWindow", "Tab_Memo_Button_Clear", "Clear")
        cmdMemoExpand.Caption = .GetINIMessage("MainWindow", "Tab_Memo_Button_Expand", "Expand")

        'labels on the Music (for games) tab
        lblPrompt(29) = GetINIMessage("MainWindow", "Tab_Music_DBMusician", "Musician") & ":"
        lblPrompt(30) = GetINIMessage("MainWindow", "Tab_Music_DBNick", "Nick") & ":"
        lblPrompt(31) = GetINIMessage("MainWindow", "Tab_Music_DBGroup", "Group") & ":"
        lblPrompt(32) = GetINIMessage("MainWindow", "Tab_Music_DBPhoto", "Photo") & ":"
        lblPrompt(33) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoTitle", "Sid-File Info")
        lblPrompt(34) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoSongs", "Songs") & ":"
        lblPrompt(35) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoDefault", "Default") & ":"
        lblPrompt(36) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoName", "Name") & ":"
        lblPrompt(37) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoAuthor", "Author") & ":"
        lblPrompt(38) = .GetINIMessage("MainWindow", "Tab_Music_Label_InfoCopyright", "Copyright") & ":"
        lblSidNotFound = ""
        lblSidNotFound2 = lblSidNotFound
        g_strSidNone = GetINIMessage("MainWindow", "Tab_Music_DBNone", "(None)")

        'labels on the Music (only, no games) tab
        lblPrompt(39) = lblPrompt(29)
        lblPrompt(40) = lblPrompt(30)
        lblPrompt(41) = lblPrompt(31)
        lblPrompt(42) = lblPrompt(32)
        lblPrompt(43) = lblPrompt(33)
        lblPrompt(44) = lblPrompt(34)
        lblPrompt(45) = lblPrompt(35)
        lblPrompt(46) = lblPrompt(36)
        lblPrompt(47) = lblPrompt(37)
        lblPrompt(48) = lblPrompt(38)

        'top level menus
        mnuFile.Caption = .GetINIMessage("MainMenus", "Menu_File", "&File")
        mnuEdit.Caption = .GetINIMessage("MainMenus", "Menu_Edit", "&Edit")
        mnuView.Caption = .GetINIMessage("MainMenus", "Menu_View", "&View")
        mnuGemus.Caption = g_xobjGEMUS.GetINIMessage("Menu", "GEMUS", "&GEMUS")
        mnuTools.Caption = .GetINIMessage("MainMenus", "Menu_Tools", "&Tools")
        mnuRecent.Caption = .GetINIMessage("MainMenus", "Menu_Recent", "&Recent")
        mnuGBPick.Caption = .GetINIMessage("MainMenus", "Menu_GameBase", "Game&Base")
        mnuHelp.Caption = .GetINIMessage("MainMenus", "Menu_Help", "&Help")

        'file menu
        mnuFileAdd.Caption = .GetINIMessage("MainMenus", "Menu_File_Add", "&Add")
        mnuFileRem.Caption = .GetINIMessage("MainMenus", "Menu_File_Remove", "&Remove")
        mnuFileExit.Caption = .GetINIMessage("MainMenus", "Menu_File_Exit", "E&xit")

        'file menu - add
        mnuFileAddGame.Caption = GetINIMessage("MainMenus", "Menu_File_Add_Game", "&Game") & "..."
        mnuFileAddSid.Caption = GetINIMessage("MainMenus", "Menu_File_Add_Music", "&Tune to Music Table") & "..."

        'file menu - remove
        mnuFileRemove.Caption = GetINIMessage("MainMenus", "Menu_File_RemoveSelected", "&Selected Game/Tune") & "..."
        mnuFileRemoveAll.Caption = GetINIMessage("MainMenus", "Menu_File_RemoveAll", "&All Games/Tunes in Current View") & "..."

        'edit menu
        mnuEditGame.Caption = GetINIMessage("MainMenus", "Menu_Edit_GameTuneInfo", "G&ame/Tune Info")
        mnuEditVersion.Caption = GetINIMessage("MainMenus", "Menu_Edit_VersionInfo", "&Version Info")
        mnuEditPersonal.Caption = GetINIMessage("MainMenus", "Menu_Edit_PersonalInfo", "Pers&onal Info")
        mnuEditGameFile.Caption = GetINIMessage("MainMenus", "Menu_Edit_GameFile", "&Game File") & "..."
        mnuEditSidFile.Caption = GetINIMessage("MainMenus", "Menu_Edit_MusicFile", "Mu&sic File") & "..."
        mnuEditScrnShots.Caption = GetINIMessage("MainMenus", "Menu_Edit_Pictures", "Pictu&res") & "..."
        mnuEditExtras.Caption = GetINIMessage("MainMenus", "Menu_Edit_Extras", "&Extras") & "..."
        mnuEditGemus.Caption = GetINIMessage("MainMenus", "Menu_Edit_Gemus", "GEMUS ""&key=value"" pairs") & "..."
        mnuEditMark.Caption = GetINIMessage("MainMenus", "Menu_Edit_MarkGame", "&Mark Game for Linking")
        mnuEditLinkSet.Caption = GetINIMessage("MainMenus", "Menu_Edit_LinkSet", "Set Game &Link") & "..."
        mnuEditLinkClear.Caption = GetINIMessage("MainMenus", "Menu_Edit_LinkClear", "&Clear Game Link") & "..."
        mnuEditFind.Caption = GetINIMessage("MainMenus", "Menu_Edit_Find", "&Find in List") & "..."
        mnuEditFindNext.Caption = GetINIMessage("MainMenus", "Menu_Edit_FindNext", "Find &Next")
        m_strAddToFavs = GetINIMessage("MainMenus", "Menu_Edit_AddToFavourites", "Add to Favo&urites")
        m_strRemoveFromFavs = GetINIMessage("MainMenus", "Menu_Edit_RemoveFromFavourites", "Remove from Favo&urites")
        mnuEditFindID.Caption = GetINIMessage("MainMenus", "Menu_Edit_FindByID", "Find &by ID")

        'edit menu > find by id
        mnuEditFindIDGame.Caption = GetINIMessage("MainMenus", "Menu_Edit_FindByID_Game", "&Game") & "..."
        mnuEditFindIDMusic.Caption = GetINIMessage("MainMenus", "Menu_Edit_FindByID_Music", "&Music") & "..."

        'edit menu > game info
        mnuEditGameCurrent.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_CurrentGameTune", "&Current Game/Tune") & "..."
        mnuEditGameAll.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_AllGamesTunes", "&All Games/Tunes in Current View") & "..."

        'edit menu > version info
        mnuEditVersionCurrent.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_CurrentGame", "&Current Game") & "..."
        mnuEditVersionAll.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_AllGames", "&All Games in Current View") & "..."

        'edit menu > personal info
        mnuEditPersonalCurrent.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_CurrentGame", "&Current Game") & "..."
        mnuEditPersonalAll.Caption = GetINIMessage("MainMenus", "Menu_Edit_Info_AllGames", "&All Games in Current View") & "..."

        'view menu
        mnuViewInfo.Caption = .GetINIMessage("MainMenus", "Menu_View_GameDetails", "&Game Details")
        mnuViewSidebar(0).Caption = .GetINIMessage("MainMenus", "Menu_View_SideBarHide", "H&ide Sidebar")
        mnuViewSidebar(1).Caption = .GetINIMessage("MainMenus", "Menu_View_SideBarLeft", "Sho&w Sidebar on Left")
        mnuViewSidebar(2).Caption = .GetINIMessage("MainMenus", "Menu_View_SideBarRight", "Show Side&bar on Right")
        mnuViewScreenshots(0).Caption = .GetINIMessage("MainMenus", "Menu_View_HidePictures", "&Hide Pictures")
        mnuViewScreenshots(1).Caption = .GetINIMessage("MainMenus", "Menu_View_SmallPictures", "&Small Pictures")
        mnuViewScreenshots(2).Caption = .GetINIMessage("MainMenus", "Menu_View_FullSizePictures", "&Full Size Pictures")
        mnuViewSSSnap.Caption = .GetINIMessage("MainMenus", "Menu_View_SnapToPictureSize", "S&nap to Picture Size")
        mnuViewExtraColumns.Caption = .GetINIMessage("MainMenus", "Menu_View_ExtraColumns", "Extra &Columns") & "..."
        mnuViewListList.Caption = .GetINIMessage("MainMenus", "Menu_View_List", "&List")
        mnuViewListDetails.Caption = .GetINIMessage("MainMenus", "Menu_View_Details", "&Details")
        mnuViewFontCols.Caption = .GetINIMessage("MainMenus", "Menu_View_FontCols", "List F&ont/Colours") & "..."
        mnuViewAddView.Caption = .GetINIMessage("MainMenus", "Menu_View_AddView", "&Add View") & "..."
        mnuViewEditView.Caption = .GetINIMessage("MainMenus", "Menu_View_EditView", "&Edit View") & "..."
        mnuViewRemoveView.Caption = .GetINIMessage("MainMenus", "Menu_View_RemoveView", "&Remove View")
        mnuViewQuickFilter.Caption = .GetINIMessage("MainMenus", "Menu_View_QuickFilter", "&Quick Filter") & "..."
        mnuViewCategories.Caption = .GetINIMessage("MainMenus", "Menu_View_CategoryWindow", "Ca&tegory Window")
        mnuViewFocusToList.Caption = .GetINIMessage("MainMenus", "Menu_View_SetFocusToList", "Set Foc&us to List")
        mnuViewExportToTextFile.Caption = .GetINIMessage("MainMenus", "Menu_View_ExportToTextFile", "E&xport to Text File")
        mnuViewExportToRomLister.Caption = .GetINIMessage("MainMenus", "Menu_View_ExportToRomLister", "Export to Ro&mLister") & "..."

        'gemus menu
        With g_xobjGEMUS
            mnuGemusManage.Caption = .GetINIMessage("Menu", "ManageEmulators", "&Manage Emulators") & "..."
            mnuGemusFileTypes.Caption = .GetINIMessage("Menu", "SupportedFileTypes", "&Default Emulator File Types") & "..."
            mnuGemusFields.Caption = .GetINIMessage("Menu", "DefaultFields", "Set Default GEMUS Script &Fields") & "..."
            mnuGemusWizard.Caption = .GetINIMessage("Menu", "Wizard", "Emulator Setup &Wizard") & "..."
            mnuGemusExport.Caption = .GetINIMessage("Menu", "ExportEmulatorDefinitions", "&Export Emulator Definitions")
        End With

        'tools menu
        mnuToolsSounds.Caption = .GetINIMessage("MainMenus", "Menu_Tools_Sounds", "Play &Sound on Classic Game")
        mnuToolsNSM.Caption = .GetINIMessage("MainMenus", "Menu_Tools_NonStopMusic", "&Non Stop Music")
        mnuToolsHideAdult.Caption = .GetINIMessage("MainMenus", "Menu_Tools_HideAdult", "&Adult Filter") & "..."
        mnuToolsAdultPassword.Caption = .GetINIMessage("MainMenus", "Menu_Tools_ResetPassword", "&Reset Password") & "..."
        mnuToolsOptions.Caption = .GetINIMessage("MainMenus", "Menu_Tools_Options", "Frontend &Options") & "..."
        mnuToolsDatabaseOptions.Caption = g_strGameBaseName & " " & .GetINIMessage("MainMenus", "Menu_Tools_Database_Options", "Op&tions") & "..."
        mnuToolsPaths.Caption = g_strGameBaseName & " " & .GetINIMessage("MainMenus", "Menu_Tools_Paths", "&Paths") & "..."
        mnuToolsVerify.Caption = .GetINIMessage("MainMenus", "Menu_Tools_Verify", "&Verify Available Files") & "..."

        'recent menu
        g_strRecentEmpty = .GetINIMessage("MainMenus", "Menu_Recent_Empty", "Nothing here yet!")

        'help menu
        mnuHelpGamebase64.Caption = .GetINIMessage("MainMenus", "Menu_Help_Help", "GameBase &Help") & "..."
        mnuHelpHome.Caption = .GetINIMessage("MainMenus", "Menu_Help_Web", "GameBase &Web Site")
        mnuHelpEmail.Caption = .GetINIMessage("MainMenus", "Menu_Help_EmailAuthor", "&Email the Author")
        mnuHelpAbout.Caption = .GetINIMessage("MainMenus", "Menu_Help_About", "&About") & "..."

        'error messages
        m_strErrorItemNotFound = .GetINIMessage("Error Messages", "Error_ItemNotFound", "The Game/Music was not found in the list!")
        m_strErrorItemNotFoundCurrent = .GetINIMessage("Error Messages", "Error_ItemNotFoundCurrent", "The Game/Music you requested was not found in the current view!")
        m_strErrorRequestRefresh = .GetINIMessage("Error Messages", "Error_RequestRefresh", "Do you want me to change to the main view so that I can show you the Game/Music?")
    End With

End Sub

Private Function ItemInListExists(strKey_in As String, ByVal lngListViewIdx_in As Long) As Boolean

    '--------------------------------------------------------------------
    'finds out if the item specified by the passed in key exists in
    'the currently selected lvMain() listview or not by forcing an error.
    '
    'Inputs:    strKey_in:          The key if the item to find
    '           lngListViewIdx_in:  The index of the listview to search
    '
    'Returns:   True if the listitem exists
    '           False otherwise
    '--------------------------------------------------------------------
        
    Dim blnDummy As Boolean     'dummy variable for test
    
    'set error handler
    On Error GoTo NoItem
    
    'test by forcing an error
    blnDummy = lvMain(lngListViewIdx_in).ListItems(strKey_in).Bold
    'test passed!
    ItemInListExists = True
    
    Exit Function
    
NoItem:
    'Item doesn't exist
    ItemInListExists = False

End Function

Private Sub LoadCustomViewsFromDB()

    '--------------------------------------------------------------
    'loads the Custom Views from the database into the memory array
    'and makes sure everything is ok (deletes corrupted views)
    '--------------------------------------------------------------

    Dim lngViewIdx As Long          'loop index
    Dim lngFilterIdx As Long        'loop index
    Dim lngViewMax As Long          'storage for UBound() of views array
    Dim lngBadIdx As Long           'corrupted views counter
    Dim a_lngBadViews() As Long     'corrupted views index to delete
    Dim strExtraColumns As String
    Dim a_strColumnWidths As Long
    Dim strKeyName As String
    Dim lngColumnIdx As Long

    '''''''''''''''''''''''''''
    'read in the list of views
    '''''''''''''''''''''''''''

    strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\CustomViews"

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM ViewData ORDER BY Ordinal", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            'there is at least one Custom View in the database
            lngViewIdx = gc_lngDEFAULTVIEWSMAX + 1
            ReDim Preserve ga_udtViewData(lngViewIdx)
            .MoveFirst
            Do Until .EOF
                ReDim Preserve ga_udtViewData(lngViewIdx)
                
                ga_udtViewData(lngViewIdx).lngID = !ID
                ga_udtViewData(lngViewIdx).lngFilterCount = !FilterCount
                ga_udtViewData(lngViewIdx).eFilterMode = !FilterMode
                ga_udtViewData(lngViewIdx).eIncludeMusicTable = !IncludeMusicTable
                ga_udtViewData(lngViewIdx).lngListViewType = !ListViewType
                ga_udtViewData(lngViewIdx).strSelectedItem = !SelectedItem
                ga_udtViewData(lngViewIdx).lngSortColumn = !SortColumn
                ga_udtViewData(lngViewIdx).lngSortOrder = !SortOrder
                ga_udtViewData(lngViewIdx).strTitle = !Title
                strExtraColumns = !ExtraColumns
                ga_udtViewData(lngViewIdx).strExtraColumns = !ExtraColumns

                g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM _
                    strKeyName, _
                    "Col1Width_" & ga_udtViewData(lngViewIdx).lngID, _
                    ga_udtViewData(lngViewIdx).lngCol1Width, _
                    False, _
                    gc_lngDEFAULTCOLUMNWIDTH

                'parse the extra columns
                ga_udtViewData(lngViewIdx).lngExtraColumnsCount = ParseExtraColumns( _
                        ga_udtViewData(lngViewIdx).a_udtExtraColumns, strExtraColumns _
                )
                ga_udtViewData(lngViewIdx).lngMaxExtraColumns = UBound(ga_udtViewData(lngViewIdx).a_udtExtraColumns)
                
                'now the number of columns is known, get the extra columns widths
                For lngColumnIdx = 0 To ga_udtViewData(lngViewIdx).lngExtraColumnsCount - 1
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM _
                        strKeyName, _
                        "Col" & (lngColumnIdx + 2) & "Width_" & ga_udtViewData(lngViewIdx).lngID, _
                        ga_udtViewData(lngViewIdx).a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                        False, _
                        gc_lngDEFAULTEXTRACOLUMNWIDTH
                Next lngColumnIdx

                lngViewIdx = lngViewIdx + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set g_rsGames = Nothing
    
    
    ''''''''''''''''''''''''''''
    'read in each view's filters
    ''''''''''''''''''''''''''''
    
    'initialise the corrupted views array
    lngBadIdx = 0
    ReDim a_lngBadViews(lngBadIdx)
    
    'loop through reading in the filters recording the corrupted views index
    lngViewMax = UBound(ga_udtViewData)
    If lngViewMax > gc_lngDEFAULTVIEWSMAX Then 'are there any custom views?
        For lngViewIdx = (gc_lngDEFAULTVIEWSMAX + 1) To lngViewMax
            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM ViewFilters WHERE ID = " & ga_udtViewData(lngViewIdx).lngID, dbOpenSnapshot)
            With g_rsGames
                If .EOF Then
                    If ga_udtViewData(lngViewIdx).lngFilterCount > 0 Then
                        'no filters in DB and there should be!
                        lngBadIdx = lngBadIdx + 1
                        ReDim Preserve a_lngBadViews(lngBadIdx)
                        a_lngBadViews(lngBadIdx) = lngViewIdx
                    End If
                Else
                    .MoveLast
                    .MoveFirst
                    If .RecordCount <> ga_udtViewData(lngViewIdx).lngFilterCount Then
                        'wrong amount of filters in the database!
                        lngBadIdx = lngBadIdx + 1
                        ReDim Preserve a_lngBadViews(lngBadIdx)
                        a_lngBadViews(lngBadIdx) = lngViewIdx
                    Else
                        ReDim ga_udtViewData(lngViewIdx).a_udtFilters(ga_udtViewData(lngViewIdx).lngFilterCount)
                        For lngFilterIdx = 1 To ga_udtViewData(lngViewIdx).lngFilterCount
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).strClauseData = !ClauseData
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).eClauseType = !ClauseType
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).strFieldName = !FieldName
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).strFieldTable = !FieldTable
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).strMusicFieldName = !MusicFieldName
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).strMusicFieldTable = !MusicFieldTable
                            ga_udtViewData(lngViewIdx).a_udtFilters(lngFilterIdx).eOperator = !Operator
                            .MoveNext
                        Next lngFilterIdx
                    End If
                End If
                .Close
            End With
            Set g_rsGames = Nothing
        Next lngViewIdx
    End If
    

    '''''''''''''''''''''''''''
    'Delete any corrupted views
    '''''''''''''''''''''''''''
    
    If lngBadIdx > 0 Then
        'delete from bottom up as we are working with the array's index
        For lngViewIdx = lngBadIdx To 1 Step -1
            Call DeleteCustomView(a_lngBadViews(lngViewIdx))
        Next lngViewIdx
    End If

End Sub
Private Sub LoadDefaultViews()

    '----------------------------------------------------------
    'Initialises the Views array, sets the default view queries
    'and reads in the default view data from the registry.
    '
    'Remarks:       If the default Views change, make sure to
    '               change the CountSQL's in FillViewSelector()
    '----------------------------------------------------------
    
    Dim lngViewIdx As Long      'loop index
    Dim lngColumnIdx As Long
    Dim strKeyName As String
    Dim strExtraColumns As String
    Dim a_strColumnWidths() As Long

    'initialise the array
    ReDim ga_udtViewData(gc_lngDEFAULTVIEWSMAX)

    strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\DefaultViews"
    
    'set common properties
    For lngViewIdx = 0 To gc_lngDEFAULTVIEWSMAX
        With ga_udtViewData(lngViewIdx)
            'general
            .eFilterMode = e_FilterMode_AND
            .lngID = 0
            .strTitle = "<" & ga_strDefaultViewNames(lngViewIdx) & ">"
            'from the registry
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "ListViewType_" & lngViewIdx, .lngListViewType, False, lvwReport
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SortColumn_" & lngViewIdx, .lngSortColumn, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SortOrder_" & lngViewIdx, .lngSortOrder, False, 0
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "SelectedItem_" & lngViewIdx, .strSelectedItem, True, "INVALID"
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "ExtraColumns_" & lngViewIdx, strExtraColumns, True, ""
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM strKeyName, "Col1Width_" & lngViewIdx, .lngCol1Width, False, gc_lngDEFAULTCOLUMNWIDTH
        
            'parse the extra columns
            .strExtraColumns = strExtraColumns
            .lngExtraColumnsCount = ParseExtraColumns(.a_udtExtraColumns, strExtraColumns)
            .lngMaxExtraColumns = UBound(.a_udtExtraColumns)

            'now the number of columns is known, get the extra columns widths
            For lngColumnIdx = 0 To .lngExtraColumnsCount - 1
                g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM _
                    strKeyName, _
                    "Col" & (lngColumnIdx + 2) & "Width_" & lngViewIdx, _
                    .a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                    False, _
                    gc_lngDEFAULTEXTRACOLUMNWIDTH
            Next lngColumnIdx
        End With
    Next lngViewIdx

    'set unique properties
    With ga_udtViewData(0)
        'All Games
        ReDim .a_udtFilters(0)
        .eIncludeMusicTable = e_Table_GamesOnly
        .lngFilterCount = 0
    End With
    With ga_udtViewData(1)
        'All Sids
        ReDim .a_udtFilters(1)
        With .a_udtFilters(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strClauseData = ""
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
        End With
        .eIncludeMusicTable = e_Table_Both
        .lngFilterCount = 1
    End With
    With ga_udtViewData(2)
        'Favourite Games/Sids
        ReDim .a_udtFilters(1)
        With .a_udtFilters(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .strFieldName = "Fav"
            .strFieldTable = "Games"
            .strClauseData = "True"
            .strMusicFieldName = "SFav"
            .strMusicFieldTable = "Music"
        End With
        .eIncludeMusicTable = e_Table_Both
        .lngFilterCount = 1
    End With
    With ga_udtViewData(3)
        'Games No clones
        ReDim .a_udtFilters(1)
        With .a_udtFilters(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .strFieldName = "CloneOf"
            .strFieldTable = "Games"
            .strClauseData = "0"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eOperator = e_Op_Equal
        End With
        .eIncludeMusicTable = e_Table_GamesOnly
        .lngFilterCount = 1
    End With

End Sub

Private Sub LoadScrnShotUnavailableImages()
    
    ''''''''''''''''''''''''''''''''''''''''
    'loads the screenshot unavailable images
    'if they exist in the gb folder
    '
    'Remarks:   This is solely for language
    '           support purposes
    ''''''''''''''''''''''''''''''''''''''''

    'big image
    If g_xobjFunc.FileExists(g_strGameBasePathWBS & "Gfx\No_Picture_320x200.gif", False, "") Then
        picNoSS.Picture = LoadPicture(g_strGameBasePathWBS & "Gfx\No_Picture_320x200.gif")
    End If
    
    'small image
    If g_xobjFunc.FileExists(g_strGameBasePathWBS & "Gfx\No_Picture_160x100.gif", False, "") Then
        picNoSSSmall.Picture = LoadPicture(g_strGameBasePathWBS & "Gfx\No_Picture_160x100.gif")
    End If

    
End Sub

Private Sub LoadExtraColumnsInfo()

    '----------------------------------------------------
    'Loads the extra columns information from GBSql.dll
    'into the extra columns field arrays
    '----------------------------------------------------

    Dim xobjExtraColumns As GBSql.clsSecCol             'the extra column object
    Dim lngMenuIdx As Long                              'loop index
    Dim lngFieldIdx As Long                             'loop index
    Dim lngTempType As Long                             'because we can't cast in vb

    'create the object
    Set xobjExtraColumns = New GBSql.clsSecCol
    Call xobjExtraColumns.SetIniFile(g_strMessagesFile)
    Call xobjExtraColumns.InitColumns

    'get the total amount of extra column menus and redim the menus array
    With xobjExtraColumns
        g_lngMaxExtraColumnArrays = .MaxArrays
        ReDim ga_udtExtraColumns(g_lngMaxExtraColumnArrays)
        'redim all the menu arrays and read in menu text
        For lngMenuIdx = 0 To g_lngMaxExtraColumnArrays
            ga_udtExtraColumns(lngMenuIdx).lngMaxColumns = .MaxFields(lngMenuIdx)
            ga_udtExtraColumns(lngMenuIdx).strMenuText = .MenuText(lngMenuIdx)
            ReDim ga_udtExtraColumns(lngMenuIdx).a_udtColumnItems(ga_udtExtraColumns(lngMenuIdx).lngMaxColumns)
        Next lngMenuIdx
    End With

    'get the extra columns field info
    For lngMenuIdx = 0 To g_lngMaxExtraColumnArrays
        With ga_udtExtraColumns(lngMenuIdx)
            For lngFieldIdx = 0 To .lngMaxColumns
                With .a_udtColumnItems(lngFieldIdx)
                    Call xobjExtraColumns.GetSecondColField(lngMenuIdx, lngFieldIdx, .strMenuText, .strColumnTitle, .strFields, .strMusicFields, lngTempType, .blnGameOnlyTable)
                    .eType = lngTempType
                End With
            Next lngFieldIdx
        End With
    Next lngMenuIdx

    'release the object
    Set xobjExtraColumns = Nothing

End Sub



Private Sub SetInitialGlobalValues()

    Dim strFileTypes As String
    Dim lngTemp As Long

    'We're not subclassed yet!
    g_blnSubclassed = False
    g_blnCatSubclassed = False
    g_blnPictureSelectorSubclassed = False
    g_blnGameSelectorSubclassed = False
    g_blnExtrasSubclassed = False

    'not resizing any windows yet
    g_blnResizingWindow = False

    'tabDetails
    g_sngCopyrightGap = Screen.TwipsPerPixelX * 4 '1986<--this-gap-in-pixels-->Mastertronic

    'recent menu
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "RecentUpdateGameMusic", g_blnRecentGameSid, False, False
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "RecentStore", g_lngRecentStore, False, gc_lngRECENT_BOTH
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "RecentUpdate", g_lngRecentUpdate, False, gc_lngRECENT_SELECTED
    If (g_lngRecentStore < gc_lngRECENT_GAMES) Or (g_lngRecentStore > gc_lngRECENT_BOTH) Then
        g_lngRecentStore = gc_lngRECENT_BOTH
    End If
    If (g_lngRecentUpdate < gc_lngRECENT_SELECTED) Or (g_lngRecentUpdate > gc_lngRECENT_PLAYED) Then
        g_lngRecentUpdate = gc_lngRECENT_SELECTED
    End If

    'loaded forms
    g_blnLoadedfrmBigScreenshot = False
    g_blnLoadedfrmCategories = False
    g_blnLoadedfrmCategories = False

    'listview selected item type
    g_eWhatIsSelected = e_SelType_IsNone

    'load image extras into the screenshots area?
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "LoadImageExtrasAsScreenshots", g_blnLoadImageExtrasAsScreenshots, False, True

    'the current view
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "ViewIndex", g_lngRegViewIndex, False, -1
    If g_lngRegViewIndex = -1 Then g_lngRegViewIndex = 0

    'boss key
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "BossKeyOn", g_blnBossKeyOn, False, False

    'rezipping games
    g_blnReZip = (g_xobjFunc.INI_GetSetting("General", "ReZipGames", "0", g_strSettingsFile) = "1")
    g_blnRezipAsk = (g_xobjFunc.INI_GetSetting("General", "ReZipAsk", "1", g_strSettingsFile) = "1")
    g_blnZipNotify = (g_xobjFunc.INI_GetSetting("General", "ReZipNotify", "1", g_strSettingsFile) = "1")
    lngTemp = Val(g_xobjFunc.INI_GetSetting("General", "ReZipAdditional", "2", g_strSettingsFile))
    If lngTemp = 0 Then
        ge_RezipAdditional = eRezipAdditionalIgnore
    ElseIf lngTemp = 1 Then
        ge_RezipAdditional = eRezipAdditionalZip
    Else
        ge_RezipAdditional = eRezipAdditionalAsk
    End If

    'show game id in title bar
    lngTemp = g_xobjFunc.INI_GetSetting("UI", "ShowGameID", "0", g_strSettingsFile)
    g_blnOptionsShowGameID = (lngTemp = 1)

    'the adult filter
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\WinPlant3", "Seed", g_blnRegHideAdultItems, False, False

    'do we use Filter64 when clicking on field data in the details tab?
    g_blnOpenFilter64 = False
    g_blnSettingCatFilter = False

    'shut down sidplayers on exit?
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "CloseMusicPlayersOnExit", g_blnShutDownSidPlayersOnExit, False, False

    'device prompts
    g_blnPromptForGamePathDevice = (g_xobjFunc.INI_GetSetting("General", "RequestGameDeviceInsert", "1", g_strSettingsFile) = "1")

    'zoom picture
    g_blnSetPicZoomForeground = (g_xobjFunc.INI_GetSetting("General", "BringZoomedPicToFront", "1", g_strSettingsFile) = "1")

    'the pictures backcolor
    g_lngPictureBackColor = Val(g_xobjFunc.INI_GetSetting("UI", "PicturesBackColor", "0", g_strSettingsFile))

    'the pictures style
    g_blnMamePicStyle = (g_xobjFunc.INI_GetSetting("General", "MamePictureStyle", "0", g_strSettingsFile) = "1")

    'double click action
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "DoubleClickAction", g_eDoubleClickAction, False, e_DoubleClick_PlayGame

    'screenshot border sizes
    imgDummySS.Picture = picNoSS.Picture
    g_sngScrnShotExtraTwipsX = imgDummySS.Width - picNoSS.Width
    g_sngScrnShotExtraTwipsY = imgDummySS.Height - picNoSS.Height
    g_lngScrnShotExtraPixelsX = g_sngScrnShotExtraTwipsX / Screen.TwipsPerPixelX
    g_lngScrnShotExtraPixelsY = g_sngScrnShotExtraTwipsY / Screen.TwipsPerPixelY
    imgDummySS.Picture = LoadPicture()

    'sidplayer handles
    g_lngTotalSidPlayerHandles = 0

    'game link field
    g_strDefaultGameLinkField = "Prequel"

    'focus enums
    g_udtEditGameInfoFocus = e_InfoFocus_Year

    'the database field names
    With g_xobjFunc
        g_strField_ExtrasName = GetINIMessage("Database", "Field_Extras_Name", "Extras Name")
        g_strField_Name = GetINIMessage("Database", "Field_Name", "Name")
        g_strField_Year = GetINIMessage("Database", "Field_Year", "Year Published")
        g_strField_Comment = GetINIMessage("Database", "Field_Comment", "Personal Comment")
        g_strField_GameFilename = GetINIMessage("Database", "Field_GameFilename", "Game Filename")
        g_strField_ScreenShotFilename = GetINIMessage("Database", "Field_PictureFilename", "Picture Filename")
        g_strField_SidFilename = GetINIMessage("Database", "Field_MusicFilename", "Music Filename")
        g_strField_Musician = GetINIMessage("Database", "Field_Musician", "Musician")
        g_strField_Genre = GetINIMessage("Database", "Field_Genre", "Genre")
        g_strField_ParentGenre = GetINIMessage("Database", "Field_ParentGenre", "Parent Genre")
        g_strField_Publisher = GetINIMessage("Database", "Field_Publisher", "Publisher")
        g_strField_Difficulty = GetINIMessage("Database", "Field_Difficulty", "Difficulty")
        g_strField_Cracker = GetINIMessage("Database", "Field_Cracker", "Cracker")
        g_strField_Programmer = GetINIMessage("Database", "Field_Programmer", "Programmer")
        g_strField_Language = GetINIMessage("Database", "Field_Language", "Language")
        g_strField_DateLastPlayed = GetINIMessage("Database", "Field_DateLastPlayed", "Last Played")
        g_strField_TimesPlayed = GetINIMessage("Database", "Field_TimesPlayed", "Times Played")
        g_strField_HighScore = GetINIMessage("Database", "Field_HighScore", "High Score")
        g_strField_Favourite = GetINIMessage("Database", "Field_Favourite", "Favourite")
        g_strField_Extras = GetINIMessage("Database", "Field_Extras", "Game Has Extras")
        g_strField_Classic = GetINIMessage("Database", "Field_Classic", "Classic Game")
        g_strField_Rating = GetINIMessage("Database", "Field_Rating", "Rating")
        g_strField_LoadingScreen = GetINIMessage("Database", "Field_LoadingScreen", "Loading Screen")
        g_strField_HighScoreSaver = GetINIMessage("Database", "Field_HighScoreSaver", "High Score Saver")
        g_strField_IncludedDocs = GetINIMessage("Database", "Field_IncludedDocs", "Included Docs")
        g_strField_PALNTSC = GetINIMessage("Database", "Field_PALNTSC", "PAL/NTSC")
        g_strField_TrueDriveEmulation = GetINIMessage("Database", "Field_TrueDriveEmulation", "True Drive Emul")
        g_strField_GameLength = GetINIMessage("Database", "Field_GameLength", "Game Length")
        g_strField_GameLengthType = GetINIMessage("Database", "Field_GameLengthType", "Game Length Type")
        g_strField_Trainers = GetINIMessage("Database", "Field_Trainers", "No. of Trainers")
        g_strField_NumPlayers = GetINIMessage("Database", "Field_NumPlayers", "No. of Players")
        g_strField_Players_Simultaneous = GetINIMessage("Field_Database", "Players_Simultaneous", "Simultaneous Play")
        g_strField_VersionComment = GetINIMessage("Database", "Field_VersionComment", "Version Comment")
        g_strField_GameVer = GetINIMessage("Database", "Field_GameVersion", "Game-File Version")
        g_strField_Prequel = GetINIMessage("Database", "Field_Precursor", "Has Precursor")
        g_strField_Sequel = GetINIMessage("Database", "Field_Sequel", "Has Sequel")
        g_strField_Related = GetINIMessage("Database", "Field_Related", "Has Related Game")
        g_strField_CloneOf = GetINIMessage("Database", "Field_CloneOf", "Is Clone")
        g_strField_Notes = GetINIMessage("Database", "Field_Notes", "Notes")
        g_strField_MusicianGroup = GetINIMessage("Database", "Field_MusicianGroup", "Musician Group")
        g_strField_MusicianNick = GetINIMessage("Database", "Field_MusicianNick", "Musician Nickname")
        g_strField_MusicianPhoto = GetINIMessage("Database", "Field_MusicianPhoto", "Musician Photo")
        g_strField_Control = GetINIMessage("Database", "Field_Control", "Control")
        g_strField_Adult = GetINIMessage("Database", "Field_Adult", "Adult")
        g_strField_Artist = GetINIMessage("Database", "Field_Artist", "Artist")
        g_strField_Developer = GetINIMessage("Database", "Field_Developer", "Developer")
        g_strField_License = GetINIMessage("Database", "Field_License", "License")
        g_strField_Rarity = GetINIMessage("Database", "Field_Rarity", "Rarity")
        g_strField_WebLinkName = GetINIMessage("Database", "Field_WebLinkName", "More Info Name")
        g_strField_WebLinkURL = GetINIMessage("Database", "Field_WebLinkURL", "More Info URL")
        g_strField_VWebLinkName = GetINIMessage("Database", "Field_VersionWebLinkName", "Download Name")
        g_strField_VWebLinkURL = GetINIMessage("Database", "Field_VersionWebLinkURL", "Download URL")
        g_strField_TitleScreen = GetINIMessage("Database", "Field_TitleScreen", "Title Screen")
        g_strField_Playable = GetINIMessage("Database", "Field_Playable", "Playable")
        g_strField_Original = GetINIMessage("Database", "Field_Original", "Original")
        g_strField_ReviewRating = GetINIMessage("Database", "Field_ReviewRating", "Review Rating")
    End With

    'other
    g_blnIsFirstItemDisplayed = False

End Sub

Public Sub SetItemInList(strRequestedKey_in As String, ByVal blnMustFindItem_in As Boolean, Optional varFullViewIdx_in As Variant, Optional varCurrentKey_in As Variant)

    '--------------------------------------------------------------------------------
    'Attempts to set the requested item in the lvMain() list
    '
    'Inputs:        strRequestedKey_in:     The key of the requested item.
    '               blnMustFindItem_in:     Flag specifying whether to refresh to
    '                                       the main view if the item can't be found.
    '               varCurrentKey_in:       The currently selected item's key (if
    '                                       there is one).
    '               varFullViewIdx_in:      The view to refresh to to find the item.
    '
    'Assumptions:   There must be at least one item in the current list in lvMain().
    '
    'Remarks:       varCurrentKey_in should be passed in if blnMustFindItem_in is set
    '               to True so that when switching to the main view we can find the
    '               old selected game if the requested game can't be found (unlikely)
    '--------------------------------------------------------------------------------

    Dim lngListViewIdx As Long      'lvMain() list view index
    Dim strCurrentKey As String     'the current item's key
    Dim lngFullViewIdx As Long      'view to refresh to
    
    'set error handler
    On Error GoTo NoItem
    
    'get the listview index to use from the current view index
    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    
    'get the view to refresh to if the item is not found
    If Not IsMissing(lngFullViewIdx) Then
        lngFullViewIdx = CLng(varFullViewIdx_in)
    Else
        lngFullViewIdx = 0
    End If
    
    'attempt to set the item
    Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(strRequestedKey_in)
    lvMain(lngListViewIdx).SelectedItem.Selected = True
    lvMain(lngListViewIdx).SelectedItem.EnsureVisible
    
    'if here then the item is found so show the item and exit the procedure
    On Error GoTo 0
    Call ShowItem
    
    Exit Sub
    
NoItem:

    '''''''''''''''''''
    'item is not found!
    '''''''''''''''''''
    
    'release the error handler
    On Error GoTo 0
        
    'test if we need to refresh and find the item
    If blnMustFindItem_in = False Then
        'not refreshing to the main default view so select the first item in the list
        If lvMain(lngListViewIdx).ListItems.Count = 0 Then
            Call SetControlsForView(False, 0)
        Else
            Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(1)
            lvMain(lngListViewIdx).SelectedItem.Selected = True
            lvMain(lngListViewIdx).SelectedItem.EnsureVisible
            Call ShowItem
        End If
    Else
        'refreshing to the main view
        If lstView.ListIndex = lngFullViewIdx Then
            '''''''''''''''''''''''''''''''''''''''''
            'already in main view and game not found!
            '''''''''''''''''''''''''''''''''''''''''
            MsgBox m_strErrorItemNotFound, vbExclamation Or vbOKOnly, gc_strAPPNAME
            'select the old item (if passed in) else select the first item in the list
            If lvMain(lngFullViewIdx).ListItems.Count = 0 Then
                Call SetControlsForView(False, 0)
            Else
                If IsMissing(varCurrentKey_in) Then
                    Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(1)
                Else
                    strCurrentKey = CStr(varCurrentKey_in)
                    Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(strCurrentKey)
                End If
                lvMain(lngListViewIdx).SelectedItem.Selected = True
                lvMain(lngListViewIdx).SelectedItem.EnsureVisible
                Call ShowItem
            End If
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''
            'ask if user wants to refresh to the main view
            ''''''''''''''''''''''''''''''''''''''''''''''
            If MsgBox(m_strErrorItemNotFoundCurrent & vbCrLf & vbCrLf & m_strErrorRequestRefresh, vbQuestion Or vbYesNoCancel, gc_strAPPNAME) = vbYes Then
                'refresh the view
                m_blnSetTheView = False 'don't want to do the code in lstView_Click() - want to do it here instead!
                lstView.ListIndex = lngFullViewIdx 'main view
                m_blnSetTheView = True 'reset the flag
                lngListViewIdx = lngFullViewIdx
                If SetView(lstView.ListIndex) = True Then 'set the view and read in the list
                    'attempt to set the item in the list
                    If ItemInListExists(strRequestedKey_in, lngListViewIdx) = True Then
                        'found game
                        Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(strRequestedKey_in)
                    Else
                        'can't find game
                        MsgBox m_strErrorItemNotFound, vbExclamation Or vbOKOnly, gc_strAPPNAME
                        'try to set to currently selected game
                        If IsMissing(varCurrentKey_in) Then
                            Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(1)
                        Else
                            strCurrentKey = Str(varCurrentKey_in)
                            Set lvMain(lngListViewIdx).SelectedItem = lvMain(lngListViewIdx).ListItems(strCurrentKey)
                        End If
                    End If
                    lvMain(lngListViewIdx).SelectedItem.Selected = True
                    lvMain(lngListViewIdx).SelectedItem.EnsureVisible
                    Call ShowItem
                Else
                    'no games in main list!
                    Call SetControlsForView(False, 0)
                End If
            End If
        End If
    End If
    
End Sub
Private Sub SetSortIcons(ByVal lngListViewIdx_in As Long)

    Dim lngIdx As Long

    With lvMain(lngListViewIdx_in)
        For lngIdx = 1 To .ColumnHeaders.Count
            If .SortKey = lngIdx - 1 Then
                If .SortOrder = lvwAscending Then
                    .ColumnHeaders(lngIdx).Icon = "down"
                Else
                    .ColumnHeaders(lngIdx).Icon = "up"
                End If
            Else
                .ColumnHeaders(lngIdx).Icon = "blank"
            End If
        Next lngIdx
    End With

End Sub

Private Sub Show0()

    Dim strTest As String

    With g_udtGame
        lbl_Info_License = .strLicense
        lbl_Info_Publisher = .strPublisher
        lbl_Info_Developer = .strDeveloper
        lbl_Info_Artist = .strArtist

        strTest = EllipseTextEnd(.strPublisher, 3650)
        lbl_Info_Publisher = strTest
        If strTest <> .strPublisher Then
            lbl_Info_Publisher.ToolTipText = .strPublisher
        Else
            lbl_Info_Publisher.ToolTipText = ""
        End If

        lbl_Info_Year = .strYear
 
        lbl_Info_Publisher.Left = lbl_Info_Year.Left + lbl_Info_Year.Width + g_sngCopyrightGap

        strTest = EllipseTextEnd(.strMusician, 4100)
        lbl_Info_Musician = strTest
        If strTest <> .strMusician Then
            lbl_Info_Musician.ToolTipText = .strMusician
        Else
            lbl_Info_Musician.ToolTipText = ""
        End If
        
        strTest = EllipseTextEnd(.strProgrammer, 4100)
        lbl_Info_Programmer = strTest
        If strTest <> .strProgrammer Then
            lbl_Info_Programmer.ToolTipText = .strProgrammer
        Else
            lbl_Info_Programmer.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strArtist, 4100)
        lbl_Info_Artist = strTest
        If strTest <> .strArtist Then
            lbl_Info_Artist.ToolTipText = .strArtist
        Else
            lbl_Info_Artist.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strDeveloper, 4100)
        lbl_Info_Developer = strTest
        If strTest <> .strDeveloper Then
            lbl_Info_Developer.ToolTipText = .strDeveloper
        Else
            lbl_Info_Developer.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strLicense, 4100)
        lbl_Info_License = strTest
        If strTest <> .strLicense Then
            lbl_Info_License.ToolTipText = .strLicense
        Else
            lbl_Info_License.ToolTipText = ""
        End If
        
        lbl_Info_Genre = .strGenre
        lbl_Info_Players = .strPlayers
        lbl_Info_Language = .strLanguage
        lbl_Info_Control = .strControl
        lbl_Info_Rarity = .strRarity
        lbl_Info_WebLink = .strWebLinkName
        
        strTest = EllipseTextEnd(.strLanguage, 2810)
        lbl_Info_Language = strTest
        If strTest <> .strLanguage Then
            lbl_Info_Language.ToolTipText = .strLanguage
        Else
            lbl_Info_Language.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strRarity, 2810)
        lbl_Info_Rarity = strTest
        If strTest <> .strRarity Then
            lbl_Info_Rarity.ToolTipText = .strRarity
        Else
            lbl_Info_Rarity.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strWebLinkName, 2810)
        lbl_Info_WebLink = strTest
        If strTest <> .strWebLinkName Then
            lbl_Info_WebLink.ToolTipText = .strWebLinkName
        Else
            lbl_Info_WebLink.ToolTipText = ""
        End If
        If .blnHasWebLink Then
            lbl_Info_WebLink.ForeColor = lblWebLink.ForeColor
        Else
            lbl_Info_WebLink.ForeColor = lbl_Info_Publisher.ForeColor
        End If

        lbl_Info_Comment = .strComment
        .blnLoaded_0 = True
    End With

End Sub
Private Sub Show1()
    
    Dim strTest As String

    With g_udtGame

        strTest = EllipseTextEnd(.strCracker, 4965)
        lbl_Info_Cracker = strTest
        If strTest <> .strCracker Then
            lbl_Info_Cracker.ToolTipText = .strCracker
        Else
            lbl_Info_Cracker.ToolTipText = ""
        End If

        strTest = EllipseTextEnd(.strVWebLinkName, 3765)
        lbl_Info_VWebLink = strTest
        If strTest <> .strVWebLinkName Then
            lbl_Info_VWebLink.ToolTipText = .strWebLinkName
        Else
            lbl_Info_VWebLink.ToolTipText = ""
        End If
        If .blnHasVWebLink Then
            lbl_Info_VWebLink.ForeColor = lblWebLink.ForeColor
        Else
            lbl_Info_VWebLink.ForeColor = lbl_Info_Publisher.ForeColor
        End If

        lbl_Info_Playable = .strPlayable
        lbl_Info_Original = .strOriginal
        lbl_Info_TitleScreen = .strTitleScreen
        lbl_Info_Trainers = .strTrainers
        lbl_Info_HSS = .strHighScoreSaver
        lbl_Info_GameLength = .strGameLength
        lbl_Info_TDE = .strTrueDriveEmulation
        lbl_Info_NTSC = .strNTSC
        lbl_Info_LoadingScreen = .strLoadingScreen
        lbl_Info_Docs = .strDocs

        lbl_Info_VComment = .strVersionComment
        .blnLoaded_1 = True
    End With

End Sub

Private Sub Show2()
    
    With g_udtGame
        lbl_Info_HighScore = .strHighScore
        lbl_Info_Difficulty = .strDifficulty
        lbl_Info_Rating = .strRating
        lbl_Info_TP = .strTimesPlayed
        lbl_Info_LP = .strDateLastPlayed
        .blnLoaded_2 = True
    End With

End Sub

Private Sub Show3()

    'load the memotext
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MemoText FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenSnapshot)
    g_udtGame.strMemo = g_rsGames!MemoText
    g_rsGames.Close
    Set g_rsGames = Nothing
    
    'validate it's length
    If Len(g_udtGame.strMemo) > gc_lngMAXMEMOTEXTLENGTH Then
        g_udtGame.strMemo = Left$(g_udtGame.strMemo, gc_lngMAXMEMOTEXTLENGTH)
    End If
    
    'show it
    txtMemo = g_udtGame.strMemo
    cmdMemoSave.Enabled = False
    cmdMemoClear.Enabled = (txtMemo <> "")
    g_udtGame.blnLoaded_3 = True

End Sub


Public Function Show4() As Boolean

    '-------------------------------------------------------
    'loads the Sid Information for a game in the Games table
    'and shows it in the tab.  Only loads and shows the info
    'if it hasn't done it already
    '
    'Returns:       True if the Sid-File exists
    '               False otherwise
    '-------------------------------------------------------

    Dim lngIdx As Long      'loop index
    Dim strTest As String

    'get the Sid file info
    With g_udtSid
    
        strTest = EllipseTextEnd(.strMusician, 2925)
        lbl_SFInfo_Musician = strTest
        If strTest <> .strMusician Then
            lbl_SFInfo_Musician.ToolTipText = .strMusician
        Else
            lbl_SFInfo_Musician.ToolTipText = ""
        End If

        lbl_SFInfo_Nick = .strNick
        lbl_SFInfo_Group = .strGroup
        lbl_SFInfo_Photo = .strPhoto
        If .blnIsPSID = True Then
            lblSidNotFound2.Visible = False
            fraSidInfo.Visible = True
            lbl_SFInfo_Name = .strSFName
            lbl_SFInfo_Author = .strSFAuthor
            lbl_SFInfo_Copyright = .strSFCopyright
            lbl_SFInfo_Songs = .lngSFTotalSongs
            lbl_SFInfo_Default = .lngSFDefaultSong
        Else
            lblSidNotFound2.Visible = False 'was true
            fraSidInfo.Visible = False
        End If
        .blnGotInfo = True
    End With

End Function

Private Sub ShowDTIcons()

    With g_udtGame
        
        'set starting position
        m_sngTabIconsX = ScaleWidth - (120 + imgAdult.Width)
           
        'classic g_udtGame icon
        If Not .blnClassic Then
            imgCup.Visible = False
            imgCup2.Visible = False
        Else
            With imgCup
                .Left = m_sngCup1X
                .Visible = True
            End With
            With imgCup2
                .Left = m_sngCup2X
                .Visible = True
            End With
        End If
    
        'adult/no adult
        If Not .blnAdult Then
            With imgAdult
                .Picture = imgNoAdultStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
            End With
        Else
            With imgAdult
                .Picture = imgAdultStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgAdult.Width)
    
        'game rating
        If .strRating = ga_strRatingLookupTable(0) Then
            With imgRatingFace
                .Picture = imgNoRating.Picture
                .ToolTipText = g_udtGame.strRating
                .Left = m_sngTabIconsX
                .Visible = True
            End With
        Else
            With imgRatingFace
                .Picture = imgFace(g_udtGame.lngRating - 1).Picture
                .ToolTipText = g_udtGame.strRating
                .Left = m_sngTabIconsX
                .Visible = True
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgRatingFace.Width)

        'related g_udtGame icon
        If .lngRelated = 0 Then
            With imgRelated
                .Picture = imgNoRelated.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbDefault
                .Tag = 0
                .ToolTipText = m_strRelatedTTT
            End With
        Else
            With imgRelated
                .Picture = imgRelatedStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbCustom
                .Tag = 1
                .ToolTipText = m_strRelatedTTT & " """ & g_udtGame.strRelated & """"
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgRelated.Width)

        'cloneof icon
        If .lngCloneOf = 0 Then
            With imgCloneOf
                .Picture = imgNoCloneOf.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbDefault
                .Tag = 0
                .ToolTipText = m_strCloneofTTT & "..."
            End With
        Else
            With imgCloneOf
                .Picture = imgCloneOfStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbCustom
                .Tag = 1
                .ToolTipText = m_strCloneofTTT & " """ & g_udtGame.strCloneOf & """"
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgCloneOf.Width)

        'sequel icon
        If .lngSequel = 0 Then
            With imgSequel
                .Picture = imgNoSequel.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbDefault
                .Tag = 0
                .ToolTipText = m_strSequelTTT
            End With
        Else
            With imgSequel
                .Picture = imgSequelStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbCustom
                .Tag = 1
                .ToolTipText = m_strSequelTTT & " """ & g_udtGame.strSequel & """"
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - (m_sngTabIconsGap + imgSequel.Width)

        'Precursor icon
        If .lngPrequel = 0 Then
            With imgPrequel
                .Picture = imgNoPrequel.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbDefault
                .Tag = 0
                .ToolTipText = m_strPrequelTTT
            End With
        Else
            With imgPrequel
                .Picture = imgPrequelStore.Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbCustom
                .Tag = 1
                .ToolTipText = m_strPrequelTTT & " """ & g_udtGame.strPrequel & """"
            End With
        End If
        m_sngTabIconsX = m_sngTabIconsX - ((m_sngTabIconsGap * 10) + imgRRStore(51).Width)

        'Review Rating bar
        If .lngReviewRating < 0 Then
            With imgRRStore(51)
                .Picture = LoadPicture()
                .Left = m_sngTabIconsX
                .Visible = False
                .MousePointer = vbDefault
                .Tag = 0
                .ToolTipText = ""
            End With
        Else
            With imgRRStore(51)
                .Picture = imgRRStore(GetReviewRatingImgIdx(g_udtGame.lngReviewRating)).Picture
                .Left = m_sngTabIconsX
                .Visible = True
                .MousePointer = vbCustom
                .Tag = 1
                .ToolTipText = g_strField_ReviewRating & ": " & g_udtGame.lngReviewRating & "%"
            End With
        End If

    End With

End Sub
Public Sub ShowPhoto()

    Dim MaxWidth As Single
    Dim MaxHeight As Single
    Dim Percent As Single
    Dim NewVal As Single
    
    If mnuViewScreenshots(0).Checked = True Then
        imgPhoto.Visible = False
    Else
        If mnuViewScreenshots(1).Checked = True Then
            MaxWidth = ((g_lngSmallPictureWidthPixels - 4) * Screen.TwipsPerPixelX)
        Else
            MaxWidth = ((g_lngFullSizePictureWidthPixels - 4) * Screen.TwipsPerPixelX)
        End If
        With imgPhoto
            MaxHeight = lvMain(0).Height
            'scale picture if needed
            If (m_sngCachedPhotoHeight > MaxHeight) Or (m_sngCachedPhotoWidth > MaxWidth) Then
                .Stretch = True
                If (m_sngCachedPhotoHeight > MaxHeight) And (m_sngCachedPhotoWidth > MaxWidth) Then
                    If (m_sngCachedPhotoHeight - MaxHeight) > (m_sngCachedPhotoWidth - MaxWidth) Then
                        Percent = MaxHeight / m_sngCachedPhotoHeight
                        .Height = MaxHeight
                        .Width = .Width * Percent
                    Else
                        Percent = MaxWidth / m_sngCachedPhotoWidth
                        .Width = MaxWidth
                        .Height = .Height * Percent
                    End If
                ElseIf m_sngCachedPhotoHeight > MaxHeight Then
                    Percent = MaxHeight / m_sngCachedPhotoHeight
                    .Height = MaxHeight
                    .Width = .Width * Percent
                Else 'm_sngCachedPhotoWidth > MaxWidth
                    Percent = MaxWidth / m_sngCachedPhotoWidth
                    .Width = MaxWidth
                    .Height = .Height * Percent
                End If
            Else
                .Stretch = False
                .Height = m_sngCachedPhotoHeight
                .Width = m_sngCachedPhotoWidth
            End If
            'position it and show
            .Visible = False
            .Left = picSS.Left + ((MaxWidth) / 2) - (.Width / 2)
            .Top = (lvMain(0).Height / 2) - (.Height / 2)
            .Visible = True
        End With
    End If

End Sub



Private Sub ClearScreenshots()

    ''''''''''''''''''''''''''''''''''''''''
    'Unload all screenshots currently loaded
    ''''''''''''''''''''''''''''''''''''''''
 
    Dim lngScrnShotIdx As Long  'loop index

    'allowed to unload the imagebox controls?
    If m_blnScrnShotBugWorkaround = False Then
        For lngScrnShotIdx = m_lngScrnShotLoaded To 1 Step -1
            Unload picSSMain(lngScrnShotIdx)
        Next lngScrnShotIdx
        'all extra shots unloaded!
        m_lngScrnShotLoaded = 0
    Else
        'just hide them
        For lngScrnShotIdx = m_lngScrnShotLoaded To 1 Step -1
            picSSMain(lngScrnShotIdx).Visible = False
        Next lngScrnShotIdx
    End If
    
    'reset screenshot vars
    ReDim ma_strScrnShots(0)
    m_lngScrnShotCount = 0
    m_blnScrnShotNone = True
    m_lngScrnShotTop = 0

End Sub

Private Sub LoadScreenshots(strScreenshotFile_in As String)
    '--------------------------------------------------------------
    'Loads the screenshots into the image box control array ready
    'for the ShowScreenshots() procedure
    '
    'Inputs:    strScreenshotFile_in        The filename (no path)
    '                                       of the first screenshot
    '                                       in a set of screenshots
    '                                       to load. i.e. no _1, _2
    '                                       etc..
    '--------------------------------------------------------------

    Dim blnFoundScrnShot As Boolean     'flag to determine that a screenshot file exists
    Dim lngScrnShotIdx As Long          'loop index
    Dim lngOldLoaded As Long            'number of screenshots already "loaded"
    Dim strScrnShotFilename As String   'path\filename storage for each screenshot
    Dim strBaseFilename As String       'base screenshot filename (no _1.gif or _2.gif etc)
    Dim strExt As String                'screenshot filename extension
    Dim lngPathIdx As Long              'loop index
    Dim lngMaxSSImage As Long
    Dim a_lngExtraIDs() As Long

    blnFoundScrnShot = False
    m_lngScrnShotCount = -1
    m_lngScrnShotTop = 0
    lngOldLoaded = m_lngScrnShotLoaded
    lngMaxSSImage = -1

    'find base screenshot
    If g_lngTotalScrnShotPaths > 0 Then
        For lngScrnShotIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(lngScrnShotIdx).strWithBackSlash & strScreenshotFile_in, False, "") Then
                blnFoundScrnShot = True
                m_lngScrnShotCount = 0
                m_lngScrnShotTop = 0
                m_blnScrnShotNone = False
                ReDim ma_strScrnShots(0)
                ma_strScrnShots(0) = ga_udtRegPathScrnShot(lngScrnShotIdx).strWithBackSlash & strScreenshotFile_in
                lngMaxSSImage = 0
                Exit For
            End If
        Next lngScrnShotIdx
        If blnFoundScrnShot = True Then
            'processed base screenshot, now find and process others
            lngScrnShotIdx = 1
            strBaseFilename = Left$(strScreenshotFile_in, InStrRev(strScreenshotFile_in, ".") - 1)
            strExt = Right$(strScreenshotFile_in, Len(strScreenshotFile_in) - Len(strBaseFilename))
            Do 'loop until there are no more extra screenshots found in any screenshot path
                strScrnShotFilename = strBaseFilename & "_" & lngScrnShotIdx & strExt
                For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                    If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(lngPathIdx).strWithBackSlash & strScrnShotFilename, False, "") Then
                        'path 1
                        ReDim Preserve ma_strScrnShots(lngScrnShotIdx)
                        ma_strScrnShots(lngScrnShotIdx) = ga_udtRegPathScrnShot(lngPathIdx).strWithBackSlash & strScrnShotFilename
                        Exit For
                    Else
                        'no more found!
                        Exit Do
                    End If
                Next lngPathIdx
                m_lngScrnShotCount = lngScrnShotIdx
                lngScrnShotIdx = lngScrnShotIdx + 1
                lngMaxSSImage = lngScrnShotIdx
            Loop
        End If
    End If
    
    'read in the image extras?
    If g_blnLoadImageExtrasAsScreenshots = True Then
        ReDim a_lngExtraIDs(0)
        If LoadImageExtrasIntoArray(a_lngExtraIDs, ma_strScrnShots, m_lngScrnShotCount, g_udtGame.lngID) > 0 Then
            blnFoundScrnShot = True
            m_blnScrnShotNone = False
        End If
    End If

    'sort out the images!
    If blnFoundScrnShot = True Then
    
        m_lngScrnShotLoaded = 0
    
        If m_lngScrnShotCount < lngOldLoaded Then
            'unload the extra image boxes if we don't have to
            'worry about the VB control unload limitation
            If m_blnScrnShotBugWorkaround = False Then
                For lngScrnShotIdx = lngOldLoaded To (m_lngScrnShotCount + 1) Step -1
                    Unload picSSMain(lngScrnShotIdx)
                Next lngScrnShotIdx
                m_lngScrnShotLoaded = m_lngScrnShotCount
            Else
                'else hide them because we can't unload them
                For lngScrnShotIdx = lngOldLoaded To (m_lngScrnShotCount + 1) Step -1
                    picSSMain(lngScrnShotIdx).Visible = False
                Next lngScrnShotIdx
                m_lngScrnShotLoaded = lngOldLoaded
            End If
        ElseIf m_lngScrnShotCount > lngOldLoaded Then
            'load the extra image boxes
            For lngScrnShotIdx = lngOldLoaded + 1 To m_lngScrnShotCount
                Load picSSMain(lngScrnShotIdx)
            Next lngScrnShotIdx
            m_lngScrnShotLoaded = m_lngScrnShotCount
        
        Else
            m_lngScrnShotLoaded = m_lngScrnShotCount
            
        End If
        'finally load the screenshots into the image boxes
        For lngScrnShotIdx = 0 To m_lngScrnShotCount
            m_blnPicExists = True
            With picSSMain(lngScrnShotIdx)
                .Appearance = 1
                .BorderStyle = 1
                .BackColor = g_lngPictureBackColor
                .Visible = False
                .Top = lngScrnShotIdx * m_sngScrnShotHeight
                .Height = m_sngScrnShotHeight
                .Width = m_sngScrnShotWidth
                .Enabled = True
                If (lngScrnShotIdx < lngMaxSSImage) Or (g_blnLoadImageExtrasAsScreenshots = False) Then
                    .Tag = 0
                Else
                    .Tag = -a_lngExtraIDs(lngScrnShotIdx)
                End If
                Call LoadPictureFromFile(ma_strScrnShots(lngScrnShotIdx), picSSMain(lngScrnShotIdx))
            End With
        Next lngScrnShotIdx
    Else
        'screenshot not found in any screenshot path
        Call ClearScreenshots
    End If
    
End Sub
Private Sub LoadScreenshotsMAME(strScreenshotFile_in As String)

    '--------------------------------------------------------------
    'Loads the screenshots into the image box control array ready
    'for the ShowScreenshots() procedure. This searches all folders
    'for the same named screenshot (MAME-style) so that flyers
    'and cabinets etc.. can be viewed simultaneously.
    '
    'Inputs:    strScreenshotFile_in        The filename (no path)
    '                                       of the screenshot to
    '                                       show.
    '--------------------------------------------------------------

    Dim lngScrnShotIdx As Long          'loop index
    Dim lngOldLoaded As Long            'number of screenshots already "loaded"
    Dim a_lngTagPath() As Long          'to remember the paths of the pictures for double-click zoom
    Dim lngMaxSSImage As Long
    Dim a_lngExtraIDs() As Long

    m_lngScrnShotCount = 0
    lngMaxSSImage = 0
    lngOldLoaded = m_lngScrnShotLoaded

    If g_lngTotalScrnShotPaths > 0 Then
        'read in all existing pictures

        m_lngScrnShotCount = 0
        m_lngScrnShotTop = 0
        m_blnScrnShotNone = False
        For lngScrnShotIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(lngScrnShotIdx).strWithBackSlash & strScreenshotFile_in, False, "") Then
                ReDim Preserve ma_strScrnShots(m_lngScrnShotCount)
                ReDim Preserve a_lngTagPath(m_lngScrnShotCount)
                ma_strScrnShots(m_lngScrnShotCount) = ga_udtRegPathScrnShot(lngScrnShotIdx).strWithBackSlash & strScreenshotFile_in
                a_lngTagPath(m_lngScrnShotCount) = lngScrnShotIdx
                m_lngScrnShotCount = m_lngScrnShotCount + 1
                lngMaxSSImage = lngMaxSSImage + 1
            End If
        Next lngScrnShotIdx
    End If
    
    'read in the image extras?
    If g_blnLoadImageExtrasAsScreenshots Then
        m_lngScrnShotCount = m_lngScrnShotCount - 1 'must pass ubound
        ReDim a_lngExtraIDs(0)
        If LoadImageExtrasIntoArray(a_lngExtraIDs, ma_strScrnShots, m_lngScrnShotCount, g_udtGame.lngID) > 0 Then
            m_blnScrnShotNone = False
        End If
        m_lngScrnShotCount = m_lngScrnShotCount + 1
    End If

    'do we have a picture?
    If m_lngScrnShotCount > 0 Then
        'sort out the images!
        m_lngScrnShotLoaded = 0
        m_lngScrnShotCount = m_lngScrnShotCount - 1
        If m_lngScrnShotCount < lngOldLoaded Then
            'unload the extra image boxes if we don't have to
            'worry about the VB control unload limitation
            If m_blnScrnShotBugWorkaround = False Then
                For lngScrnShotIdx = lngOldLoaded To (m_lngScrnShotCount + 1) Step -1
                    Unload picSSMain(lngScrnShotIdx)
                Next lngScrnShotIdx
                m_lngScrnShotLoaded = m_lngScrnShotCount
            Else
                'else hide them because we can't unload them
                For lngScrnShotIdx = lngOldLoaded To (m_lngScrnShotCount + 1) Step -1
                    picSSMain(lngScrnShotIdx).Visible = False
                Next lngScrnShotIdx
                m_lngScrnShotLoaded = lngOldLoaded
            End If
        ElseIf m_lngScrnShotCount > lngOldLoaded Then
            'load the extra image boxes
            For lngScrnShotIdx = lngOldLoaded + 1 To m_lngScrnShotCount
                Load picSSMain(lngScrnShotIdx)
            Next lngScrnShotIdx
            m_lngScrnShotLoaded = m_lngScrnShotCount
        Else
            m_lngScrnShotLoaded = m_lngScrnShotCount
        End If
        'finally load the screenshots into the image boxes
        For lngScrnShotIdx = 0 To m_lngScrnShotCount
            m_blnPicExists = True
            With picSSMain(lngScrnShotIdx)
                .Appearance = 1
                .BorderStyle = 1
                .BackColor = g_lngPictureBackColor
                .Visible = False
                .Top = lngScrnShotIdx * m_sngScrnShotHeight
                .Height = m_sngScrnShotHeight
                .Width = m_sngScrnShotWidth
                .Enabled = True
                If lngScrnShotIdx < lngMaxSSImage Then
                    .Tag = a_lngTagPath(lngScrnShotIdx)
                ElseIf g_blnLoadImageExtrasAsScreenshots = True Then
                    .Tag = -a_lngExtraIDs(lngScrnShotIdx)
                Else
                    .Tag = 0
                End If
                Call LoadPictureFromFile(ma_strScrnShots(lngScrnShotIdx), picSSMain(lngScrnShotIdx))
            End With
        Next lngScrnShotIdx
    Else
        'no pictures
        Call ClearScreenshots
    End If

End Sub

Private Sub CalculateMinWindowSizes()

    '----------------------------------------------------------------
    'Set the minimum window sizes for different window configurations
    'remember user could be in large or small fonts
    '----------------------------------------------------------------
    
    Dim ExtraTwips As Single
    Dim TempTwips As Single
    
    'get picture sizes from settings.ini
    g_lngFullSizePictureWidthPixels = Val(g_xobjFunc.INI_GetSetting("UI", "FullSizePictureWidthPixels", Trim$(gc_lngSCRNSHOTWIDTH), g_strSettingsFile))
    g_lngFullSizePictureHeightPixels = Val(g_xobjFunc.INI_GetSetting("UI", "FullSizePictureHeightPixels", Trim$(gc_lngSCRNSHOTHEIGHT), g_strSettingsFile))
    g_lngSmallPictureWidthPixels = Val(g_xobjFunc.INI_GetSetting("UI", "SmallPictureWidthPixels", Trim$(gc_lngSCRNSHOTWIDTHHALF), g_strSettingsFile))
    g_lngSmallPictureHeightPixels = Val(g_xobjFunc.INI_GetSetting("UI", "SmallPictureHeightPixels", Trim$(gc_lngSCRNSHOTHEIGHTHALF), g_strSettingsFile))
    g_lngZoomedPictureWidthPixels = Val(g_xobjFunc.INI_GetSetting("UI", "ZoomedWidthPixels", Trim$(gc_lngSCRNSHOTWIDTHx2), g_strSettingsFile))
    g_lngZoomedPictureHeightPixels = Val(g_xobjFunc.INI_GetSetting("UI", "ZoomedPictureHeightPixels", Trim$(gc_lngSCRNSHOTHEIGHTx2), g_strSettingsFile))
    If g_lngFullSizePictureWidthPixels < 1 Then g_lngFullSizePictureWidthPixels = 1
    If g_lngFullSizePictureHeightPixels < 1 Then g_lngFullSizePictureHeightPixels = 1
    If g_lngSmallPictureWidthPixels < 1 Then g_lngSmallPictureWidthPixels = 1
    If g_lngSmallPictureHeightPixels < 1 Then g_lngSmallPictureHeightPixels = 1
    If g_lngZoomedPictureWidthPixels < 1 Then g_lngZoomedPictureWidthPixels = 1
    If g_lngZoomedPictureHeightPixels < 1 Then g_lngZoomedPictureHeightPixels = 1

    With m_udtMinSizes
        'width
        .lngTwipsX = Screen.TwipsPerPixelX
        .sngScaleWidth_Twips = ScaleWidth
        .lngScaleWidth_Pixels = .sngScaleWidth_Twips / .lngTwipsX
        .sngWidth_Twips = ScaleWidth + (Width - ScaleWidth)
        .lngWidth_Pixels = .sngWidth_Twips / .lngTwipsX
        .lngFullSSWidth_Pixels = g_lngFullSizePictureWidthPixels + g_lngScrnShotExtraPixelsX
        .sngFullSSWidth_Twips = (g_lngFullSizePictureWidthPixels * .lngTwipsX) + g_sngScrnShotExtraTwipsX
        .lngSmallSSWidth_Pixels = g_lngSmallPictureWidthPixels + g_lngScrnShotExtraPixelsX
        .sngSmallSSWidth_Twips = (g_lngSmallPictureWidthPixels * .lngTwipsX) + g_sngScrnShotExtraTwipsX
        .lngScreenWidth_Pixels = Screen.Width / .lngTwipsX
        'height
        .lngTwipsY = Screen.TwipsPerPixelY
        .lngFullSSHeight_Pixels = g_lngFullSizePictureHeightPixels + g_lngScrnShotExtraPixelsY
        .sngFullSSHeight_Twips = (g_lngFullSizePictureHeightPixels * .lngTwipsY) + g_sngScrnShotExtraTwipsY
        .lngSmallSSHeight_Pixels = g_lngSmallPictureHeightPixels + g_lngScrnShotExtraPixelsY
        .sngSmallSSHeight_Twips = (g_lngSmallPictureHeightPixels * .lngTwipsY) + g_sngScrnShotExtraTwipsY
        ExtraTwips = Height - ScaleHeight
        .lngScreenHeight_Pixels = Screen.Height / .lngTwipsY
        'get height with details less screenshot
        TempTwips = sbStatus.Height
        TempTwips = TempTwips + mc_TwipsBetweenTABDETAILSAndSBSTATUS
        TempTwips = TempTwips + tabDetails.Height
        TempTwips = TempTwips + mc_TwipsBetweenFRAPANELAndTABDETAILS
        TempTwips = TempTwips + fraPanel.Height
        TempTwips = TempTwips + mc_TwipsBetweenLVMAINAndFRAPANEL
        .sngScaleHeightLessSSDetails_Twips = TempTwips
        'get height without details less screenshot
        TempTwips = sbStatus.Height
        TempTwips = TempTwips + mc_TwipsBetweenFRAPANELAndSBSTATUS
        TempTwips = TempTwips + fraPanel.Height
        TempTwips = TempTwips + mc_TwipsBetweenLVMAINAndFRAPANEL
        .sngScaleHeightLessSSNoDetails_Twips = TempTwips
        'big screenshot with details
        .sngHeight_FullSSDetails_Twips = .sngFullSSHeight_Twips + .sngScaleHeightLessSSDetails_Twips + ExtraTwips
        .sngScaleHeight_FullSSDetails_Twips = .sngHeight_FullSSDetails_Twips - ExtraTwips
        .lngHeight_FullSSDetails_Pixels = .sngHeight_FullSSDetails_Twips / .lngTwipsY
        'small screenshot with details
        .sngHeight_SmallSSDetails_Twips = .sngSmallSSHeight_Twips + .sngScaleHeightLessSSDetails_Twips + ExtraTwips
        .sngScaleHeight_SmallSSDetails_Twips = .sngHeight_SmallSSDetails_Twips - ExtraTwips
        .lngHeight_SmallSSDetails_Pixels = .sngHeight_SmallSSDetails_Twips / .lngTwipsY
        'big screenshots no details
        .sngHeight_FullSSNoDetails_Twips = .sngFullSSHeight_Twips + .sngScaleHeightLessSSNoDetails_Twips + ExtraTwips
        .sngScaleHeight_FullSSNoDetails_Twips = .sngHeight_FullSSNoDetails_Twips - ExtraTwips
        .lngHeight_FullSSNoDetails_Pixels = .sngHeight_FullSSNoDetails_Twips / .lngTwipsY
        'small screenshot no details
        .sngHeight_SmallSSNoDetails_Twips = .sngSmallSSHeight_Twips + .sngScaleHeightLessSSNoDetails_Twips + ExtraTwips
        .sngScaleHeight_SmallSSNoDetails_Twips = .sngHeight_SmallSSNoDetails_Twips - ExtraTwips
        .lngHeight_SmallSSNoDetails_Pixels = .sngHeight_SmallSSNoDetails_Twips / .lngTwipsY
    End With

End Sub

Private Sub SetScreenshotArrows(ByVal lngImageBoxIdx_in As Integer, ByVal sngMousePosX_in As Single, ByVal sngMousePosY_in As Single, lngShift_in As Integer)

    Dim blnUP As Boolean
    Dim blnDOWN As Boolean
    Dim sngHalfway As Single
    Dim sngCurrent As Single
    
    blnUP = picSSUp.Visible
    blnDOWN = picSSDown.Visible
    
    'calculate if up arrow or down arrow
    sngHalfway = (m_sngScrnShotHeight * (m_lngScrnShotMaxVisible + 1)) / 2 'm_lngScrnShotMaxVisible starts at 0 remember
    sngCurrent = (m_sngScrnShotHeight * (lngImageBoxIdx_in - m_lngScrnShotTop)) + (sngMousePosY_in * Screen.TwipsPerPixelY)

    With picSSMain(lngImageBoxIdx_in)
        If (sngMousePosX_in * Screen.TwipsPerPixelX) > (m_sngScrnShotWidth / 2) Then
            If blnDOWN And Not blnUP Then
                'only down arrow for next screenshot
                If sngCurrent < sngHalfway Then
                    .MousePointer = vbNormal
                Else
                    .MousePointer = vbCustom
                    .MouseIcon = imgCurDown.MouseIcon
                End If
            
            ElseIf blnUP And Not blnDOWN Then
                'only up arrow for previous screenshot
                If sngCurrent < sngHalfway Then
                    .MousePointer = vbCustom
                    .MouseIcon = imgCurUp.MouseIcon
                Else
                    .MousePointer = vbNormal
                End If
            
            ElseIf blnUP And blnDOWN Then
                'up and down arrows for both next and previous screenshots
                If sngCurrent < sngHalfway Then
                    .MousePointer = vbCustom
                    .MouseIcon = imgCurUp.MouseIcon
                Else
                    .MousePointer = vbCustom
                    .MouseIcon = imgCurDown.MouseIcon
                End If
            Else
                'no arrows - no extra screenshots
                .MousePointer = vbNormal
                
            End If
        ElseIf m_blnScrnShotNone = False Then
            .MousePointer = vbCustom
            .MouseIcon = imgZoomCur.MouseIcon
            
        End If
    End With
    
End Sub

Private Sub SetExtraColumns(ByVal lngViewIdx_in As Long)

    '------------------------------------------------------------
    'Sets the extra columns properties for the view engine to use
    'when reading in the list query from the database and sets
    'all the check-marks and enabled properties for the menus
    '
    'Inputs:    lngViewIdx_in       The index of the current view
    '------------------------------------------------------------
    
    Dim lngMenuIdx As Long          'the second-column menu index
    Dim lngFieldIdx As Long         'the second-column field index
    Dim blnEnableMenu As Boolean    'flag to enable the parent menu
    Dim lngExtraColumnIdx As Long
    Dim blnOK As Boolean
    Dim strFields As String
    
    With ga_udtViewData(lngViewIdx_in)
    
        If .lngExtraColumnsCount = 0 Then
            'no extra column
            With g_udtExtraColumnFields
                .strFields = ""
                .strMusicFields = ""
            End With
        Else
            'check all columns are ok
            blnOK = True
            For lngExtraColumnIdx = 0 To .lngMaxExtraColumns
                With .a_udtExtraColumns(lngExtraColumnIdx)
                    If (ga_udtViewData(lngViewIdx_in).eIncludeMusicTable = e_Table_SidOnly) _
                     And (ga_udtExtraColumns(.lngTypeIndex).a_udtColumnItems(.lngFieldIndex).blnGameOnlyTable) Then
                        blnOK = False
                        Exit For
                     End If
                End With
            Next lngExtraColumnIdx
            If Not blnOK Then
                With g_udtExtraColumnFields
                    .strFields = ""
                    .strMusicFields = ""
                End With
            Else
                'generate the field strings for the extra columns
                g_udtExtraColumnFields.strFields = ""
                g_udtExtraColumnFields.strMusicFields = ""
                For lngExtraColumnIdx = 0 To .lngMaxExtraColumns
                    With .a_udtExtraColumns(lngExtraColumnIdx)
                        'game fields
                        strFields = ga_udtExtraColumns(.lngTypeIndex).a_udtColumnItems(.lngFieldIndex).strFields
                        strFields = Replace(strFields, "ColX", "Col" & (lngExtraColumnIdx + 2))
                        g_udtExtraColumnFields.strFields = g_udtExtraColumnFields.strFields & strFields & ", "
                        'music fields
                        strFields = ga_udtExtraColumns(.lngTypeIndex).a_udtColumnItems(.lngFieldIndex).strMusicFields
                        If strFields <> "" Then
                            strFields = Replace(strFields, "ColX", "Col" & (lngExtraColumnIdx + 2))
                            g_udtExtraColumnFields.strMusicFields = g_udtExtraColumnFields.strMusicFields & strFields & ", "
                        End If
                    End With
                Next lngExtraColumnIdx
                If Right$(g_udtExtraColumnFields.strFields, 2) = ", " Then
                    g_udtExtraColumnFields.strFields = Left$(g_udtExtraColumnFields.strFields, Len(g_udtExtraColumnFields.strFields) - 2)
                End If
                If Right$(g_udtExtraColumnFields.strMusicFields, 2) = ", " Then
                    g_udtExtraColumnFields.strMusicFields = Left$(g_udtExtraColumnFields.strMusicFields, Len(g_udtExtraColumnFields.strMusicFields) - 2)
                End If
            End If
        End If
    End With

End Sub
Private Sub LoadViewPrefs(ByVal lngViewIdx_in As Long, lngListViewIdx_in As Long)

    '---------------------------------------------------------------------------
    'Loads the custom view preferences into the listview from the
    'custom views array in memory.
    '
    'Inputs:        lngViewIdx_in       the actual custom view index
    '               lngListViewIdx_in   the view's corresponding list view index
    '
    'Remarks:       The sorting is not loaded and done here because the
    '               listview has to be filled first, it's done in
    '               ReadViewIntoList().  Also, the selecteditem is not done
    '               here for the same reason.
    '---------------------------------------------------------------------------
    
    'set the listview prefs for the selected view
    With ga_udtViewData(lngViewIdx_in)
        lvMain(lngListViewIdx_in).View = .lngListViewType
        Call SetExtraColumns(lngViewIdx_in)
        lvMain(lngListViewIdx_in).Sorted = False
    End With

End Sub



Private Sub SaveViewPrefs(ByVal lngViewIdx_in As Long, lngListViewIdx_in As Long)

    '---------------------------------------------------------------------------
    'Saves the preferences for the current view to the custom view array in
    'memory and to the database/registry when the user moves to a different
    'view or quits the program.
    '
    'Inputs:        lngViewIdx_in       the actual custom view index
    '               lngListViewIdx_in   the view's corresponding list view index
    '
    'Remarks:       Will save only the custom-view prefs to the database.
    '               Prefs for the default views and any loaded Quick Filter view
    '               will be saved to the registry.
    '               The Second Column info is saved to the views array when the
    '               user selects it from the menus, so we only need to save it
    '               to the database/registry here.
    '---------------------------------------------------------------------------

    Dim lngColumnIdx As Long
    Dim strKeyName As String

    'save prefs to the custom views array in memory
    With ga_udtViewData(lngViewIdx_in)
        .lngListViewType = lvMain(lngListViewIdx_in).View
        .lngSortColumn = lvMain(lngListViewIdx_in).SortKey
        .lngSortOrder = lvMain(lngListViewIdx_in).SortOrder
        If .lngListViewType = lvwReport Then
            .lngCol1Width = lvMain(lngListViewIdx_in).ColumnHeaders(1).Width
            If lvMain(lngListViewIdx_in).ColumnHeaders.Count > 1 Then
                For lngColumnIdx = 0 To .lngExtraColumnsCount - 1
                    .a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth _
                        = lvMain(lngListViewIdx_in).ColumnHeaders(lngColumnIdx + 2).Width
                Next lngColumnIdx
            End If
        End If
        If lvMain(lngListViewIdx_in).SelectedItem Is Nothing Then
            .strSelectedItem = "INVALID"
        Else
            .strSelectedItem = lvMain(lngListViewIdx_in).SelectedItem.Key
        End If
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'save to the registry/database depending on
    'whether the view is one of the defaults or not
    '''''''''''''''''''''''''''''''''''''''''''''''

    If lngViewIdx_in <= gc_lngDEFAULTVIEWSMAX Then
        'Default view: save prefs to the registry
        strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\DefaultViews"
        With g_xobjRegistry
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "ExtraColumns_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).strExtraColumns, True
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "Col1Width_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).lngCol1Width, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "ListViewType_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).lngListViewType, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SortColumn_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).lngSortColumn, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SortOrder_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).lngSortOrder, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SelectedItem_" & lngViewIdx_in, ga_udtViewData(lngViewIdx_in).strSelectedItem, True
            For lngColumnIdx = 0 To ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1
                g_xobjRegistry.SaveSetting _
                    "HKEY_CURRENT_USER", _
                    strKeyName, _
                    "Col" & (lngColumnIdx + 2) & "Width_" & lngViewIdx_in, _
                    ga_udtViewData(lngViewIdx_in).a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                    False
            Next lngColumnIdx
        End With
    ElseIf (lngViewIdx_in = UBound(ga_udtViewData)) And (m_blnLastViewIsQuick = True) Then
        'quick filter view
        strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter"
        With g_xobjRegistry
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "ExtraColumns", ga_udtViewData(lngViewIdx_in).strExtraColumns, True
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "Col1Width", ga_udtViewData(lngViewIdx_in).lngCol1Width, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "ListViewType", ga_udtViewData(lngViewIdx_in).lngListViewType, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SortColumn", ga_udtViewData(lngViewIdx_in).lngSortColumn, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SortOrder", ga_udtViewData(lngViewIdx_in).lngSortOrder, False
            .SaveSetting "HKEY_CURRENT_USER", strKeyName, "SelectedItem", ga_udtViewData(lngViewIdx_in).strSelectedItem, True
            For lngColumnIdx = 0 To ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1
                g_xobjRegistry.SaveSetting _
                    "HKEY_CURRENT_USER", _
                    strKeyName, _
                    "Col" & (lngColumnIdx + 2) & "Width", _
                    ga_udtViewData(lngViewIdx_in).a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                    False
            Next lngColumnIdx
        End With
    Else
        'Custom view: save prefs to the database
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM ViewData WHERE ID = " & ga_udtViewData(lngViewIdx_in).lngID, dbOpenDynaset)
        With g_rsGames
            .MoveFirst
            .Edit
            !ListViewType = ga_udtViewData(lngViewIdx_in).lngListViewType
            !SortColumn = ga_udtViewData(lngViewIdx_in).lngSortColumn
            !SortOrder = ga_udtViewData(lngViewIdx_in).lngSortOrder
            !SelectedItem = ga_udtViewData(lngViewIdx_in).strSelectedItem
            !ExtraColumns = ga_udtViewData(lngViewIdx_in).strExtraColumns
            .Update
            .Close
        End With
        Set g_rsGames = Nothing

        'column widths to the registry
        strKeyName = "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\CustomViews"
        g_xobjRegistry.SaveSetting _
            "HKEY_CURRENT_USER", _
            strKeyName, _
            "Col1Width_" & ga_udtViewData(lngViewIdx_in).lngID, _
            ga_udtViewData(lngViewIdx_in).lngCol1Width, _
            False

        For lngColumnIdx = 0 To ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1
            g_xobjRegistry.SaveSetting _
                "HKEY_CURRENT_USER", _
                strKeyName, _
                "Col" & (lngColumnIdx + 2) & "Width_" & ga_udtViewData(lngViewIdx_in).lngID, _
                ga_udtViewData(lngViewIdx_in).a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth, _
                False
        Next lngColumnIdx
    End If

End Sub


Private Function SetView(ByVal lngNewViewIdx_in As Long) As Boolean

    '-------------------------------------------------------------------------
    'Saves the current view's preferences, loads the new view's preferences,
    'builds the new view's SQL query and loads the list of games/sids from the
    'database.
    '
    'Inputs:        lngNewViewIdx_in    the custom view index for the new view
    '
    'Returns:       True if the view has at least one item
    '               False if the view is empty
    '-------------------------------------------------------------------------
    
    Dim lngListViewIdx As Long
    Dim lngIdx As Long
    Dim blnListViewHasGames As Boolean

    'set the caption
    Caption = g_strWINDOWCAPTION

    'save the current view
    If (m_blnFormIsLoading = False) And (m_blnSaveViewPrefs = True) Then
        Call SaveViewPrefs(m_lngLastViewIdx, GetListViewIdxFromViewIdx(m_lngLastViewIdx))
    End If
    
    'store the new view index for next time the view changes
    m_lngLastViewIdx = lngNewViewIdx_in
    
    'get which listview control to use
    lngListViewIdx = GetListViewIdxFromViewIdx(lngNewViewIdx_in)
    
    'disable controls as change the mousepointer
    MousePointer = vbHourglass
    If g_blnSettingCatFilter = False Then
        DoEvents
    End If
    
    'load the new view's prefs for the listview
    Call LoadViewPrefs(lngNewViewIdx_in, lngListViewIdx)
    
    'make the view's SQL string if the view is not already loaded
    'and read the games/sids into the listview
    If ga_blnViewsLoaded(lngListViewIdx) = False Then
        Call ReadViewIntoList(lngNewViewIdx_in, lngListViewIdx)
    End If

    'adjust menus
    With ga_udtViewData(lngNewViewIdx_in)
        'the View menu
        mnuViewListList.Checked = (.lngListViewType = lvwList)
        mnuViewListDetails.Checked = (.lngListViewType = lvwReport)
    End With
   
    'hide and disable all non-relevant listview controls
    'and show the relevant one
    For lngIdx = 0 To (gc_lngDEFAULTVIEWSMAX + 1)
        With lvMain(lngIdx)
            If lngListViewIdx = lngIdx Then
                .Visible = True
                .Enabled = True
            Else
                .Visible = False
                .Enabled = False
            End If
        End With
    Next lngIdx

    'disable/enable controls depending on if the view has items in the listview or not
    SetView = (lvMain(lngListViewIdx).ListItems.Count > 0)
    Call SetControlsForView(SetView, lngListViewIdx)
    
    'set controls depending on what view we have
    If lngNewViewIdx_in = gc_lngFAVOURITESINDEX Then
        'favourites menu item
        mnuEditFav.Caption = m_strRemoveFromFavs
    Else
        mnuEditFav.Caption = m_strAddToFavs
    End If
    
    'save the new view index to the registry
    If m_blnFormIsLoading = False Then
        g_lngRegViewIndex = lngNewViewIdx_in
        g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "ViewIndex", g_lngRegViewIndex, False
    End If

    'enable all the controls and change the mousepointer
    MousePointer = vbNormal

End Function
Private Function ReadViewIntoList(ByVal lngViewIdx_in As Long, ByVal lngListViewIdx_in As Long) As Boolean

    Dim fldName As Field
    Dim fldID As Field
    Dim fldCCode As Field
    Dim fldColX As Field
    Dim fldColXExtra As Field
    Dim fldColXExtra2 As Field
    Dim fldCol() As Field
    Dim fldColExtra() As Field
    Dim fldColExtra2() As Field
    Dim fldBold1 As Field
    Dim fldBold2 As Field
    Dim LItem As ListItem
    Dim lngTotal As Long
    Dim intCCode As Integer
    Dim strCount As String
    Dim a_strFilterBuffers() As String
    Dim strQuery As String
    Dim intLengthType As Integer
    Dim lngTempIdx As Long
    Dim lngExtraColumnIdx As Long
    Dim udtView As ViewDataType
    Dim lngTypeIdx As Long
    Dim lngFieldIdx As Long
    Dim lngSubItemIdx As Long

    'debug code
    g_xobjClock.StartClock
    
    '''''''''''''''''''''
    'PREPARE THE LISTVIEW
    '''''''''''''''''''''

    'column headers, sort order, etc
    With lvMain(lngListViewIdx_in)
        
        .Visible = False
        .Enabled = False
        
        'clear the listview:
        '(don't use lvmain(lngViewIdx_in).Clear because it takes ages!)
        Call SendMessageAny(.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
        
        'init column numero uno
        .ColumnHeaders(1).Text = ga_udtViewData(lngViewIdx_in).strTitle
        .ColumnHeaders(1).Width = ga_udtViewData(lngViewIdx_in).lngCol1Width
        
        'start with a single column
        For lngTempIdx = .ColumnHeaders.Count To 2 Step -1
            .ColumnHeaders.Remove lngTempIdx
        Next lngTempIdx

        'add extra columns
        With ga_udtViewData(lngViewIdx_in)
            For lngTempIdx = 0 To .lngExtraColumnsCount - 1
                lvMain(lngListViewIdx_in).ColumnHeaders.Add , , _
                    ga_udtExtraColumns(.a_udtExtraColumns(lngTempIdx).lngTypeIndex).a_udtColumnItems(.a_udtExtraColumns(lngTempIdx).lngFieldIndex).strColumnTitle, _
                        .a_udtExtraColumns(lngTempIdx).lngColumnHeaderWidth
            Next lngTempIdx
        End With

        'set the sort column and order and load the sort icons
        If ga_udtViewData(lngViewIdx_in).lngSortColumn > .ColumnHeaders.Count - 1 Then
            ga_udtViewData(lngViewIdx_in).lngSortColumn = .ColumnHeaders.Count - 1
        End If
        .SortKey = ga_udtViewData(lngViewIdx_in).lngSortColumn
        .SortOrder = ga_udtViewData(lngViewIdx_in).lngSortOrder
        .Sorted = True
        SetSortIcons lngListViewIdx_in

    End With

    '''''''''''''''
    ' GAMES TABLE '
    '''''''''''''''

    'Read the list from the SQL string into memory then load the entries into the listbox
    If ga_udtViewData(lngViewIdx_in).eIncludeMusicTable <> e_Table_SidOnly Then
        With ga_udtViewData(lngViewIdx_in)
            Call MakeGB64SQLFilterString(.a_udtFilters(), .lngFilterCount, a_strFilterBuffers())
            strQuery = m_xobjSQLEngine.GetViewQueryString(True, g_udtExtraColumnFields.strFields, .lngFilterCount, a_strFilterBuffers(), (.eFilterMode = e_FilterMode_AND), g_blnRegHideAdultItems)
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Info", "LastGamesSQL", strQuery, True
            If g_blnInVBIDE = True Then
                Debug.Print "(Games) " & strQuery
            End If
        End With
        Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
        If Not g_rsGames.EOF Then
            g_rsGames.MoveLast
            g_rsGames.MoveFirst
            lngTotal = g_rsGames.RecordCount
            If lngTotal > gc_lngACCEPTABLEWAIT And m_blnFormIsLoading = False Then
                'show the wait message
                With picPleaseWait
                    .Top = 0
                    .Left = 0
                    .Width = lvMain(lngListViewIdx_in).Width
                    .Height = lvMain(lngListViewIdx_in).Height
                    lblPleaseWait.Top = (.Height / 2) - (lblPleaseWait.Height / 2)
                    lblPleaseWait.Left = 0
                    lblPleaseWait.Width = .Width
                    .Visible = True
                End With
                If g_blnSettingCatFilter = False Then
                    DoEvents
                End If
            End If
            'setup field objects for quick access
            Set fldName = g_rsGames!Name
            Set fldID = g_rsGames!GA_Id
            Set fldCCode = g_rsGames!CCode
            Set fldBold1 = g_rsGames!Bold1
            Set fldBold2 = g_rsGames!Bold2
            If ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount > 0 Then
                ReDim Preserve fldCol(ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1)
                ReDim Preserve fldColExtra(ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1)
                ReDim Preserve fldColExtra2(ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1)
                On Error Resume Next
                For lngExtraColumnIdx = 0 To ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1
                    Set fldCol(lngExtraColumnIdx) = g_rsGames.Fields("Col" & lngExtraColumnIdx + 2)
                    Set fldColExtra(lngExtraColumnIdx) = g_rsGames.Fields("Col" & lngExtraColumnIdx + 2 & "Extra")
                    Set fldColExtra2(lngExtraColumnIdx) = g_rsGames.Fields("Col" & lngExtraColumnIdx + 2 & "Extra2")
                Next lngExtraColumnIdx
                On Error GoTo 0
            End If
            
            'FOUND GAMES - FILL LISTVIEW
            If ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount = 0 Then
                'Only One Column
                With lvMain(lngListViewIdx_in).ListItems
                    Do Until g_rsGames.EOF
                        intCCode = CInt(fldCCode)
                        Set LItem = .Add(, "G" & fldID, fldName, intCCode, intCCode)
                        LItem.Bold = Abs(fldBold1 Or fldBold2)
                        Set LItem = Nothing
                        g_rsGames.MoveNext
                    Loop
                End With
            Else
                'Multiple Columns
                udtView = ga_udtViewData(lngViewIdx_in)
                Do Until g_rsGames.EOF
                    intCCode = CInt(fldCCode)
                    Set LItem = lvMain(lngListViewIdx_in).ListItems.Add(, "G" & fldID, fldName, intCCode, intCCode)
                    LItem.Bold = Abs(fldBold1 Or fldBold2)
                    For lngExtraColumnIdx = 0 To udtView.lngExtraColumnsCount - 1

                        lngSubItemIdx = lngExtraColumnIdx + 1
                        lngTypeIdx = udtView.a_udtExtraColumns(lngExtraColumnIdx).lngTypeIndex
                        lngFieldIdx = udtView.a_udtExtraColumns(lngExtraColumnIdx).lngFieldIndex

                        Set fldColX = fldCol(lngExtraColumnIdx)
                        Set fldColXExtra = fldColExtra(lngExtraColumnIdx)
                        Set fldColXExtra2 = fldColExtra2(lngExtraColumnIdx)

                        With LItem
                            Select Case ga_udtExtraColumns(lngTypeIdx).a_udtColumnItems(lngFieldIdx).eType
                                
                            Case e_SecColType_Normal
                                .SubItems(lngSubItemIdx) = fldColX
                            
                            Case e_SecColType_Publisher
                                .SubItems(lngSubItemIdx) = fldColX
                            
                            Case e_SecColType_Year
                                .SubItems(lngSubItemIdx) = ga_strYearLookupTable(fldColX)

                            Case e_SecColType_Genres
                                .SubItems(lngSubItemIdx) = ga_strGenreParentLookupTable(fldColXExtra) & fldColX
                            
                            Case e_SecColType_YesNo
                                .SubItems(lngSubItemIdx) = ga_strYesNoUnknownLookupTable(Abs(fldColX))
                            
                            Case e_SecColType_Players
                                If fldColX <> fldColXExtra Then
                                    If fldColX = gc_lngPLAYERSUNKNOWN Then
                                        .SubItems(lngSubItemIdx) = "? - " & fldColXExtra & ga_strPlayerSimLookupTable(Abs(fldColXExtra2))
                                    ElseIf fldColXExtra = gc_lngPLAYERSUNKNOWN Then
                                        .SubItems(lngSubItemIdx) = fldColX & " - ?" & ga_strPlayerSimLookupTable(Abs(fldColXExtra2))
                                    Else
                                        .SubItems(lngSubItemIdx) = fldColX & " - " & fldColXExtra & ga_strPlayerSimLookupTable(Abs(fldColXExtra2))
                                    End If
                                Else
                                    If fldColX > gc_lngPLAYERSUNKNOWN Then
                                        .SubItems(lngSubItemIdx) = fldColX & "P Only" & ga_strPlayerSimLookupTable(Abs(fldColXExtra2))
                                    Else
                                        .SubItems(lngSubItemIdx) = "(" & g_strUnknown & ")"
                                    End If
                                End If

                            Case e_SecColType_Rating
                                If fldBold2 = False Then
                                    .SubItems(lngSubItemIdx) = ga_strRatingLookupTable(fldColX)
                                Else
                                    .SubItems(lngSubItemIdx) = ga_strRatingLookupTable(6)
                                End If

                            Case e_SecColType_Trainers
                                If fldColX > gc_lngTRAINERSUNKNOWN Then
                                    .SubItems(lngSubItemIdx) = fldColX
                                Else
                                    .SubItems(lngSubItemIdx) = "(" & g_strUnknown & ")"
                                End If

                            Case e_SecColType_ReviewRating
                                If fldColX > gc_lngTRAINERSUNKNOWN Then
                                    .SubItems(lngSubItemIdx) = fldColX & "%"
                                Else
                                    .SubItems(lngSubItemIdx) = "(" & g_strUnknown & ")"
                                End If

                            Case e_SecColType_Control
                                .SubItems(lngSubItemIdx) = ga_strControlLookupTable(fldColX)
                            
                            Case e_SecColType_PALNTSC
                                .SubItems(lngSubItemIdx) = ga_strPalNTSCLookupTable(fldColX)
                            
                            Case Else 'e_SecColType_Length
                                If fldColXExtra < 19 And fldColXExtra > -1 Then
                                    .SubItems(lngSubItemIdx) = fldColX & " " & ga_strLengthTypeLookupTable(intLengthType)
                                Else
                                    .SubItems(lngSubItemIdx) = fldColX & " " & ga_strLengthTypeLookupTable(19)
                                End If
                            
                            End Select
                        End With

                    Next lngExtraColumnIdx
                    Set LItem = Nothing
                    g_rsGames.MoveNext
                Loop
            End If
            'release quick-access field objects
            Set fldName = Nothing
            Set fldID = Nothing
            Set fldCCode = Nothing
            Set fldBold1 = Nothing
            Set fldBold2 = Nothing
            For lngExtraColumnIdx = 0 To udtView.lngExtraColumnsCount - 1
                Set fldCol(lngExtraColumnIdx) = Nothing
                Set fldColExtra(lngExtraColumnIdx) = Nothing
                Set fldColExtra2(lngExtraColumnIdx) = Nothing
            Next lngExtraColumnIdx
        Else
            'NO GAMES
            lngTotal = 0
        End If
        g_rsGames.Close
        Set g_rsGames = Nothing
    Else
        'NO GAMES
        lngTotal = 0
    End If

    '''''''''''''''
    ' MUSIC TABLE '
    '''''''''''''''

    'If there's the Music Table to read in, now's the time to do it!
    If ga_udtViewData(lngViewIdx_in).eIncludeMusicTable <> e_Table_GamesOnly Then
        With ga_udtViewData(lngViewIdx_in)
            Call MakeGB64SQLFilterString(.a_udtFilters(), .lngFilterCount, a_strFilterBuffers())
            strQuery = m_xobjSQLEngine.GetViewQueryString(False, g_udtExtraColumnFields.strMusicFields, .lngFilterCount, a_strFilterBuffers(), (.eFilterMode = e_FilterMode_AND), g_blnRegHideAdultItems)
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Info", "LastMusicSQL", strQuery, True
            Debug.Print "(Music) " & strQuery
        End With
        Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
        If Not g_rsGames.EOF Then
            g_rsGames.MoveLast
            g_rsGames.MoveFirst
            lngTotal = lngTotal + g_rsGames.RecordCount
            If lngTotal > gc_lngACCEPTABLEWAIT And m_blnFormIsLoading = False Then
                If lblPleaseWait.Visible = False Then
                    'show the wait message
                    With picPleaseWait
                        .Top = 0
                        .Left = 0
                        .Width = lvMain(lngListViewIdx_in).Width
                        .Height = lvMain(lngListViewIdx_in).Height
                        lblPleaseWait.Top = (.Height / 2) - (lblPleaseWait.Height / 2)
                        lblPleaseWait.Left = 0
                        lblPleaseWait.Width = .Width
                        .Visible = True
                    End With
                    If g_blnSettingCatFilter = False Then
                        DoEvents
                    End If
                End If
            End If
            'setup field objects for quick access
            Set fldName = g_rsGames!Name
            Set fldID = g_rsGames!GA_Id
            Set fldBold1 = g_rsGames!Bold1
            If ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount > 0 Then
                ReDim Preserve fldCol(ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1)
                On Error Resume Next
                For lngExtraColumnIdx = 0 To ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount - 1
                    Set fldCol(lngExtraColumnIdx) = g_rsGames.Fields("Col" & lngExtraColumnIdx + 2)
                Next lngExtraColumnIdx
                On Error GoTo 0
            End If

            If ga_udtViewData(lngViewIdx_in).lngExtraColumnsCount = 0 Then
                'Only One Column
                With lvMain(lngListViewIdx_in).ListItems
                    Do Until g_rsGames.EOF
                        Set LItem = .Add(, "M" & fldID, fldName, 5, 5)
                        LItem.Bold = fldBold1
                        Set LItem = Nothing
                        g_rsGames.MoveNext
                    Loop
                End With
            Else
                udtView = ga_udtViewData(lngViewIdx_in)
                Do Until g_rsGames.EOF
                    Set LItem = lvMain(lngListViewIdx_in).ListItems.Add(, "M" & fldID, fldName, 5, 5)
                    With LItem
                        .Bold = fldBold1
                        For lngExtraColumnIdx = 0 To udtView.lngExtraColumnsCount - 1
                            If Not fldCol(lngExtraColumnIdx) Is Nothing Then
                                .SubItems(lngExtraColumnIdx + 1) = fldCol(lngExtraColumnIdx)
                            Else
                                .SubItems(lngExtraColumnIdx + 1) = ""
                            End If
                        Next lngExtraColumnIdx
                    End With
                    Set LItem = Nothing
                    g_rsGames.MoveNext
                Loop
            End If
            'release quick-access field objects
            Set fldName = g_rsGames!Name
            Set fldID = g_rsGames!GA_Id
            Set fldBold1 = g_rsGames!Bold1
            For lngExtraColumnIdx = 0 To udtView.lngExtraColumnsCount - 1
                Set fldCol(lngExtraColumnIdx) = Nothing
            Next lngExtraColumnIdx
        End If
        g_rsGames.Close
        Set g_rsGames = Nothing
    End If

    'hide the wait form if it's showing
    If picPleaseWait.Visible = True Then
        picPleaseWait.Visible = False
    End If

    'show and enable the listview again
    With lvMain(lngListViewIdx_in)
        .Visible = True
        .Enabled = True
    End With

    'set the view to being loaded if not a custom view
    If lngListViewIdx_in <= gc_lngDEFAULTVIEWSMAX Then
        ga_blnViewsLoaded(lngListViewIdx_in) = True
    End If
    
    'update the item count (incase wrong) in lstView and any controls that need enabling/disabling
    'also show the number of games/items in the columnheader
    With lstView
        strCount = .List(.ListIndex)
        If lngTotal > 0 Then
            strCount = Left$(strCount, InStrRev(strCount, "(")) & lngTotal & ")"
            lvMain(lngListViewIdx_in).ColumnHeaders(1).Text = lvMain(lngListViewIdx_in).ColumnHeaders(1).Text & "  (" & lngTotal & ")"
            ReadViewIntoList = True
        Else
            strCount = Left$(strCount, InStrRev(strCount, "(")) & g_strViewNone & ")"
            lvMain(lngListViewIdx_in).ColumnHeaders(1).Text = lvMain(lngListViewIdx_in).ColumnHeaders(1).Text & "  (" & g_strViewNone & ")"
            ReadViewIntoList = False
        End If
        .List(.ListIndex) = strCount
    End With
    
    'debug code
    Debug.Print "ListView took " & g_xobjClock.TimeTaken & " Secs to fill"
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Info", "LastViewTimeToFill", Trim$(g_xobjClock.TimeTaken) & " seconds", True

    Exit Function

End Function
Private Sub ShowTheBigPicture(ByVal lngPicIdx_in As Long, strCaption_in As String)

    Dim strFilename As String
    Dim strExt As String
    Dim lngPathIdx As Long
    Dim blnFound As Boolean
    Dim lngExtraID As Long
    Dim strPath As String
    Dim lngExtrasPathIdx As Long

    On Error GoTo Err_Handler

    blnFound = False

    If picSSMain(lngPicIdx_in).Tag > -1 Then

        'screenshot

        strFilename = g_udtGame.strScreenshotFilename
        strExt = Right$(g_udtGame.strScreenshotFilename, Len(g_udtGame.strScreenshotFilename) - InStrRev(g_udtGame.strScreenshotFilename, "."))
    
        If Not g_blnMamePicStyle Then
            'get correct filename
            If lngPicIdx_in > 0 Then
                strFilename = Left$(strFilename, InStrRev(strFilename, ".") - 1)
                strFilename = strFilename & "_" & Trim$(lngPicIdx_in) & "." & strExt
            End If
        
            'get full path\filename
            blnFound = False
            For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(lngPathIdx).strWithBackSlash & strFilename, False, "") Then
                    strFilename = ga_udtRegPathScrnShot(lngPathIdx).strWithBackSlash & strFilename
                    blnFound = True
                    Exit For
                End If
            Next lngPathIdx
        Else
            'get the correct path and filename
            blnFound = False
            If picSSMain(lngPicIdx_in).Tag < g_lngTotalScrnShotPaths Then
                If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(picSSMain(lngPicIdx_in).Tag).strWithBackSlash & strFilename, False, "") Then
                    strFilename = ga_udtRegPathScrnShot(picSSMain(lngPicIdx_in).Tag).strWithBackSlash & strFilename
                    blnFound = True
                End If
            End If
        End If
    Else
        'image extra
        
        blnFound = False
        strPath = ""
        lngExtraID = Abs(picSSMain(lngPicIdx_in).Tag)
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Name, Path FROM Extras WHERE EA = True AND EX_Id = " & lngExtraID, dbOpenSnapshot)
        With g_rsGames
            If Not .EOF Then
                .MoveFirst
                strPath = !Path
                strCaption_in = strCaption_in & " [" & !Name & "]"
                .Close
            End If
        End With
        Set g_rsGames = Nothing

        If strPath <> "" Then
            If FindExtrasFile(lngExtrasPathIdx, strPath) = "" Then
                strFilename = ga_udtRegPathExtras(lngExtrasPathIdx).strWithBackSlash & strPath
                blnFound = True
            End If
        End If
        
    End If

    If blnFound = True Then
        'load the screenshot into picfull then call the zoom form
        Select Case UCase$(strExt)
            Case "PNG"
                picFull.Cls
                picFull.Picture = LoadPicture
                picFull.AutoSize = False
                png_load strFilename
                picFull.Width = png_width * Screen.TwipsPerPixelX
                picFull.Height = png_height * Screen.TwipsPerPixelY
                png_draw picFull.hDC, 0, 0
                picFull.Picture = picFull.Image
            Case Else
                picFull.AutoSize = True
                picFull.Picture = LoadPicture(strFilename)
        End Select
        If g_blnSetPicZoomForeground = True Then
            frmBigScreenshot.Show vbModeless, Me
            frmBigScreenshot.ShowPicture picFull, strCaption_in
        Else
            frmBigScreenshot.ShowPicture picFull, strCaption_in
        End If
    Else
        Call MsgBox(g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInPicturePaths", "The file was not found in any of your Picture Paths!") & vbCrLf & vbCrLf & "'" & strFilename & "'", vbExclamation, gc_strAPPNAME)
    End If

Err_Handler:

End Sub

Private Sub UnCacheViews()

    ga_blnViewsLoaded(0) = False
    ga_blnViewsLoaded(1) = False
    ga_blnViewsLoaded(2) = False
    ga_blnViewsLoaded(3) = False

End Sub
Public Function GetListViewIdxFromViewIdx(ByVal lngListViewIdx_in As Long) As Long
    
    '----------------------------------------------------------
    'Gets the correct listview index for lvMain() based on the
    'passed in actual custom view index
    '
    'Inputs:        lngListViewIdx_in   The custom view index
    '
    'Returns:       the corresponding lvMain() listview index
    '----------------------------------------------------------
    
    If lngListViewIdx_in > gc_lngDEFAULTVIEWSMAX Then
        GetListViewIdxFromViewIdx = gc_lngDEFAULTVIEWSMAX + 1
    Else
        GetListViewIdxFromViewIdx = lngListViewIdx_in
    End If

End Function
Private Sub SetWindowPosition()

    If m_sngRegTop < 0 Or m_sngRegTop > Screen.Height Or m_sngRegLeft < 0 Or m_sngRegLeft > Screen.Width Then
        m_sngRegTop = (Screen.Height / 2) - (Height / 2)
        m_sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    If m_blnRegMaximized = False Then
        Top = m_sngRegTop
        Left = m_sngRegLeft
    Else
        WindowState = vbMaximized
    End If

End Sub


Private Sub SetInitialUIElements()

    '----------------------------------------
    'sets the initial user interface elements
    '----------------------------------------

    Dim lngIdx As Long              'loop index
    Dim lngTab As Long              'current tab index in registry
    Dim lngBackCol As Long          'listview background colour
    Dim lngForeCol As Long          'listview foreground colour
    Dim blnFontBold As Boolean      'listview font bold property
    Dim blnFontItalic As Boolean    'listview font italic property
    Dim strFontName As String       'listview font name property
    Dim lngFontSize As Long         'listview font size property
    Dim a_strGBN() As String        'for storing the gamebase (custom) names
    Dim lngTotalGBs As Long         'total num of gbs in array
    Dim strDir As String            'for getting the GB names
    Dim lngAttr As Long

    'set the visible game details tab
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "GameTab", lngTab, False, 0
    If lngTab = tabDetails.Tab Then
        m_blnDoneTabDetails = True
    End If
    tabDetails.Tab = lngTab

    'pictures popup menu text
    pmnuScrnShot(0).Caption = g_xobjFunc.GetINIMessage("Pictures Popup Menu", "Zoom", "Zoom In")
    pmnuScrnShot(1).Caption = g_xobjFunc.GetINIMessage("Pictures Popup Menu", "PictureSelector", "Picture Selector") & "..."

    'Database menu
    lngTotalGBs = 0
    ReDim ma_strGameBaseNames(0)
    ReDim a_strGBN(0)
    strDir = Dir(g_strAppPathWithBackslash, vbDirectory)
    Do Until strDir = ""
        lngAttr = GetAttr(g_strAppPathWithBackslash & strDir)
        If lngAttr And vbDirectory Then
            If (strDir <> ".") And (strDir <> "..") And (UCase$(strDir) <> "DOCS") Then
                If g_xobjFunc.FileExists(g_strAppPathWithBackslash & strDir & "\" & strDir & ".mdb", False, "") = True Then
                    ReDim Preserve ma_strGameBaseNames(lngTotalGBs)
                    ReDim Preserve a_strGBN(lngTotalGBs)
                    ma_strGameBaseNames(lngTotalGBs) = strDir
                    a_strGBN(lngTotalGBs) = GetCustomDatabaseName(strDir)
                    lngTotalGBs = lngTotalGBs + 1
                End If
            End If
        End If
        strDir = Dir$()
    Loop
    If lngTotalGBs < 2 Then
        mnuGBPick.Enabled = False
        mnuGBPick.Visible = False
    Else
        For lngIdx = 1 To (lngTotalGBs - 1)
            Load mnuGB(lngIdx)
        Next lngIdx
        For lngIdx = 0 To (lngTotalGBs - 1)
            mnuGB(lngIdx).Caption = a_strGBN(lngIdx)
            If UCase$(ma_strGameBaseNames(lngIdx)) = UCase$(g_strGameBaseName) Then
                mnuGB(lngIdx).Enabled = False
            End If
        Next lngIdx
    End If

    'please wait message
    lblPleaseWait = g_strPleaseWait

    'bold comments
    lbl_Info_Comment.FontBold = (g_xobjFunc.INI_GetSetting("UI", "Tab1BoldComment", "0", g_strSettingsFile) = "1")
    lbl_Info_VComment.FontBold = (g_xobjFunc.INI_GetSetting("UI", "Tab2BoldComment", "0", g_strSettingsFile) = "1")

    'gfx on the details tabs
    imgCup.Top = m_sngTabIconsCupTop
    imgCup2.Top = m_sngTabIconsCupTop
    imgPrequel.Top = m_sngTabIconsTop
    imgRelated.Top = m_sngTabIconsTop
    imgCloneOf.Top = m_sngTabIconsTop
    imgSequel.Top = m_sngTabIconsTop
    imgRatingFace.Top = m_sngTabIconsTop
    imgAdult.Top = m_sngTabIconsTop
    imgRRStore(51).Top = m_sngTabIconsTop

    'menu items
    mnuToolsHideAdult.Checked = g_blnRegHideAdultItems
    mnuToolsSounds.Checked = m_blnRegSounds
    sbStatus.Panels(5).Text = IIf((m_blnRegSounds), "S", "")
    mnuToolsNSM.Checked = m_blnRegNonStopMusic
    sbStatus.Panels(6).Text = IIf((m_blnRegNonStopMusic), "N", "")
    mnuToolsVerify.Enabled = AllowVerifyFiles()

    'popup views menu
    pmnuView(0).Caption = mnuViewAddView.Caption
    pmnuView(1).Caption = mnuViewEditView.Caption
    pmnuView(2).Caption = mnuViewRemoveView.Caption
    pmnuView(4).Caption = mnuViewQuickFilter.Caption
    pmnuView(6).Caption = g_xobjFunc.GetINIMessage("MainMenus", "Menu_View_SortCustomViews", "&Sort Custom Views") & "..."
    
    'picture backcolour
    picSSMain(0).BackColor = g_lngPictureBackColor
    
    'status bar
    sbStatus.Panels(7).ToolTipText = gc_strAPPNAME & " ID"

    'replace menu checkmarks with radio buttons
    'menus and menuitems are zero-indexed, splitter bars are included
    Call SetRadioMenuChecks(Me, 2, 2)   'view menu - sidebar hidden
    Call SetRadioMenuChecks(Me, 2, 3)   'view menu - sidebar on left
    Call SetRadioMenuChecks(Me, 2, 4)   'view menu - sidebar on right
    Call SetRadioMenuChecks(Me, 2, 6)   'view menu - hide screenshots
    Call SetRadioMenuChecks(Me, 2, 7)   'view menu - small screenshots
    Call SetRadioMenuChecks(Me, 2, 8)   'view menu - full size screenshots
    Call SetRadioMenuChecks(Me, 2, 14)  'view menu - list
    Call SetRadioMenuChecks(Me, 2, 15)  'view menu - details

    'listview colours and font
    With g_xobjRegistry
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "BackColour", lngBackCol, False, lstView.BackColor
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "ForeColour", lngForeCol, False, lstView.ForeColor
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontBold", blnFontBold, False, lstView.Font.Bold
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontItalic", blnFontBold, False, lstView.Font.Italic
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontName", strFontName, True, lstView.Font.Name
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontSize", lngFontSize, False, lstView.Font.Size
    End With
    For lngIdx = 0 To lvMain.Count - 1
        With lvMain(lngIdx)
            .BackColor = lngBackCol
            .ForeColor = lngForeCol
            .Font.Bold = blnFontBold
            .Font.Italic = blnFontItalic
            .Font.Name = strFontName
            .Font.Size = lngFontSize
        End With
    Next lngIdx
    With lvLinks
        .BackColor = lngBackCol
        .ForeColor = lngForeCol
        .Font.Bold = blnFontBold
        .Font.Italic = blnFontItalic
        .Font.Name = strFontName
        .Font.Size = lngFontSize
    End With

End Sub
Private Function GetCustomDatabaseName(ByVal strDBName_in As String) As String

    On Error GoTo Err_Handler

    Dim dbGamebase As Database
    Dim rsGames As Recordset

    GetCustomDatabaseName = strDBName_in

    Set dbGamebase = OpenDatabase(g_strAppPathWithBackslash & strDBName_in & "\" & strDBName_in & ".mdb")

    dbGamebase.Execute ("UPDATE Config SET DatabaseName="""" WHERE DatabaseName IS NULL")
    
    Set rsGames = dbGamebase.OpenRecordset("SELECT DatabaseName FROM Config", dbOpenSnapshot)
    With rsGames
        If Not .EOF Then
            .MoveFirst
            If Trim$(!DatabaseName) <> "" Then
                GetCustomDatabaseName = Trim$(!DatabaseName)
            End If
            .Close
        End If
    End With
    Set rsGames = Nothing
    
    dbGamebase.Close
    Set dbGamebase = Nothing

    Exit Function

Err_Handler:

End Function
Private Sub FillViewSelector()
    
    '-------------------------------------------------------------
    'Reads in the global array of Views into the View Selector
    '(lstView) and calculates the total items for each view.
    '
    'Remarks:       The a_strCountSQL Array is based upon what the
    '               Default Views are set to in LoadDefaultViews()
    '-------------------------------------------------------------
    
    Dim a_strCountSQL(gc_lngDEFAULTVIEWSMAX) As String  'quick SQL count queries for the default views
    Dim strCountSQLFavSid As String                     'extra SQL count query for the favourite view
    Dim a_strItemCount()                                'array to hold each view's item count
    Dim lngViewIdx As Long                              'loop index
    Dim lngTempItemCount As Long                        'temporary item counter
    Dim lngViewMax As Long                              'UBound() of the global views array
    Dim rsFav As Recordset                              'extra recordset for the Favourites view
    Dim strCustomViewSQL As String                      'SQL query for each Custom View
    Dim a_strFilterBuffers() As String                  'array for sending the filters to GBSql.clsQuery
    Dim lngCalcTotals As Long                           'calculate view totals?
    
    'start timer
    g_xobjClock.StartClock

    '''''''''''''''''''''
    'Initialise Variables
    '''''''''''''''''''''

    'are we getting the totals on startup?
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "CalcTotals", lngCalcTotals, False, 0

    'set the CountSQL strings to quickly
    'total the number of items in each view
    a_strCountSQL(0) = "SELECT Fav FROM Games"
    a_strCountSQL(1) = "SELECT Fav FROM Games WHERE SidFilename <> """" UNION ALL SELECT SFav FROM Music WHERE Filename <> """""
    a_strCountSQL(2) = "SELECT Fav FROM Games WHERE Fav = True"
    a_strCountSQL(3) = "SELECT Fav FROM Games WHERE CloneOf = 0"
    strCountSQLFavSid = "SELECT SFav FROM Music WHERE SFav = True"

    'initialise the item count array
    lngViewMax = UBound(ga_udtViewData)
    ReDim a_strItemCount(lngViewMax)

    'get the totals if option says we can
    If lngCalcTotals = 1 Then
        
        '''''''''''''''''''''''''''''''''
        'get the default view item counts
        '''''''''''''''''''''''''''''''''

        For lngViewIdx = 0 To gc_lngDEFAULTVIEWSMAX
            Set g_rsGames = g_dbGamebase.OpenRecordset(a_strCountSQL(lngViewIdx), dbOpenSnapshot)
            With g_rsGames
                If lngViewIdx = gc_lngFAVOURITESINDEX Then

                    'Favourite Games and Sids
                    '------------------------
                    'First the game favourites
                    If .EOF Then
                        lngTempItemCount = 0
                    Else
                        .MoveLast
                        .MoveFirst
                        lngTempItemCount = .RecordCount
                    End If

                    'now the sid favourites
                    Set rsFav = g_dbGamebase.OpenRecordset(strCountSQLFavSid, dbOpenSnapshot)
                    With rsFav
                        If .EOF Then
                            a_strItemCount(lngViewIdx) = lngTempItemCount
                        Else
                            .MoveLast
                            .MoveFirst
                            a_strItemCount(lngViewIdx) = lngTempItemCount + .RecordCount
                        End If
                        .Close
                    End With
                    Set rsFav = Nothing
                Else
                    'Gamebase Collection
                    If .EOF Then
                        a_strItemCount(lngViewIdx) = 0
                    Else
                        .MoveLast
                        .MoveFirst
                        a_strItemCount(lngViewIdx) = .RecordCount
                    End If
                End If
                .Close
            End With
            Set g_rsGames = Nothing
        Next lngViewIdx

        ''''''''''''''''''''''''''''''''
        'get the custom view item counts
        ''''''''''''''''''''''''''''''''
        
        For lngViewIdx = (gc_lngDEFAULTVIEWSMAX + 1) To lngViewMax

            'GAMES TABLE
            If ga_udtViewData(lngViewIdx).eIncludeMusicTable <> e_Table_SidOnly Then
                'get the view SQL string
                With ga_udtViewData(lngViewIdx)
                    Call MakeGB64SQLFilterString(.a_udtFilters(), .lngFilterCount, a_strFilterBuffers())
                    strCustomViewSQL = m_xobjSQLEngine.GetViewQueryString(True, "", .lngFilterCount, a_strFilterBuffers(), (.eFilterMode = e_FilterMode_AND), g_blnRegHideAdultItems)
                End With
                'read in the items
                Set g_rsGames = g_dbGamebase.OpenRecordset(strCustomViewSQL, dbOpenSnapshot)
                With g_rsGames
                    If Not .EOF Then
                        'games found in query
                        .MoveLast
                        .MoveFirst
                        lngTempItemCount = .RecordCount
                    Else
                        'no games returned in query
                        lngTempItemCount = 0
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
            Else
                'no query run on games table
                lngTempItemCount = 0
            End If

            'SID TABLE
            If ga_udtViewData(lngViewIdx).eIncludeMusicTable <> e_Table_GamesOnly Then
                'get the view SQL string
                With ga_udtViewData(lngViewIdx)
                    Call MakeGB64SQLFilterString(.a_udtFilters(), .lngFilterCount, a_strFilterBuffers())
                    strCustomViewSQL = m_xobjSQLEngine.GetViewQueryString(False, "", .lngFilterCount, a_strFilterBuffers(), (.eFilterMode = e_FilterMode_AND), g_blnRegHideAdultItems)
                End With
                'read in the items
                Set g_rsGames = g_dbGamebase.OpenRecordset(strCustomViewSQL, dbOpenSnapshot)
                With g_rsGames
                    If Not g_rsGames.EOF Then
                        g_rsGames.MoveLast
                        g_rsGames.MoveFirst
                        lngTempItemCount = lngTempItemCount + g_rsGames.RecordCount
                    End If
                    g_rsGames.Close
                End With
                Set g_rsGames = Nothing
            End If
            
            'store the total items
            a_strItemCount(lngViewIdx) = lngTempItemCount
            
        Next lngViewIdx

        'finally populate the View Selector (lstView)
        With lstView
            .Clear
            For lngViewIdx = 0 To lngViewMax
                If a_strItemCount(lngViewIdx) > 0 Then
                    .List(lngViewIdx) = ga_udtViewData(lngViewIdx).strTitle & "  (" & a_strItemCount(lngViewIdx) & ")"
                Else
                    .List(lngViewIdx) = ga_udtViewData(lngViewIdx).strTitle & "  (" & g_strViewNone & ")"
                End If
            Next lngViewIdx
        End With
        'debug code
        Debug.Print "View Selector took " & g_xobjClock.TimeTaken & " Secs to fill"
        
    Else
        'just populate the listview without the view totals
        With lstView
            .Clear
            For lngViewIdx = 0 To lngViewMax
                If a_strItemCount(lngViewIdx) > 0 Then
                    .List(lngViewIdx) = ga_udtViewData(lngViewIdx).strTitle & "  (?)"
                Else
                    .List(lngViewIdx) = ga_udtViewData(lngViewIdx).strTitle & "  (?)"
                End If
            Next lngViewIdx
        End With
    End If

End Sub
Private Sub ResizeWindowWithDetails()

    Dim NewScaleWidth As Single
    Dim NewScaleHeight As Single
    Dim HalfNewScaleWidth As Single
    Dim sngTemp As Single
    Dim SS As Long
    Dim sngListViewWidth As Single
    Dim sngListViewLeft As Single
    
    If WindowState <> vbMinimized Then
        If WindowState = vbMaximized Then
            mnuViewSSSnap.Enabled = False
        Else
            mnuViewSSSnap.Enabled = (m_lngRegScrnShots > 0)
        End If
        NewScaleWidth = ScaleWidth
        NewScaleHeight = ScaleHeight
        HalfNewScaleWidth = NewScaleWidth / 2
        
        '==============
        'width controls
        '==============
        
        'screenshots and listviews
        m_blnLastScrnShotWasNone = False
        m_blnLastPhotoWasNone = False
        If m_lngRegScrnShots = 0 Then
            sngTemp = NewScaleWidth - m_sngScrnShotWidth
            picSSMain(0).Left = sngTemp
        Else
            sngTemp = NewScaleWidth - m_sngScrnShotWidth - mc_TwipsBetweenLVMAINAndpicSSMain
            picSS.Left = sngTemp + mc_TwipsBetweenLVMAINAndpicSSMain
            picSSMain(0).Left = 0
        End If
        m_sngLastListViewSideBarAvailableWidth = sngTemp
        For SS = 1 To m_lngScrnShotCount
            picSSMain(SS).Left = 0
        Next SS
        
        If m_lngRegSidebar = 0 Then
            'sidebar hidden
            With picListViews
                .Left = 0
                .Width = sngTemp
                lvMain(0).Width = sngTemp
                lvMain(0).Width = sngTemp
                lvMain(1).Width = sngTemp
                lvMain(2).Width = sngTemp
                lvMain(3).Width = sngTemp
                lvMain(4).Width = sngTemp
            End With
        Else
            'sidebar visible
            If m_sngSidebarWidth + m_sngSplitterWidth > sngTemp Then
                'sidebar is now wider than the area to display both it and the listviews
                'so shrink it a bit so both fit
                m_sngSidebarWidth = sngTemp - m_sngSplitterWidth - mc_sngMINSIDEBARWIDTH
            End If
            If m_sngSidebarWidth < mc_sngMINSIDEBARWIDTH Then m_sngSidebarWidth = mc_sngMINSIDEBARWIDTH
            sngListViewWidth = sngTemp - m_sngSidebarWidth - m_sngSplitterWidth
            If m_lngRegSidebar = 1 Then
                'on left
                picSidebar.Left = 0
                picSidebar.Width = m_sngSidebarWidth
                picSplitter.Left = m_sngSidebarWidth
                sngListViewLeft = m_sngSidebarWidth + m_sngSplitterWidth
                With picListViews
                    .Width = sngListViewWidth
                    .Left = sngListViewLeft
                End With
            Else
                'on right
                picSplitter.Left = sngListViewWidth
                picSidebar.Left = sngListViewWidth + m_sngSplitterWidth
                picSidebar.Width = m_sngSidebarWidth
                With picListViews
                    .Left = 0
                    .Width = sngListViewWidth
                End With
            End If
            lvMain(0).Width = sngListViewWidth
            lvMain(1).Width = sngListViewWidth
            lvMain(2).Width = sngListViewWidth
            lvMain(3).Width = sngListViewWidth
            lvMain(4).Width = sngListViewWidth
            lvMain(0).Left = 0
            lvMain(1).Left = 0
            lvMain(2).Left = 0
            lvMain(3).Left = 0
            lvMain(4).Left = 0
        End If
        
        'frapanel / lstview / frabuttons
        fraPanel.Width = NewScaleWidth
        fraButtons.Left = NewScaleWidth - fraButtons.Width
        lstView.Width = fraButtons.Left - mc_TwipsBetweenLVMAINAndpicSSMain - cmdCustViews.Width
        
        'icons
        Call RepositionDTIconsX
        
        'tabDetails and tabSid
        With tabDetails
            .Left = HalfNewScaleWidth - (.Width / 2)
            tabSid.Left = .Left
        End With
        
        'status bar
        With sbStatus
            If g_eWhatIsSelected = e_SelType_IsSid Then
                .Panels(1).Width = 1
                .Panels(2).Width = (ScaleWidth - m_sngStatusBarOffset) - 1
                .Panels(3).Width = 1
                .Panels(4).Width = 1
            Else
                .Panels(1).Width = (ScaleWidth - m_sngStatusBarOffset - 800) / 3
                .Panels(2).Width = .Panels(1).Width
                .Panels(3).Width = 800
                .Panels(4).Width = .Panels(1).Width
            End If
        End With
        
        '===============
        'height controls
        '===============

        'sbstatus
        sbStatus.Top = NewScaleHeight - sbStatus.Height

        'tabDetails and tabSid, and lblName
        With tabDetails
            .Top = sbStatus.Top - mc_TwipsBetweenTABDETAILSAndSBSTATUS - .Height
            tabSid.Top = .Top
            lblName.Top = .Top - 345
            'gfx on the details tabs
            m_sngTabIconsCupTop = lblName.Top + ((lblName.Height / 2) - ((Screen.TwipsPerPixelY * 11) / 2)) + (Screen.TwipsPerPixelY / 2)
            imgCup.Top = m_sngTabIconsCupTop
            imgCup2.Top = m_sngTabIconsCupTop
            imgPrequel.Top = m_sngTabIconsCupTop
            imgRelated.Top = m_sngTabIconsCupTop
            imgCloneOf.Top = m_sngTabIconsCupTop
            imgSequel.Top = m_sngTabIconsCupTop
            imgRatingFace.Top = m_sngTabIconsCupTop
            imgAdult.Top = m_sngTabIconsCupTop
            imgRRStore(51).Top = m_sngTabIconsCupTop
        End With

        'frapanel / lstview / frabuttons
        fraPanel.Top = tabDetails.Top - mc_TwipsBetweenFRAPANELAndTABDETAILS - fraPanel.Height

        'lvmain and sidebar
        sngTemp = fraPanel.Top - mc_TwipsBetweenLVMAINAndFRAPANEL
        With picListViews
            .Height = sngTemp
            lvMain(0).Height = sngTemp
            lvMain(1).Height = sngTemp
            lvMain(2).Height = sngTemp
            lvMain(3).Height = sngTemp
            lvMain(4).Height = sngTemp
            If m_lngRegSidebar > 0 Then
                picSplitter.Height = sngTemp
                picSidebar.Height = sngTemp
            End If
        End With

        'calculate how many vertical screenshots can be seen
        If m_lngRegScrnShots = 0 Then
            m_lngScrnShotMaxVisible = 1
            picSS.Height = picSSMain(0).Height
        Else
            m_lngScrnShotMaxVisible = Int(lvMain(0).Height / m_sngScrnShotHeight) - 1
            If m_lngScrnShotMaxVisible < 0 Then m_lngScrnShotMaxVisible = 0
            picSS.Height = (m_lngScrnShotMaxVisible + 1) * m_sngScrnShotHeight
        End If
        If g_eWhatIsSelected = e_SelType_IsGame Then
            If Not m_ShowingPhoto Then
                ShowScreenshots
            Else
                imgPhoto.Visible = False
                imgPhoto.Stretch = False
                ShowPhoto
            End If
        ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
            imgPhoto.Visible = False
            imgPhoto.Stretch = False
            ShowPhoto
        End If
        
    End If
    
End Sub

Private Sub ResizeWindowWithoutDetails()

    Dim NewScaleWidth As Single
    Dim NewScaleHeight As Single
    Dim HalfNewScaleWidth As Single
    Dim sngTemp As Single
    Dim SS As Long
    Dim sngListViewWidth As Single
    Dim sngListViewLeft As Single
    
    If WindowState <> vbMinimized Then
        If WindowState = vbMaximized Then
            mnuViewSSSnap.Enabled = False
        Else
            mnuViewSSSnap.Enabled = (m_lngRegScrnShots > 0)
        End If
        NewScaleWidth = ScaleWidth
        NewScaleHeight = ScaleHeight
        HalfNewScaleWidth = NewScaleWidth / 2

        '==============
        'width controls
        '==============
        
        'screenshots and listviews
        m_blnLastScrnShotWasNone = False
        m_blnLastPhotoWasNone = False
        If m_lngRegScrnShots = 0 Then
            sngTemp = NewScaleWidth - m_sngScrnShotWidth
            picSSMain(0).Left = sngTemp
        Else
            sngTemp = NewScaleWidth - m_sngScrnShotWidth - mc_TwipsBetweenLVMAINAndpicSSMain
            picSS.Left = sngTemp + mc_TwipsBetweenLVMAINAndpicSSMain
            picSSMain(0).Left = 0
        End If
        m_sngLastListViewSideBarAvailableWidth = sngTemp
        For SS = 1 To m_lngScrnShotCount
            picSSMain(SS).Left = 0
        Next SS

        If m_lngRegSidebar = 0 Then
            'sidebar hidden
            With picListViews
                .Left = 0
                .Width = sngTemp
                lvMain(0).Width = sngTemp
                lvMain(0).Width = sngTemp
                lvMain(1).Width = sngTemp
                lvMain(2).Width = sngTemp
                lvMain(3).Width = sngTemp
                lvMain(4).Width = sngTemp
            End With
        Else
            'sidebar visible
            If m_sngSidebarWidth + m_sngSplitterWidth > sngTemp Then
                'sidebar is now wider than the area to display both it and the listviews
                'so shrink it a bit so both fit
                m_sngSidebarWidth = sngTemp - m_sngSplitterWidth - mc_sngMINSIDEBARWIDTH
            End If
            If m_sngSidebarWidth < mc_sngMINSIDEBARWIDTH Then m_sngSidebarWidth = mc_sngMINSIDEBARWIDTH
            sngListViewWidth = sngTemp - m_sngSidebarWidth - m_sngSplitterWidth
            If m_lngRegSidebar = 1 Then
                'on left
                picSidebar.Left = 0
                picSidebar.Width = m_sngSidebarWidth
                picSplitter.Left = m_sngSidebarWidth
                sngListViewLeft = m_sngSidebarWidth + m_sngSplitterWidth
                With picListViews
                    .Width = sngListViewWidth
                    .Left = sngListViewLeft
                End With
            Else
                'on right
                picSplitter.Left = sngListViewWidth
                picSidebar.Left = sngListViewWidth + m_sngSplitterWidth
                picSidebar.Width = m_sngSidebarWidth
                With picListViews
                    .Left = 0
                    .Width = sngListViewWidth
                End With
            End If
            lvMain(0).Width = sngListViewWidth
            lvMain(1).Width = sngListViewWidth
            lvMain(2).Width = sngListViewWidth
            lvMain(3).Width = sngListViewWidth
            lvMain(4).Width = sngListViewWidth
            lvMain(0).Left = 0
            lvMain(1).Left = 0
            lvMain(2).Left = 0
            lvMain(3).Left = 0
            lvMain(4).Left = 0
        End If
        
        'frapanel / lstview / frabuttons
        fraPanel.Width = NewScaleWidth
        fraButtons.Left = NewScaleWidth - fraButtons.Width
        lstView.Width = fraButtons.Left - mc_TwipsBetweenLVMAINAndpicSSMain - cmdCustViews.Width
        
        'icons
        Call RepositionDTIconsX
        
        'status bar
        With sbStatus
            If g_eWhatIsSelected = e_SelType_IsSid Then
                .Panels(1).Width = 1
                .Panels(2).Width = (ScaleWidth - m_sngStatusBarOffset) - 1
                .Panels(3).Width = 1
                .Panels(4).Width = 1
            Else
                .Panels(1).Width = (ScaleWidth - m_sngStatusBarOffset - 800) / 3
                .Panels(2).Width = .Panels(1).Width
                .Panels(3).Width = 800
                .Panels(4).Width = .Panels(1).Width
            End If
        End With
        
        '===============
        'height controls
        '===============
        
        'sbstatus
        sbStatus.Top = NewScaleHeight - sbStatus.Height

        'lblName and icons
        With lblName
            .Top = sbStatus.Top - 345
            'gfx on the details tabs
            m_sngTabIconsCupTop = .Top + ((.Height / 2) - ((Screen.TwipsPerPixelY * 11) / 2)) + (Screen.TwipsPerPixelY / 2)
            imgCup.Top = m_sngTabIconsCupTop
            imgCup2.Top = m_sngTabIconsCupTop
            imgPrequel.Top = m_sngTabIconsCupTop
            imgRelated.Top = m_sngTabIconsCupTop
            imgCloneOf.Top = m_sngTabIconsCupTop
            imgSequel.Top = m_sngTabIconsCupTop
            imgRatingFace.Top = m_sngTabIconsCupTop
            imgAdult.Top = m_sngTabIconsCupTop
            imgRRStore(51).Top = m_sngTabIconsCupTop
        End With

        'frapanel / lstview / frabuttons
        fraPanel.Top = sbStatus.Top - mc_TwipsBetweenFRAPANELAndSBSTATUS - fraPanel.Height

        'lvmain and sidebar
        sngTemp = fraPanel.Top - mc_TwipsBetweenLVMAINAndFRAPANEL
        With picListViews
            .Height = sngTemp
            lvMain(0).Height = sngTemp
            lvMain(1).Height = sngTemp
            lvMain(2).Height = sngTemp
            lvMain(3).Height = sngTemp
            lvMain(4).Height = sngTemp
            If m_lngRegSidebar > 0 Then
                picSplitter.Height = sngTemp
                picSidebar.Height = sngTemp
            End If
        End With
        
        'calculate how many vertical screenshots can be seen
        If m_lngRegScrnShots = 0 Then
            m_lngScrnShotMaxVisible = 1
            picSS.Height = picSSMain(0).Height
        Else
            m_lngScrnShotMaxVisible = Int(lvMain(0).Height / m_sngScrnShotHeight) - 1
            If m_lngScrnShotMaxVisible < 0 Then m_lngScrnShotMaxVisible = 0
            picSS.Height = (m_lngScrnShotMaxVisible + 1) * m_sngScrnShotHeight
        End If
        If g_eWhatIsSelected = e_SelType_IsGame Then
            If Not m_ShowingPhoto Then
                ShowScreenshots
            Else
                imgPhoto.Visible = False
                imgPhoto.Stretch = False
                ShowPhoto
            End If
        ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
            imgPhoto.Visible = False
            imgPhoto.Stretch = False
            ShowPhoto
        End If
        
    End If
    
End Sub

Private Sub SetWindowSize(ByVal lngSS_in As Long, ByVal blnDetails_in As Boolean, ByVal sngWidth_in As Single, ByVal sngHeight_in As Single, ByVal lngSidebar_in As Long)

    'un-subclass the window
    If g_blnInVBIDE = False Then
        If g_blnSubclassed = True Then
            Call UnSubClass(Me.hwnd)
        End If
    End If

    'get the new min size for the subclass
    With m_udtMinSizes
        
        'width
        If sngWidth_in < .sngWidth_Twips Then
            sngWidth_in = .sngWidth_Twips
        End If
        lngWindowMinX = .lngWidth_Pixels
        lngWindowMaxX = .lngScreenWidth_Pixels
        
        'height
        Select Case blnDetails_in
            
            Case False 'HIDE DETAILS
                
                tabDetails.Visible = False
                tabDetails.Enabled = False
                tabSid.Visible = False
                tabSid.Enabled = False
                
                Select Case lngSS_in
                
                    Case 0, 1 'hidden or small Screenshots
                        If sngHeight_in < .sngHeight_SmallSSNoDetails_Twips Then
                            sngHeight_in = .sngHeight_SmallSSNoDetails_Twips
                        End If
                        lngWindowMinY = .lngHeight_SmallSSNoDetails_Pixels
                        
                    Case Else 'big screenshots
                        If sngHeight_in < .sngHeight_FullSSNoDetails_Twips Then
                            sngHeight_in = .sngHeight_FullSSNoDetails_Twips
                        End If
                        lngWindowMinY = .lngHeight_FullSSNoDetails_Pixels
                    
                End Select
                
            Case True 'SHOW DETAILS
         
                Select Case lngSS_in
                
                    Case 0, 1 'hidden or small Screenshots
                        If sngHeight_in < .sngHeight_SmallSSDetails_Twips Then
                            sngHeight_in = .sngHeight_SmallSSDetails_Twips
                        End If
                        lngWindowMinY = .lngHeight_SmallSSDetails_Pixels
                        
                    Case Else 'big screenshots
                        If sngHeight_in < .sngHeight_FullSSDetails_Twips Then
                            sngHeight_in = .sngHeight_FullSSDetails_Twips
                        End If
                        lngWindowMinY = .lngHeight_FullSSDetails_Pixels
                    
                End Select
            
        End Select
        lngWindowMaxY = .lngScreenHeight_Pixels
        
        'store the width/height of the screenshots for resize events
        If lngSS_in = 0 Then
            m_sngScrnShotWidth = 0
            m_sngScrnShotHeight = 0
            picSSMain(0).Visible = False
            picSS.Visible = False
        ElseIf lngSS_in = 1 Then
            m_sngScrnShotWidth = .sngSmallSSWidth_Twips
            m_sngScrnShotHeight = .sngSmallSSHeight_Twips
            picSSMain(0).Visible = True
            picSSMain(0).Width = .sngSmallSSWidth_Twips
            picSSMain(0).Height = .sngSmallSSHeight_Twips
            picSS.Width = m_sngScrnShotWidth
            picSS.Visible = True
        Else
            m_sngScrnShotWidth = .sngFullSSWidth_Twips
            m_sngScrnShotHeight = .sngFullSSHeight_Twips
            picSSMain(0).Visible = True
            picSSMain(0).Width = .sngFullSSWidth_Twips
            picSSMain(0).Height = .sngFullSSHeight_Twips
            picSS.Width = m_sngScrnShotWidth
            picSS.Visible = True
        End If

        'sidebar
        If lngSidebar_in = 0 Then
            picSidebar.Visible = False
            picSplitter.Visible = False
        Else
            picSidebar.Top = 0
            picSplitter.Top = 0
            picSidebar.Visible = True
            picSplitter.Visible = True
        End If

    End With
    
    're-subclass the window
    If g_blnInVBIDE = False Then
        Call SubClass(Me.hwnd)
    End If
    
    'set the menu checkmarks
    mnuViewScreenshots(0).Checked = (lngSS_in = 0)
    mnuViewScreenshots(1).Checked = (lngSS_in = 1)
    mnuViewScreenshots(2).Checked = (lngSS_in = 2)
    mnuViewSidebar(0).Checked = (lngSidebar_in = 0)
    mnuViewSidebar(1).Checked = (lngSidebar_in = 1)
    mnuViewSidebar(2).Checked = (lngSidebar_in = 2)
    mnuViewInfo.Checked = blnDetails_in
    
    'force a control resize incase the window size didn't actually change
    m_blnResizeEnable = False
    If WindowState = vbNormal Then
        Move Left, Top, sngWidth_in, sngHeight_in
    End If
    m_blnResizeEnable = True
    ShowItem
    If blnDetails_in = True Then
        ResizeWindowWithDetails
    Else
        ResizeWindowWithoutDetails
    End If

End Sub

Public Sub ShowItem()

    '----------------------------------------------------------
    'Shows the selected item in the selected listview (lvMain)
    '
    'Remarks:   This procedure is called whenever a user clicks
    '           on a game or sid in the selected view
    '----------------------------------------------------------

    Dim strQuery As String              'the SQL Query to fetch for the selected item's info
    Dim itmSelectedItem As ListItem     'the selected listitem
    Dim lngID As Long                   'the selected listitem's database ID (GA_Id)
    Dim strName As String               'temp variable for formatting the game name
    Dim strScrnShot As String           'storage for screenshot filename
    Dim lngIdx As Long                  'loop index
    Dim blnSidFileExists As Boolean     'flag for if the selected item's sid-file exists
    Dim blnPSID As Boolean              'linked sid is a PSID
    Dim blnHasMemo As Boolean           'flag true if game has memotext (for icon change)
    Dim strVersion As String            'text for the game version number in the statusbar (panel 1)
    Dim strTest As String               'for shortening and ellipsing the musician label on the sid-only details tab
    
    '''''''''''''''''''''''''''''''''''''
    'PROCESS GENERAL STUFF PER ITEM-CLICK
    '''''''''''''''''''''''''''''''''''''
    
    'hide the Icons on the detail tabs (otherwise looks messy)
    'imgCup(lngID).Visible = False
    imgPrequel.Visible = False
    imgRelated.Visible = False
    imgSequel.Visible = False
    imgCloneOf.Visible = False
    imgRatingFace.Visible = False
    imgAdult.Visible = False
    'allow time to hide before continuing
    If g_blnSettingCatFilter = False Then
        DoEvents
    End If

    'grab the selected item and ID
    If lstView.ListIndex = gc_lngNOLISTINDEX Then Exit Sub
    Set itmSelectedItem = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If itmSelectedItem Is Nothing Then Exit Sub
    With itmSelectedItem
        lngID = Val(Right$(.Key, Len(.Key) - 1))
    End With

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'test if an item from the Games table or Music table was clicked
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If itmSelectedItem.Icon = 5 Then

        '''''''''''''''''''''''''''''''''''''''''''
        'ITEM CLICKED IS A SID FROM THE MUSIC TABLE
        '''''''''''''''''''''''''''''''''''''''''''

        'store global variable so rest of program knows what's selected
        g_eWhatIsSelected = e_SelType_IsSid
        m_ShowingPhoto = True
        
        'menus
        mnuEditGame.Enabled = False
        mnuEditScrnShots.Enabled = False
        mnuEditExtras.Enabled = False
        mnuEditGemus.Enabled = False
        mnuEditVersion.Enabled = False
        mnuEditPersonal.Enabled = False
        mnuEditGame.Enabled = True
        mnuEditLinkSet.Enabled = False
        mnuEditMark.Enabled = False
        mnuEditLinkClear.Enabled = False
        
        'buttons
        cmdPlay.Enabled = False
        cmdExtra.Enabled = False

        'construct the SELECT query
        '--------------------------
        'fields
        strQuery = "SELECT Music.GA_Id, Music.Name, Music.Filename, Music.SFav, Music.Adult, "
        strQuery = strQuery & "Musicians.Photo, Musicians.Grp, Musicians.Nick, Musicians.MU_Id, "
        strQuery = strQuery & "Musicians.Musician "
        'tables and inner joins
        strQuery = strQuery & "FROM Musicians INNER JOIN Music "
        strQuery = strQuery & "ON Musicians.MU_Id = Music.MU_Id "
        'where clause
        strQuery = strQuery & "WHERE Music.GA_Id = " & lngID

        'Run the query, get the Sid info from the database and store globally
        Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
        If g_rsGames.EOF Then
            MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_InvalidGameMusic", "The selected game/music has missing or invalid data in one or more of it's linked fields.") & vbCrLf & vbCrLf & "(Music)" & vbTab & "GA_Id = " & lngID & vbTab & "Name = """ & itmSelectedItem.Text & """", vbCritical, gc_strAPPNAME
            Call UnloadAllForms
            If g_blnInVBIDE = False Then
                If g_blnSubclassed = True Then
                    Call UnSubClass(Me.hwnd)
                End If
            End If
            End
        End If
        With g_rsGames
            .MoveFirst
            g_udtSid.lngID = lngID
            If g_blnOptionsShowGameID Then
                Caption = g_strWINDOWCAPTION & "    {" & lngID & "}"
            End If
            sbStatus.Panels(7).Text = lngID
            g_udtSid.strName = !Name
            g_udtSid.strFilename = !Filename
            If g_udtSid.strFilename = "" Then
                g_udtSid.strFilename = g_strNone
            End If
            g_udtSid.blnGotInfo = True
            g_udtSid.blnGameTable = False
            g_udtSid.lngMU_Id = !MU_Id
            g_udtSid.strMusician = !Musician
            g_udtSid.strNick = !Nick
            g_udtSid.strGroup = !Grp
            g_udtSid.strPhoto = !Photo
            g_udtSid.blnAdult = !Adult
            g_udtSid.blnFav = !SFav
            'amend field data
            If g_udtSid.strGroup = "" Then g_udtSid.strGroup = g_strNone
            If g_udtSid.strNick = "" Then g_udtSid.strNick = g_strNone
            If g_udtSid.strPhoto = "" Then g_udtSid.strPhoto = g_strNone
            .Close
        End With
        Set g_rsGames = Nothing
        
        'get the Sid file info and store whether the file exists (for the Play Sid button)
        With g_udtSid
            blnSidFileExists = LoadSidFileInfo(blnPSID)
            If blnSidFileExists = True Then
                .blnFileExists = True
                If blnPSID = True Then
                    lblSidNotFound.Visible = False
                    fraSidFile.Visible = True
                    lbl_SInfo_Name = .strSFName
                    lbl_SInfo_Author = .strSFAuthor
                    lbl_SInfo_Copyright = .strSFCopyright
                    lbl_SInfo_Songs = .lngSFTotalSongs
                    lbl_SInfo_Default = .lngSFDefaultSong
                Else
                    lblSidNotFound.Visible = True
                    fraSidFile.Visible = False
                End If
            Else
                .blnFileExists = False
                lblSidNotFound.Visible = True
                fraSidFile.Visible = False
            End If
        End With
        
        'show musician photo (if any)
        Call ClearScreenshots
        picSS.Visible = False
        picSSMain(0).Visible = False
        picSSUp.Visible = False
        picSSDown.Visible = False
        With imgPhoto
            .Stretch = False
            .Visible = False
            If g_udtSid.strPhoto = g_strNone Then
                'no photo
                If m_blnLastPhotoWasNone = False Then
                    .Picture = imgNoPhoto.Picture
                    m_blnLastPhotoWasNone = True
                End If
            ElseIf g_xobjFunc.FileExists(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto, False, gc_strAPPNAME) = False Then
                'photo referenced in DB but not found in photo path
                If m_blnLastPhotoWasNone = False Then
                    .Picture = imgNoPhoto.Picture
                    m_blnLastPhotoWasNone = True
                End If
            Else
                'photo found!
                .Picture = LoadPicture(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto)
                m_blnLastPhotoWasNone = False
            End If
            m_sngCachedPhotoHeight = .Height
            m_sngCachedPhotoWidth = .Width
            Call ShowPhoto
        End With
        
        'show the Sid Info tab and hide the game tabs
        tabDetails.Visible = False
        tabDetails.Enabled = False
        If m_blnRegDetails Then
            tabSid.Visible = True
            tabSid.Enabled = True
        End If
        
        'show the sid details
        With g_udtSid
            'status bar
            sbStatus.Panels(1) = ""
            sbStatus.Panels(1).Width = 1
            sbStatus.Panels(2).Width = (ScaleWidth - m_sngStatusBarOffset) - 1
            sbStatus.Panels(3).Width = 1
            sbStatus.Panels(4).Width = 1
            sbStatus.Panels(2) = EllipseText(.strFilename, sbStatus.Panels(2).Width)
            If sbStatus.Panels(2) <> .strFilename Then
                sbStatus.Panels(2).ToolTipText = .strFilename
            Else
                sbStatus.Panels(2).ToolTipText = ""
            End If
            'tab display
            strTest = EllipseTextEnd(.strMusician, 2925)
            lbl_SInfo_Musician = strTest
            If strTest <> .strMusician Then
                lbl_SInfo_Musician.ToolTipText = .strMusician
            Else
                lbl_SInfo_Musician.ToolTipText = ""
            End If
            lbl_SInfo_Nick = .strNick
            lbl_SInfo_Group = .strGroup
            lbl_SInfo_Photo = .strPhoto
        End With

        'add to the recent list?
        If (Not m_blnFormIsLoading) And (g_lngRecentUpdate = gc_lngRECENT_SELECTED) And (g_lngRecentStore <> gc_lngRECENT_GAMES) Then
            Call AddToRecentList(itmSelectedItem.Key, g_udtSid.strName)
        End If
        
    Else
    
        ''''''''''''''''''''''''''''''''''''''''''''
        'ITEM CLICKED IS A GAME FROM THE GAMES TABLE
        ''''''''''''''''''''''''''''''''''''''''''''
    
        'store global variable so rest of program knows what's selected
        g_eWhatIsSelected = e_SelType_IsGame
        
        'flags
        m_blnNotesByCode = True
        m_blnNotesChanged = False
    
        'menus
        mnuEditGame.Enabled = True
        mnuEditScrnShots.Enabled = True
        mnuEditExtras.Enabled = True
        mnuEditGemus.Enabled = True
        mnuEditVersion.Enabled = True
        mnuEditPersonal.Enabled = True
        mnuEditVersion.Enabled = True
        mnuEditPersonal.Enabled = True
        mnuEditLinkSet.Enabled = ((m_udtMarkedGame.strName <> "") And (m_udtMarkedGame.lngID <> lngID))
        mnuEditMark.Enabled = True

        'construct the SELECT query
        '--------------------------
        'fields
        strQuery = "SELECT " & _
            "Years.Year, Programmers.Programmer, Difficulty.Difficulty, " & _
            "Genres.Genre, Genres.GE_Id, Genres.PG_Id, Musicians.Musician, " & _
            "Musicians.Grp, Musicians.Nick, Musicians.Photo, Games.PU_Id, Games.LA_Id, Games.PR_Id, Games.YE_Id, " & _
            "Games.AR_Id, Games.DE_Id, Games.LI_Id, Games.RA_Id, " & _
            "Games.V_TitleScreen, Games.V_Playable, Games.V_Original, " & _
            "Games.WebLink_Name, Games.WebLink_URL, Games.V_WebLink_Name, Games.V_WebLink_URL, " & _
            "Games.CR_Id, Crackers.Cracker, Publishers.Publisher, Languages.Language, " & _
            "Games.DI_Id, Games.GA_Id, Games.Name, Games.Comment, Games.Rating, Games.Control, " & _
            "Games.Classic, Games.Filename, Games.FileToRun, Games.FilenameIndex, Games.Version, " & _
            "Games.SidFilename, Games.ScrnShotFilename, Games.MU_Id, " & _
            "Games.HighScore, Games.TimesPlayed, Games.DateLastPlayed, " & _
            "Games.SA, Games.FA, Games.Extras, Games.V_Comment, " & _
            "Games.PlayersTo, Games.PlayersFrom, Games.V_LoadingScreen, " & _
            "Games.V_HighScoreSaver, Games.V_IncludedDocs, Games.V_PalNTSC, " & _
            "Games.V_Length, Games.V_LengthType, Games.V_Trainers, Games.PlayersSim, Games.MemoText, Games.Gemus, " & _
            "Games.V_TrueDriveEmu, Games.Prequel, Games.Related, Games.CloneOf, Games.Sequel, Games.Adult, Games.Fav, " & _
            "Games.ReviewRating, " & _
            "Artists.Artist, Developers.Developer, Licenses.License, Rarities.Rarity, " & _
            "P.Name AS PrequelName , S.Name AS SequelName, R.Name AS RelatedName, C.Name AS CloneOfName "

        strQuery = strQuery & _
            "FROM Years " & _
            "INNER JOIN (Rarities " & _
            "INNER JOIN (Developers " & _
            "INNER JOIN (Languages " & _
            "INNER JOIN (Licenses " & _
            "INNER JOIN (Artists " & _
            "INNER JOIN (Programmers " & _
            "INNER JOIN (Publishers " & _
            "INNER JOIN (Musicians " & _
            "INNER JOIN (Genres " & _
            "INNER JOIN (Crackers " & _
            "INNER JOIN (Difficulty " & _
            "RIGHT JOIN (Games AS P " & _
            "RIGHT JOIN (Games AS S " & _
            "RIGHT JOIN (Games AS R " & _
            "RIGHT JOIN (Games AS C " & _
            "RIGHT JOIN Games "

        strQuery = strQuery & _
            "ON C.GA_Id = Games.CloneOf) " & _
            "ON R.GA_Id = Games.Related) " & _
            "ON S.GA_Id = Games.Sequel) " & _
            "ON P.GA_Id = Games.Prequel) " & _
            "ON Difficulty.DI_Id = Games.DI_Id) " & _
            "ON Crackers.CR_Id = Games.CR_Id) " & _
            "ON Genres.GE_Id = Games.GE_Id) " & _
            "ON Musicians.MU_Id = Games.MU_Id) " & _
            "ON Publishers.PU_Id = Games.PU_Id) " & _
            "ON Programmers.PR_Id = Games.PR_Id) " & _
            "ON Artists.AR_Id = Games.AR_Id) " & _
            "ON Licenses.LI_Id = Games.LI_Id) " & _
            "ON Languages.LA_Id = Games.LA_Id) " & _
            "ON Developers.DE_Id = Games.DE_Id) " & _
            "ON Rarities.RA_Id = Games.RA_Id) " & _
            "ON Years.YE_Id = Games.YE_Id "

        strQuery = strQuery & _
            "WHERE Games.GA_Id = " & lngID

        'Run the query, get the Game info from the database and store globally
        Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
        If g_rsGames.EOF Then
            MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_InvalidGameMusic", "The selected game/music has missing or invalid data in one or more of it's linked fields.") & vbCrLf & vbCrLf & "(Games)" & vbTab & "GA_Id = " & lngID & vbTab & "Name = """ & itmSelectedItem.Text & """", vbCritical, gc_strAPPNAME
            Call UnloadAllForms
            If g_blnInVBIDE = False Then
                If g_blnSubclassed = True Then
                    Call UnSubClass(Me.hwnd)
                End If
            End If
            End
        End If
        With g_rsGames
            With g_udtGame
                .lngID = lngID
                If g_blnOptionsShowGameID Then
                    Caption = g_strWINDOWCAPTION & "    <" & lngID & ">"
                End If
                sbStatus.Panels(7).Text = lngID
                .blnLoaded_0 = False
                .blnLoaded_1 = False
                .blnLoaded_2 = False
                .blnLoaded_3 = False
                .blnLoaded_Extras = False
            End With
            
            g_udtGame.strArtist = !Artist
            g_udtGame.lngArtistID = !AR_Id
            g_udtGame.strDeveloper = !Developer
            g_udtGame.lngDeveloperID = !DE_Id
            g_udtGame.strLicense = !License
            g_udtGame.lngLicenseID = !LI_Id
            g_udtGame.strRarity = !Rarity
            g_udtGame.lngRarityID = !RA_Id

            g_udtGame.lngReviewRating = !ReviewRating
            With g_udtGame
                If .lngReviewRating < -1 Then
                    .lngReviewRating = -1
                ElseIf .lngReviewRating > 100 Then
                    .lngReviewRating = 100
                End If
            End With

            g_udtGame.strWebLinkURL = !WebLink_URL
            g_udtGame.strWebLinkName = !WebLink_Name
            If Trim$(!WebLink_Name) = "" Then
                g_udtGame.strWebLinkName = g_udtGame.strWebLinkURL
            End If
            g_udtGame.blnHasWebLink = (Trim$(g_udtGame.strWebLinkURL) <> "")
            If Not g_udtGame.blnHasWebLink Then
                g_udtGame.strWebLinkName = "(" & g_strNone & ")"
            End If

            g_udtGame.strVWebLinkURL = !V_WebLink_URL
            g_udtGame.strVWebLinkName = !V_WebLink_Name
            If Trim$(!V_WebLink_Name) = "" Then
                g_udtGame.strVWebLinkName = g_udtGame.strVWebLinkURL
            End If
            g_udtGame.blnHasVWebLink = (Trim$(g_udtGame.strVWebLinkURL) <> "")
            If Not g_udtGame.blnHasVWebLink Then
                g_udtGame.strVWebLinkName = "(" & g_strNone & ")"
            End If

            g_udtGame.lngTitleScreen = !V_TitleScreen
            g_udtGame.strTitleScreen = ga_strYesNoUnknownLookupTable(!V_TitleScreen)
            g_udtGame.lngPlayable = !V_Playable
            g_udtGame.strPlayable = ga_strYesNoUnknownLookupTable(!V_Playable)
            g_udtGame.lngOriginal = !V_Original
            g_udtGame.strOriginal = ga_strYesNoUnknownLookupTable(!V_Original)

            g_udtGame.strPairs = !Gemus
            g_udtGame.blnHasExtras = !Extras
            g_udtGame.blnAdult = !Adult
            g_udtGame.blnFav = !Fav
            g_udtGame.lngDiffID = !DI_Id
            g_udtGame.lngVer = !Version
            g_udtGame.strGameFilename = !Filename
            If g_udtGame.strGameFilename = "" Then
                g_udtGame.strGameFilename = g_strNone
                g_udtGame.blnGameFileExists = False
                strVersion = ""
            Else
                strVersion = "[v" & Format$(g_udtGame.lngVer, "00") & "] "
            End If
            g_udtGame.blnGameFileExists = ((!FA = True) And (g_blnValidGamePathExists = True))
            g_udtGame.strFileToRun = !FileToRun
            g_udtGame.lngGameFilenameIndex = !FilenameIndex
            g_udtGame.strScreenshotFilename = !ScrnShotFilename
            If g_udtGame.strScreenshotFilename = "" Then
                g_udtGame.strScreenshotFilename = g_strNone
            End If
            g_udtGame.strComment = !Comment
            g_udtGame.strCracker = !Cracker
            g_udtGame.lngCrackerID = !CR_Id
            g_udtGame.strDateLastPlayed = GetFormattedDate(!DateLastPlayed)
            If g_udtGame.strDateLastPlayed = "" Then
                g_udtGame.strDateLastPlayed = g_strNever
            End If
            g_udtGame.strDifficulty = !Difficulty
    
            g_udtGame.strDocs = ga_strYesNoUnknownLookupTable(Abs(!V_IncludedDocs))
            g_udtGame.lngDocs = !V_IncludedDocs

            g_udtGame.lngGameLength = !V_Length
            g_udtGame.lngGameLengthType = !V_LengthType
            If g_udtGame.lngGameLengthType < 0 Or g_udtGame.lngGameLengthType > 19 Then
                g_udtGame.lngGameLengthType = 19 'unknown
            End If
            g_udtGame.strGameLength = g_udtGame.lngGameLength & " " & ga_strLengthTypeLookupTable(g_udtGame.lngGameLengthType)

            g_udtGame.strGenre = ga_strGenreParentLookupTable(!PG_Id) & !Genre
            g_udtGame.lngGenreID = !GE_Id
            g_udtGame.strHighScore = !HighScore
            If g_udtGame.strHighScore = "" Then
                g_udtGame.strHighScore = g_strNotEntered
            End If
            g_udtGame.strHighScoreSaver = ga_strYesNoUnknownLookupTable(Abs(!V_HighScoreSaver))
            g_udtGame.lngHSS = !V_HighScoreSaver
            g_udtGame.strLanguage = !Language
            g_udtGame.lngLanguageID = !LA_Id
            g_udtGame.strLoadingScreen = ga_strYesNoUnknownLookupTable(Abs(!V_LoadingScreen))
            g_udtGame.lngLoadingScreen = !V_LoadingScreen
            g_udtGame.blnPlayersSim = !PlayersSim
            g_udtGame.lngPlayersFrom = !PlayersFrom
            g_udtGame.lngPlayersTo = !PlayersTo
            g_udtGame.strMemo = !MemoText
            blnHasMemo = (Len(!MemoText) > 0)
            g_udtGame.strMusician = !Musician
            g_udtGame.strName = FormatGameName(!Name)
            g_udtGame.strNTSC = ga_strPalNTSCLookupTable(!V_PalNTSC)
            g_udtGame.lngNTSC = !V_PalNTSC
            g_udtGame.blnClassic = !Classic
            g_udtGame.lngPrequel = !Prequel
            If Not IsNull(!PrequelName) Then
                g_udtGame.strPrequel = !PrequelName
            Else
                g_udtGame.strPrequel = ""
            End If

            If Not IsNull(!SequelName) Then
                g_udtGame.strSequel = !SequelName
            Else
                g_udtGame.strSequel = ""
            End If
            If Not IsNull(!RelatedName) Then
                g_udtGame.strRelated = !RelatedName
            Else
                g_udtGame.strRelated = ""
            End If
            If Not IsNull(!CloneOfName) Then
                g_udtGame.strCloneOf = !CloneOfName
            Else
                g_udtGame.strCloneOf = ""
            End If
            g_udtGame.lngSequel = !Sequel
            g_udtGame.lngRelated = !Related
            g_udtGame.lngCloneOf = !CloneOf
            If !PlayersTo = !PlayersFrom Then
                If !PlayersTo = gc_lngPLAYERSUNKNOWN Then
                    g_udtGame.strPlayers = "(" & g_strUnknown & ")"
                Else
                    g_udtGame.strPlayers = !PlayersFrom & "P Only"
                End If
            Else
                If !PlayersFrom = gc_lngPLAYERSUNKNOWN Then
                    g_udtGame.strPlayers = "?"
                Else
                    g_udtGame.strPlayers = !PlayersFrom
                End If
                g_udtGame.strPlayers = g_udtGame.strPlayers & " - "
                If !PlayersTo = gc_lngPLAYERSUNKNOWN Then
                    g_udtGame.strPlayers = g_udtGame.strPlayers & "?"
                Else
                    g_udtGame.strPlayers = g_udtGame.strPlayers & !PlayersTo
                End If
            End If
            If !PlayersSim = True Then
                g_udtGame.strPlayers = g_udtGame.strPlayers & " (Simultaneous)"
            End If
            g_udtGame.blnPlayersSim = !PlayersSim
            g_udtGame.strProgrammer = !Programmer
            g_udtGame.lngProgrammerID = !PR_Id
            g_udtGame.strPublisher = !Publisher
            g_udtGame.lngPublisherID = !PU_Id
            If g_udtGame.blnClassic = False Then
                g_udtGame.strRating = ga_strRatingLookupTable(!Rating)
            Else
                g_udtGame.strRating = ga_strRatingLookupTable(6)
            End If
            g_udtGame.lngControl = !Control
            g_udtGame.strControl = ga_strControlLookupTable(!Control)
            g_udtGame.lngRating = !Rating
            g_udtGame.strTimesPlayed = !TimesPlayed
            g_udtGame.strTrueDriveEmulation = ga_strYesNoUnknownLookupTable(Abs(!V_TrueDriveEmu))
            g_udtGame.lngTrueDriveEmulation = !V_TrueDriveEmu
            If !V_Trainers < 0 Then
                g_udtGame.strTrainers = "(" & g_strUnknown & ")"
            Else
                g_udtGame.strTrainers = !V_Trainers
            End If
            g_udtGame.lngTrainers = !V_Trainers
            g_udtGame.strVersionComment = !V_Comment
            g_udtGame.strYear = ga_strYearLookupTable(!YE_ID)
            g_udtGame.lngYearID = !YE_ID
            g_udtSid.strFilename = !SidFilename
            If g_udtSid.strFilename = "" Then
                g_udtSid.strFilename = g_strNone
            End If
            g_udtSid.blnGameTable = True
            g_udtSid.blnGotInfo = False
            g_udtSid.strName = g_udtGame.strName
            g_udtSid.lngID = lngID
            g_udtSid.lngMU_Id = !MU_Id
            g_udtSid.strMusician = !Musician
            g_udtSid.strGroup = !Grp
            g_udtSid.strNick = !Nick
            g_udtSid.strPhoto = !Photo
            If g_udtSid.strGroup = "" Then g_udtSid.strGroup = g_strNone
            If g_udtSid.strNick = "" Then g_udtSid.strNick = g_strNone
            If g_udtSid.strPhoto = "" Then g_udtSid.strPhoto = g_strNone
            blnSidFileExists = LoadSidFileInfo(blnPSID)
            g_udtSid.blnFileExists = blnSidFileExists
            .Close
        End With
        Set g_rsGames = Nothing
        
        'show or clear game screenshots
        If tabDetails.Tab <> 4 Then
        
            imgPhoto.Visible = False
            m_blnPicExists = False
            If Not g_blnMamePicStyle Then
                Call LoadScreenshots(g_udtGame.strScreenshotFilename)
            Else
                Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
            End If
            Call ShowScreenshots
            m_ShowingPhoto = False
        Else
        
            Call ClearScreenshots
            picSS.Visible = False
            picSSMain(0).Visible = False
            picSSUp.Visible = False
            picSSDown.Visible = False
            With imgPhoto
                .Stretch = False
                .Visible = False
                If g_udtSid.strPhoto = g_strNone Then
                    'no photo
                    If m_blnLastPhotoWasNone = False Then
                        .Picture = imgNoPhoto.Picture
                        m_blnLastPhotoWasNone = True
                    End If
                ElseIf g_xobjFunc.FileExists(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto, False, gc_strAPPNAME) = False Then
                    'photo referenced in DB but not found in photo path
                    If m_blnLastPhotoWasNone = False Then
                        .Picture = imgNoPhoto.Picture
                        m_blnLastPhotoWasNone = True
                    End If
                Else
                    'photo found!
                    .Picture = LoadPicture(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto)
                    m_blnLastPhotoWasNone = False
                End If
                m_sngCachedPhotoHeight = .Height
                m_sngCachedPhotoWidth = .Width
                Call ShowPhoto
                m_ShowingPhoto = True

            End With
        
        End If
        
        'show the Game Info tab and hide the Sid tab
        tabSid.Visible = False
        tabSid.Enabled = False
        If m_blnRegDetails Then
            tabDetails.Visible = True
            tabDetails.Enabled = True
        End If
        If blnHasMemo Then
            tabDetails.TabPicture(3) = imgTab3Y.Picture
        Else
            tabDetails.TabPicture(3) = imgTab3N.Picture
        End If
        
        'show the game details
        With g_udtGame
            'name
            lblName.Visible = False
            lblName = .strName
            With lblName
                If g_udtGame.blnClassic Then
                    .Left = 120 + m_sngTabIconsCupGap + imgCup.Width
                    m_sngCup1X = (.Left - m_sngTabIconsCupGap) - m_sngTabIconsCupWidth
                    m_sngCup2X = .Left + .Width + m_sngTabIconsCupGap
                Else
                    .Left = 120
                End If
                .Visible = True
            End With
            'status bar
            sbStatus.Panels(1).Width = (ScaleWidth - m_sngStatusBarOffset - 800) / 3
            sbStatus.Panels(2).Width = sbStatus.Panels(1).Width
            sbStatus.Panels(3).Width = 800
            sbStatus.Panels(4).Width = sbStatus.Panels(1).Width
            sbStatus.Panels(1) = strVersion & EllipseText(.strGameFilename, sbStatus.Panels(1).Width - 320)
            sbStatus.Panels(2) = EllipseText(g_udtSid.strFilename, sbStatus.Panels(2).Width)
            If sbStatus.Panels(1) <> g_strNone Then
                sbStatus.Panels(1).ToolTipText = .strGameFilename
            Else
                sbStatus.Panels(1).ToolTipText = ""
            End If
            If sbStatus.Panels(2) <> g_udtSid.strFilename Then
                sbStatus.Panels(2).ToolTipText = g_udtSid.strFilename
            Else
                sbStatus.Panels(2).ToolTipText = ""
            End If
            'only need to show the visible tab
            cmdMemoSave.Enabled = False
            cmdMemoExpand.Enabled = True
            If g_blnLoadedfrmNotes = True Then
                frmNotes.Caption = lblName
                Call frmNotes.SetNotes(.strMemo)
            End If
            Select Case tabDetails.Tab
                Case 0
                    'Game tab
                    .blnLoaded_0 = True
                    Show0
                Case 1
                    'version tab
                    .blnLoaded_1 = True
                    Show1
                Case 2
                    'stats tab
                    .blnLoaded_2 = True
                    Show2
                Case 3
                    'memo tab
                    .blnLoaded_3 = True
                    Show3
                Case 4
                    'Sid tab
                    g_udtSid.blnGotInfo = True
                    Show4
            End Select
            'icons
            Call ShowDTIcons
            'menus
            mnuEditLinkClear.Enabled = ((.lngPrequel > 0) Or (.lngSequel > 0) Or (.lngRelated > 0) Or (.lngCloneOf > 0))
        End With

        'play game button
        cmdPlay.Enabled = (g_udtGame.blnGameFileExists And (g_xobjGEMUS.TotalEmulators > 0))

        'extras button
        cmdExtra.Enabled = (g_udtGame.blnHasExtras And (g_lngTotalExtrasPaths > 0))

        'play a random sound if the game is a classic
        If g_udtGame.blnClassic And _
            m_blnRegSounds And _
            (m_strLastSelectedKey <> itmSelectedItem.Key) And _
            mnuToolsSounds.Enabled And _
            (Not (g_blnIsFirstItemDisplayed = False And g_blnIsWavOnStart = True)) _
        Then
            Call PlayRandomSound
        End If
        m_strLastSelectedKey = itmSelectedItem.Key
        
        'add to the recent list?
        If (Not m_blnFormIsLoading) And (g_lngRecentUpdate = gc_lngRECENT_SELECTED) And (g_lngRecentStore <> gc_lngRECENT_SIDS) Then
            Call AddToRecentList(itmSelectedItem.Key, g_udtGame.strName)
        End If
        
    End If

    'release the listitem
    Set itmSelectedItem = Nothing

    'quick links window
    If m_lngRegSidebar > 0 Then
        Call ShowQuickLinks
    End If

    'play sid button
    cmdSid.Enabled = (blnSidFileExists And (g_lngTotalSidPlayers > 0))

    'menus
    mnuEditSidFile.Enabled = (g_udtRegPathSid.ePathType <> e_PathLocation_Invalid)
    mnuEditGameFile = (g_blnValidGamePathExists) And (g_lngTotalGamePaths > 0 And (g_eWhatIsSelected = e_SelType_IsGame))

    'non stop music?
    If m_blnRegNonStopMusic And cmdSid.Enabled Then
        If (m_strLastSidFilePlayed <> UCase$(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename)) Or (m_blnSettingView = False) Then
            m_strLastSidFilePlayed = UCase$(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename)
            cmdSid.Value = True
        End If
    End If
    
    g_blnIsFirstItemDisplayed = True

End Sub
Private Sub ShowScreenshots()

    '------------------------------------------------------------
    'Shows the currently loaded screenshots according to size and
    'how many can fit in the viewable window area
    '------------------------------------------------------------
    
    Dim lngScrnShotIdx As Long      'loop index
    Dim lngShowingCount As Long     'visible screenshots counter
    Dim udtSnug As SnugFitType

    picSS.Visible = (mnuViewScreenshots(0).Checked = False)
    If m_blnScrnShotNone = True Then
    
        '''''''''''''''''''''''
        'no screenshot for game
        '''''''''''''''''''''''
        
        If m_blnLastScrnShotWasNone = False Then
            'screenshot unavailable pic not already showing
            If m_lngRegScrnShots = mc_SCRNSHOTSBIG Then

                With udtSnug
                    .lngDestH = picSSMain(0).ScaleHeight
                    .lngDestW = picSSMain(0).ScaleWidth
                    .lngSourceH = picNoSS.ScaleHeight
                    .lngSourceW = picNoSS.ScaleWidth
                    Call GetSnugFitImageDimensions(udtSnug)
                    picSSMain(0).Picture = LoadPicture
                    picSSMain(0).PaintPicture picNoSS.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picNoSS.ScaleWidth, picNoSS.ScaleHeight
                End With
            Else

                With udtSnug
                    .lngDestH = picSSMain(0).ScaleHeight
                    .lngDestW = picSSMain(0).ScaleWidth
                    .lngSourceH = picNoSSSmall.ScaleHeight
                    .lngSourceW = picNoSSSmall.ScaleWidth
                    Call GetSnugFitImageDimensions(udtSnug)
                    picSSMain(0).Picture = LoadPicture
                    picSSMain(0).PaintPicture picNoSSSmall.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picNoSSSmall.ScaleWidth, picNoSSSmall.ScaleHeight
                End With
            End If
            'for next time
            m_blnLastScrnShotWasNone = True
        End If
        'show the pic and hide navigation arrows
        picSSMain(0).Visible = picSS.Visible
        picSSUp.Visible = False
        picSSDown.Visible = False
    Else
        'reset flag because this time there IS a screenshot
        m_blnLastScrnShotWasNone = False
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        'the top screenshot showing is not the first one so
        'hide the previous screenshots and show the
        'prev-scrnshot navigation button
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If m_lngScrnShotTop > 0 Then
            For lngScrnShotIdx = 0 To m_lngScrnShotTop - 1
                picSSMain(lngScrnShotIdx).Visible = False
            Next lngScrnShotIdx
            With picSSUp
                .Top = g_sngScrnShotExtraTwipsY / 2
                .Left = picSSMain(0).Left + m_sngScrnShotWidth - (m_ScrnShotNavButtonWidth * Screen.TwipsPerPixelX) - (g_sngScrnShotExtraTwipsX / 2)
                .Visible = picSS.Visible
            End With
        Else
            'the top visible screenshot is the first one so
            'hide the prev-scrnshot navigation button
            picSSUp.Visible = False
        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'show number of screenshots that fit the maximum visible range
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        lngShowingCount = 0
        For lngScrnShotIdx = m_lngScrnShotTop To m_lngScrnShotCount
            If lngShowingCount <= m_lngScrnShotMaxVisible Then
                picSSMain(lngScrnShotIdx).Top = lngShowingCount * m_sngScrnShotHeight
                picSSMain(lngScrnShotIdx).Visible = picSS.Visible
                lngShowingCount = lngShowingCount + 1
            Else
                'reached maximum showing - don't check anymore
                Exit For
            End If
        Next lngScrnShotIdx
        
        ''''''''''''''''''''''''''''''''''''''''''''
        'hide surplus screenshots and see if we need
        'to show the next-scrnshot navigation button
        ''''''''''''''''''''''''''''''''''''''''''''
        
        If (m_lngScrnShotCount - m_lngScrnShotTop) > m_lngScrnShotMaxVisible Then
            For lngScrnShotIdx = (m_lngScrnShotMaxVisible + m_lngScrnShotTop + 1) To m_lngScrnShotCount
                picSSMain(lngScrnShotIdx).Visible = False
            Next lngScrnShotIdx
            With picSSDown
                .Top = picSSMain(m_lngScrnShotMaxVisible + m_lngScrnShotTop).Top + picSSMain(m_lngScrnShotMaxVisible + m_lngScrnShotTop).Height - (m_ScrnShotNavButtonHeight * Screen.TwipsPerPixelY) - (g_sngScrnShotExtraTwipsY / 2)
                .Left = picSSMain(0).Left + m_sngScrnShotWidth - (m_ScrnShotNavButtonWidth * Screen.TwipsPerPixelX) - (g_sngScrnShotExtraTwipsX / 2)
                .Visible = picSS.Visible
            End With
        Else
            'no more screenshots so hide the next-scrnshot navigation button
            picSSDown.Visible = False
        End If

    End If

End Sub
Private Sub Shutdown()

    Visible = False
    DoEvents
    
    'shutdown any of the sidplayers
    Call ReleaseSidPlayers(g_blnShutDownSidPlayersOnExit)
    
    'save the recent menu list
    Call SaveRecentMenuList

    'save the current view
    Call SaveViewPrefs(m_lngLastViewIdx, GetListViewIdxFromViewIdx(lstView.ListIndex))
    Call UnloadAllForms
    
    Call SendMessageAny(lvMain(0).hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
    Call SendMessageAny(lvMain(1).hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
    Call SendMessageAny(lvMain(2).hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
    Call SendMessageAny(lvMain(3).hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)

    g_dbGamebase.Close

    If g_blnInVBIDE = False Then
        If g_blnSubclassed = True Then
            Call UnSubClass(Me.hwnd)
        End If
    End If
       
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "MainMaximized", True, False
        End If
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "GameTab", tabDetails.Tab, False
    End With

    Set m_xobjSQLEngine = Nothing
    Set g_xobjGEMUS = Nothing
    Set g_xobjRegistry = Nothing
    Set g_xobjClock = Nothing
    Set g_dbGamebase = Nothing
    Set g_xobjFunc = Nothing
    'Set g_xobjPNG = Nothing  'removed
    
    End

End Sub


Private Sub UpdateRecentMenu()

    Dim lngMenuIdx As Long          'loop index
    Dim lngHideIdx As Long          'loop index
    Dim lngMenuHandle As Long       'for adding bitmaps to menu
    Dim lngSubMenuHandle As Long    'for adding bitmaps to menu

    If ma_udtRecent(0).strKey = "" Then
        'no recent items
        mnuRec(0).Caption = g_strRecentEmpty
        For lngMenuIdx = 1 To 9
            With mnuRec(lngMenuIdx)
                .Visible = False
                .Enabled = False
            End With
        Next lngMenuIdx
    Else
        'show all available recent items
        For lngMenuIdx = 0 To 9
            With mnuRec(lngMenuIdx)
                If ma_udtRecent(lngMenuIdx).strKey <> "" Then
                    .Caption = " " & ma_udtRecent(lngMenuIdx).strName
                    lngMenuHandle = GetMenu(hwnd)
                    lngSubMenuHandle = GetSubMenu(lngMenuHandle, 5)
                    Call SetMenuItemBitmaps(lngSubMenuHandle, lngMenuIdx, MF_BYPOSITION, imgRecent(Abs(ma_udtRecent(lngMenuIdx).blnIsGame)).Picture, imgRecent(Abs(ma_udtRecent(lngMenuIdx).blnIsGame)).Picture)
                    .Enabled = True
                    .Visible = True
                Else
                    .Caption = ""
                    .Visible = False
                    .Enabled = False
                    Exit For
                End If
            End With
        Next lngMenuIdx
        'hide the surplus menus
        For lngHideIdx = (lngMenuIdx + 1) To 9
            With mnuRec(lngHideIdx)
                .Visible = False
                .Enabled = False
            End With
        Next lngHideIdx
    End If

End Sub

Private Sub WriteLoadUpRegSettings()

    '--------------------------------------------
    'Writes any registry settings that need to be
    'written each time GameBase loads up
    '--------------------------------------------

    Dim strDate As String

    With g_xobjRegistry
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Info", "Exe", g_strAppPathWithBackslash & App.EXEName & ".exe", True
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Info", "Version", g_strAPPVERSION & IIf((g_blnBETA), " Beta", ""), True
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "InstallDate", strDate, True, Format$(Now, "dd-mm-yyyy")
    End With

End Sub

Private Sub cmdCustViews_Click()
    
    '----------------------------------------------------
    'set the menu items depending on which view we are on
    '----------------------------------------------------
    
    'edit(1) and remove(2)
    If lstView.ListIndex > gc_lngDEFAULTVIEWSMAX Then
        pmnuView(1).Enabled = True
        pmnuView(2).Enabled = True
    Else
        pmnuView(1).Enabled = False
        pmnuView(2).Enabled = False
    End If
    
    'show the popup menu
    PopupMenu pmnuViews, , , , pmnuView(4)
    
End Sub

Private Sub cmdExtra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngMenuIdx As Long

    If g_lngTotalExtrasPaths > 0 Then
        If g_udtGame.blnHasExtras = True Then
            'load extras if needed
            If g_udtGame.blnLoaded_Extras = False Then
                'clear extras menu
                For lngMenuIdx = 1 To (pmnuExtra.Count - 1)
                    Unload pmnuExtra(lngMenuIdx)
                Next lngMenuIdx
                'read in the extras
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT EX_Id, Name, EA FROM Extras WHERE GA_Id = " & g_udtGame.lngID & " ORDER BY DisplayOrder", dbOpenSnapshot)
                With g_rsGames
                    If Not .EOF Then
                        'add to the popup menu
                        .MoveFirst
                        pmnuExtra(0).Caption = !Name
                        pmnuExtra(0).Tag = !EX_Id
                        pmnuExtra(0).Enabled = !EA
                        .MoveNext
                        lngMenuIdx = 1
                        Do Until .EOF
                            Load pmnuExtra(lngMenuIdx)
                            pmnuExtra(lngMenuIdx).Caption = !Name
                            pmnuExtra(lngMenuIdx).Tag = !EX_Id
                            pmnuExtra(lngMenuIdx).Enabled = !EA
                            lngMenuIdx = lngMenuIdx + 1
                            .MoveNext
                        Loop
                        .Close
                        g_udtGame.blnLoaded_Extras = True
                    Else
                        'no extras!
                        cmdExtra.Enabled = False
                        .Close
                        Set g_rsGames = Nothing
                        Call RemoveExtras(g_udtGame.lngID, True, False)
                    End If
                End With
                Set g_rsGames = Nothing
            End If
            'show popup menu of extras
            Call PopupMenu(pmnuExtras)
        Else
            cmdExtra.Enabled = False
        End If
    Else
        cmdExtra.Enabled = False
    End If

End Sub

Private Sub cmdMemoClear_Click()

    txtMemo = ""

End Sub



Private Sub cmdMemoExpand_Click()

    frmNotes.Caption = lblName
    Call frmNotes.SetNotes(g_udtGame.strMemo)
    frmNotes.Show vbModeless

End Sub

Private Sub cmdMemoSave_Click()

    'show activity and set button enabled to false
    cmdMemoSave.Enabled = False
    cmdMemoClear.Enabled = False
    MousePointer = vbHourglass
    DoEvents

    'save to the global variable
    g_udtGame.strMemo = txtMemo.Text

    'save the text in the memo textbox back to the database
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MemoText FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
    With g_rsGames
        .Edit
        !MemoText = g_udtGame.strMemo
        .Update
        .Close
    End With
    Set g_rsGames = Nothing
    
    If g_blnLoadedfrmNotes Then
        frmNotes.Caption = lblName
        Call frmNotes.SetNotes(g_udtGame.strMemo)
    End If

    'finished!
    cmdMemoSave.Enabled = False
    cmdMemoClear.Enabled = (txtMemo <> "")
    MousePointer = vbDefault

End Sub


Private Sub cmdPlay_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
    Call ShowTip(gc_lngTIPS_RUNGAME)
    Call RunGame(-1)

End Sub

Private Sub cmdPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '--------------------------------
    'show the emulator selection menu
    '--------------------------------

    If (Button = vbRightButton) And (cmdPlay.Enabled = True) Then
        If g_xobjGEMUS.TotalEmulators > 0 Then
            Call PopupMenu(pmnuEmulators)
        End If
    End If
    
End Sub


Private Sub cmdRandom_Click()

    Randomize
    With lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex))
        Set .SelectedItem = .ListItems(Int((.ListItems.Count - 1 + 1) * Rnd + 1))
        .SelectedItem.EnsureVisible
        ShowItem
    End With

End Sub

Private Sub cmdSid_Click()

    '------------------------------------------------------------------------
    'play the selected sid with the first valid player in the list of players
    '------------------------------------------------------------------------

    Dim lngPlayerIdx As Long    'loop index
    Dim dblProcessID As Double  'process ID of shelled player
    Dim lngTotalPlayed As Long  'for stats
    Dim lngHandle As Long

    On Error GoTo Err_Handler

    If g_lngTotalSidPlayers > 0 Then
        If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, False, "") Then
            'find first valid player
            lngPlayerIdx = GetAppropriatePlayer(g_udtSid.strFilename)
            If lngPlayerIdx > -1 Then
                With ga_udtSidPlayers(lngPlayerIdx)
                    'show tip?
                    Call ShowTip(gc_lngTIPS_PLAYSID)
                    'now shell the player
                    If Not .blnUseShortFilenames Then
                        dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, False), vbNormalFocus)
                    Else
                        dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, True), vbNormalFocus)
                    End If
                    'store the process handle
                    lngHandle = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                    Call AddToSidPlayerHandles(lngHandle)
                    'modify stats
                    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed, False, 0
                    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed + 1, False
                    'add to the recent list?
                    If (g_eWhatIsSelected = e_SelType_IsGame) And (g_blnRecentGameSid = True) Then
                        If (Not m_blnFormIsLoading) And (g_lngRecentUpdate = gc_lngRECENT_PLAYED) And (g_lngRecentStore <> gc_lngRECENT_SIDS) Then
                            Call AddToRecentList("G" & g_udtGame.lngID, g_udtGame.strName)
                        End If
                    ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
                        If (Not m_blnFormIsLoading) And (g_lngRecentUpdate = gc_lngRECENT_PLAYED) And (g_lngRecentStore <> gc_lngRECENT_GAMES) Then
                            Call AddToRecentList("M" & g_udtSid.lngID, g_udtSid.strName)
                        End If
                    End If
                End With
            Else
                If m_blnRegNonStopMusic = False Then
                    Call MsgBox(g_xobjFunc.GetINIMessage("Error_Messages", "Error_NoMusicPlayersForType", "There were no Music Players found that support this Music File.  You can set the supported file types in Options!"), vbInformation, gc_strAPPNAME)
                End If
            End If
        End If
    End If
    
    'set focus back to the listview
    lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SetFocus

    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbExclamation, gc_strAPPNAME

End Sub
Private Sub cmdSid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '---------------------------------
    'show the sidplayer selection menu
    '---------------------------------

    If (Button = vbRightButton) And (cmdSid.Enabled = True) Then
        If g_lngTotalSidPlayers > 0 Then
            Call PopupMenu(pmnuSidPlayers)
        End If
    End If

End Sub

Private Sub Form_Activate()

    Dim lngListViewIdx As Long

    'We finished loading!
    m_blnFormIsLoading = False

    If m_blnActivated Then
        'coming back to the form, set focus to the listbox
        Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
    Else
        m_blnActivated = True
    End If

End Sub

Private Sub Form_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim lngNewTabIdx As Long

    If Shift = 0 Then 'no control/shift key pressed
        If KeyCode = vbKeyEscape Then
            If (m_blnFormIsLoading = False) And (g_blnBossKeyOn = True) Then
                Unload Me
            End If
        ElseIf KeyCode > 115 And KeyCode < 121 Then 'F5 to F9 - TABDETAILS SWITCH
            With tabDetails
                If tabDetails.Visible = True And tabDetails.Enabled = True Then
                    lngNewTabIdx = KeyCode - 116
                    If lngNewTabIdx <> .Tab Then
                        If .TabVisible(lngNewTabIdx) Then .Tab = lngNewTabIdx
                    End If
                End If
            End With
        ElseIf KeyCode = 122 Then 'F11 - PREV SCREENSHOT
            If picSSUp.Visible = True And picSSUp.Enabled = True Then
                imgSSUp_Click
            End If
        ElseIf KeyCode = 123 Then 'F12 - NEXT SCREENSHOT
            If picSSDown.Visible = True And picSSDown.Enabled = True Then
                imgSSDown_Click
            End If
        End If
    ElseIf Shift = 4 Then
        If KeyCode = vbKeyReturn Then
            If WindowState = vbMaximized Then
                WindowState = vbNormal
            ElseIf WindowState = vbNormal Then
                WindowState = vbMaximized
            End If
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = 2 Then
        If KeyCode = 107 Or KeyCode = 109 Then
            'CTRL +
            'increase version of current gamefile
            'CTRL -
            'decrease version of current gamefile
            If g_eWhatIsSelected = e_SelType_IsGame Then
                If g_udtGame.lngID > 0 Then
                    If g_udtGame.strGameFilename <> g_strNone Then
                        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Version FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                        With g_rsGames
                            If Not .EOF Then
                                .MoveFirst
                                .Edit
                                If KeyCode = 107 Then
                                    If !Version < 998 Then
                                        !Version = !Version + 1
                                    End If
                                Else
                                    If !Version > 0 Then
                                        !Version = !Version - 1
                                    End If
                                End If
                                sbStatus.Panels(1) = EllipseText("[v" & !Version & "] " & g_udtGame.strGameFilename, sbStatus.Panels(1).Width)
                                .Update
                            End If
                            .Close
                        End With
                        Set g_rsGames = Nothing
                    End If
                End If
            End If
        End If
    End If

End Sub


Private Sub Form_Load()

    Dim strTextFile As String
    Dim lngTimesLoaded As Long
    Dim strFLMessage As String
    Dim blnFLGemusAsk As Boolean
    Dim lngSplash As Long
    Dim lngVCheck As Long
    Dim lngIsQuickLinksVisible As Long

    'We're loading!
    m_blnFormIsLoading = True
    m_blnDoneTabDetails = False
    m_blnActivated = False

    'create ini file if it doesn't exist
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "\GBText.ini", False, "") = False Then
        strTextFile = GetResourceFileAsString(102)
        Call CreateFile(g_strAppPathWithBackslash & "\GBText.ini", strTextFile)
    End If

    'select the database ini file if present
    Call g_xobjFunc.SetGBTextINIFile(g_strGameBasePathWBS & "GBText.ini")
    Call CreateObjects                  'create global/modular objects
    Call SetInitialGlobalValues         'set the initial values for global variables
    Call SetInitialModularValues        'set the initial values for modular variables
    Call WriteLoadUpRegSettings         'write any load-up registry settings
    Call GetINIMessages                 'get all ini messages (global and frmMain)
    Call LoadCustomGfx                  'load all the custom images
    Call SetupHelpMenu                  'setup the database specific help menu items

    'open the database
    Set g_dbGamebase = OpenDatabase(g_strGameBaseDBPath)
    'grab the database version info
    Set g_rsGames = g_dbGamebase.OpenRecordset("Config", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            g_lngDBMajorVersion = !MajorVersion
            g_lngDBMinorVersion = !MinorVersion
            g_lngOfficialDBUpdate = !OfficialUpdate
            .Close
        Else
            .Close
            g_dbGamebase.Close
            MsgBox "Missing Version/Update info in Config Table!", vbCritical, gc_strAPPNAME
            Unload frmSplash
            Set frmSplash = Nothing
            End
        End If
    End With
    Set g_rsGames = Nothing
    g_sngDBVersion = g_lngDBMajorVersion + (CSng(g_lngDBMinorVersion / 10))

    'update the database if needed
    If g_sngDBVersion < gc_sngREQ_DB_VERSION Then
        If MsgBox("GameBase needs to update your database to version " & gc_sngREQ_DB_VERSION & "." & vbCrLf & vbCrLf & "Once updated, you will no longer be able to run this database in a previous version of GameBase." & vbCrLf & vbCrLf & "If you are unsure, you should backup the file """ & g_strGameBaseDBPath & """ before continuing." & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo Or vbExclamation, gc_strAPPNAME) = vbNo Then
            Unload frmSplash
            Set frmSplash = Nothing
            End
        End If
        frmDBUpdate.Show vbModal, Me
        If g_strErrorUpdatingDatabase <> "" Then
            MsgBox "Error updating database! Your database may be corrupt!" & vbCrLf & vbCrLf & "[" & g_strErrorUpdatingDatabase & "]" & vbCrLf & vbCrLf & "A backup of your original database should be stored here: -" & vbCrLf & g_strGameBaseDBPath & ".bak" & vbCrLf & vbCrLf & "DO NOT RUN GAMEBASE AGAIN UNTIL YOU HAVE" & vbCrLf & "RENAMED THIS BACKUP DATABASE BACK TO ITS" & vbCrLf & "ORIGINAL FILENAME (" & g_strGameBaseName & ".mdb)." & vbCrLf & vbCrLf & "IF YOU DON'T RENAME YOUR DATABASE BACK, YOUR" & vbCrLf & "WORKING DATABASE MAY BE OVERWRITTEN!", vbCritical, gc_strAPPNAME
            Unload frmSplash
            Set frmSplash = Nothing
            End
        Else
            g_xobjFunc.DeletePathSilent (g_strGameBaseDBPath & ".bak")
        End If
    ElseIf g_sngDBVersion > gc_sngREQ_DB_VERSION Then
        'oops - we can't handle this newer database!
        MsgBox "You need to download a newer version of GameBase to run this database!" & vbCrLf & vbCrLf & "Download the latest version from http://www.bu22.com/", vbInformation, gc_strAPPNAME
        Unload frmSplash
        Set frmSplash = Nothing
        End
    End If

    Call LoadDatabaseConfig             'load database config (custom name, gamebase windowtitle, etc)
    Call g_xobjGEMUS.Init(g_strGameBaseName)   'Initialise the GEMUS engine
    Call ReadInRecentMenuList           'the recent menu
    Call LoadTips                       'load the main tips
    Call LoadPaths                      'load all the paths used in the program
    Call CreateLookupTables             'create the lookup tables for speed
    Call LoadDetailsTabPositions        'positions of labels/fields on details tabs
    Call SetInitialUIElements           'set the initial User Interface elements
    Call LoadScrnShotUnavailableImages  'load any Screenshot Unavailable images in the GameBase folder
    Call LoadExtraColumnsInfo           'load the second column information for the view engine
    Call CalculateMinWindowSizes        'set the minimum sizes the GB window can go
    Call SetWindowSize(m_lngRegScrnShots, m_blnRegDetails, m_sngRegWidth, m_sngRegHeight, m_lngRegSidebar) 'Set the Main GB Window size and state
    Call SetWindowPosition              'Position the GB Window on the screen
    'Call LoadDetailsTabPositions        'positions of labels/fields on details tabs
    Call LoadDefaultViews               'load in all the Default Views
    Call LoadCustomViewsFromDB          'load in all the Custom Views from the database
    Call LoadQuickFilterView            'loads the stored Quick Filter View if it exists
    Call SetSortCustomViewsEnabledState 'only allowed to sort custom views if there are any
    Call UnCacheViews                   'Un-cache all the views
    Call FillViewSelector               'Fill the View Selector with the views

    'set the view (force with gc_lngNOLISTINDEX)
    With lstView
        If g_lngRegViewIndex > .ListCount - 1 Then
            g_lngRegViewIndex = .ListCount - 1
        End If
        m_lngLastViewIdx = g_lngRegViewIndex
        .ListIndex = gc_lngNOLISTINDEX
        .ListIndex = g_lngRegViewIndex
    End With

    'Unload the splash screen
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "ShowSplash", lngSplash, False, 1
    If lngSplash > 0 Then
        If g_lngStartupWavSecondsToWait = 0 Then
            Do Until g_xobjStartupClock.TimeTaken > g_lngStartupSplashSecondsToWait
                DoEvents
            Loop
        Else
            Do Until g_xobjStartupWavClock.TimeTaken > g_lngStartupWavSecondsToWait
                DoEvents
            Loop
        End If
    End If
    Unload frmSplash
    Set frmSplash = Nothing

    'Finally show the GB window
    Show
    Visible = True

    'show the firstload message and/or gemus prompt?
    lngTimesLoaded = Val(g_xobjFunc.INI_GetSetting("Stats", "TimesLoaded", "0", g_strSettingsFile))
    If lngTimesLoaded = 0 Then
        Call GetFirstLoadStuff(strFLMessage, blnFLGemusAsk)
        If Len(strFLMessage) > 0 Then
            'first load message to show
            With g_udtFirstLoad
                .blnGemus = blnFLGemusAsk
                .strMessage = strFLMessage
                frmFirstLoad.Show vbModal, Me
                If .blnGemus = True Then
                    If mnuGemusWizard.Enabled = True Then
                        mnuGemusWizard_Click
                    End If
                End If
            End With

        ElseIf blnFLGemusAsk = True Then
            'no first load message, but we need to ask if the gemus wizard should be run
            If MsgBox(g_xobjFunc.GetINIMessage("FirstLoad Window", "GemusPrompt", "Would you like to setup your emulators now with the GEMUS Setup Wizard?"), vbYesNo Or vbQuestion, g_strGameBaseName) = vbYes Then
                If mnuGemusWizard.Enabled = True Then
                    mnuGemusWizard_Click
                End If
            End If
        End If
    End If

    'update the times loaded
    Call g_xobjFunc.INI_WriteSetting("Stats", "TimesLoaded", Trim$(lngTimesLoaded + 1), g_strSettingsFile)

    'show the version checker form?
    If g_blnSkipVersionCheck = False Then
        lngVCheck = Val(g_xobjFunc.INI_GetSetting("Startup", "VersionCheck", "-1", g_strGameBaseDotINI))
        If lngVCheck = -1 Then
            frmVersionCheck.Show vbModal, Me
        ElseIf lngVCheck = 1 Then
            sbStatus.Panels(4) = "Checking for new version..."
            DoEvents
            frmSocket.Show vbModeless, Me
        End If
    End If

End Sub

Private Sub Form_Resize()

    g_blnResizingWindow = True
    
    If m_blnResizeEnable = True Then
        If m_blnRegDetails = True Then
            ResizeWindowWithDetails
        Else
            ResizeWindowWithoutDetails
        End If
    End If
    
    g_blnResizingWindow = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Shutdown

End Sub


Private Sub fraButtons_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))

End Sub

Private Sub fraPanel_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))

End Sub

Private Sub imgAdult_Click()

    g_udtEditPersonalInfoFocus = e_PersonalFocus_Adult
    Call ShowEditPersonalInfo(False)

End Sub



Private Sub imgCloneOf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mbRes As VbMsgBoxResult

    If Button = vbRightButton Then
        
    Else
        If cmdMemoSave.Enabled = True Then
            mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
            If mbRes = vbYes Then
                cmdMemoSave.Value = True
            End If
        End If
        'find and select the game's cloneof game in the game list
        If (imgCloneOf.Tag = 1) And (g_udtGame.lngCloneOf > 0) Then
            Call SetItemInList("G" & g_udtGame.lngCloneOf, True, 0, "G" & g_udtGame.lngID)
        End If
    End If

End Sub


Private Sub imgPhoto_Click()

    If g_xobjFunc.DirExists(g_udtRegPathPhotos.strRealPath) = True Then
        Call ShowMusicianProps("Photo")
    Else
        MsgBox g_xobjFunc.GetINIMessage("Messages", "NoPhotoPath", "You must set your Photo path in 'Paths' first!"), vbInformation, gc_strAPPNAME
    End If

End Sub

Private Sub imgPrequel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mbRes As VbMsgBoxResult

    If Button = vbRightButton Then
        
    Else
        
        If cmdMemoSave.Enabled = True Then
            mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
            If mbRes = vbYes Then
                cmdMemoSave.Value = True
            End If
        End If
        
        'find and select the game's prequel in the game list
        If (imgPrequel.Tag = 1) And (g_udtGame.lngPrequel > 0) Then
            Call SetItemInList("G" & g_udtGame.lngPrequel, True, 0, "G" & g_udtGame.lngID)
        End If
    End If

End Sub


Private Sub imgRatingFace_Click()

    g_udtEditPersonalInfoFocus = e_PersonalFocus_Rating
    Call ShowEditPersonalInfo(False)

End Sub

Private Sub imgRelated_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mbRes As VbMsgBoxResult

    If Button = vbRightButton Then
        
    Else
        If cmdMemoSave.Enabled = True Then
            mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
            If mbRes = vbYes Then
                cmdMemoSave.Value = True
            End If
        End If
        'find and select the game's related game in the game list
        If (imgRelated.Tag = 1) And (g_udtGame.lngRelated > 0) Then
            Call SetItemInList("G" & g_udtGame.lngRelated, True, 0, "G" & g_udtGame.lngID)
        End If
    End If
    
End Sub


Private Sub imgRRStore_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngMin As Long
    Dim lngMax As Long

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected no.players in a quickview
            lngMin = g_udtGame.lngReviewRating - 6
            lngMax = g_udtGame.lngReviewRating + 6
            If lngMin < 0 Then lngMin = 0
            If lngMax > 100 Then lngMax = 100
            With g_udtQuickFilter
                'approx X%
                .strTitle = g_strField_ReviewRating & " ~ " & g_udtGame.lngReviewRating & "%"
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 3
                'must be > -1
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "ReviewRating"
                .strClauseData(0) = -1
                .eOperator(0) = e_Op_IsMoreThan
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                'must be more than 6 less than X%
                .strFieldTable(1) = "Games"
                .strFieldName(1) = "ReviewRating"
                .strClauseData(1) = g_udtGame.lngReviewRating - 6
                .eOperator(1) = e_Op_IsMoreThan
                .eClauseType(1) = e_ClauseType_IsDatabaseField
                .strMusicFieldTable(1) = ""
                .strMusicFieldName(1) = ""
                'must be less than 6 more than X%
                .strFieldTable(2) = "Games"
                .strFieldName(2) = "ReviewRating"
                .strClauseData(2) = g_udtGame.lngReviewRating + 6
                .eOperator(2) = e_Op_IsLessThan
                .eClauseType(2) = e_ClauseType_IsDatabaseField
                .strMusicFieldTable(2) = ""
                .strMusicFieldName(2) = ""
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditGameInfoFocus = e_InfoFocus_ReviewRating
        Call ShowEditGameInfo(False)
    End If

End Sub


Private Sub imgSequel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mbRes As VbMsgBoxResult

    If Button = vbRightButton Then

    Else
        If cmdMemoSave.Enabled = True Then
            mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
            If mbRes = vbYes Then
                cmdMemoSave.Value = True
            End If
        End If

        'find and select the game's sequel in the game list
        If (imgSequel.Tag = 1) And (g_udtGame.lngSequel > 0) Then
            Call SetItemInList("G" & g_udtGame.lngSequel, True, 0, "G" & g_udtGame.lngID)
        End If
    End If
    
End Sub


Private Sub imgSSDown_Click()

    m_lngScrnShotTop = m_lngScrnShotTop + 1
    ShowScreenshots
  
End Sub

Private Sub lbl_Info_Artist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_Artist.ToolTipText = "" Then
        strText = lbl_Info_Artist.Caption
    Else
        strText = lbl_Info_Artist.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected artist in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Artist & " = *" & g_udtGame.strArtist & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Artists"
                .strFieldName(0) = "Artist"
                .strClauseData(0) = g_udtGame.strArtist
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Artist)
    End If

End Sub


Private Sub lbl_Info_Control_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected game length in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Control & " = " & lbl_Info_Control
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "Control"
                .strClauseData(0) = g_udtGame.lngControl
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditGameInfoFocus = e_InfoFocus_Control
        Call ShowEditGameInfo(False)
    End If

End Sub


Private Sub lbl_Info_Developer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_Developer.ToolTipText = "" Then
        strText = lbl_Info_Developer.Caption
    Else
        strText = lbl_Info_Developer.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected Developer in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Developer & " = *" & g_udtGame.strDeveloper & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Developers"
                .strFieldName(0) = "Developer"
                .strClauseData(0) = g_udtGame.strDeveloper
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Developer)
    End If

End Sub


Private Sub lbl_Info_License_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_License.ToolTipText = "" Then
        strText = lbl_Info_License.Caption
    Else
        strText = lbl_Info_License.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected License in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_License & " = *" & g_udtGame.strLicense & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Licenses"
                .strFieldName(0) = "License"
                .strClauseData(0) = g_udtGame.strLicense
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_License)
    End If

End Sub







Private Sub lbl_Info_Original_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected original value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Original & " = " & lbl_Info_Original
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_Original"
                .strClauseData(0) = g_udtGame.lngOriginal
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_Original
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_Playable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected playable value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Playable & " = " & lbl_Info_Playable
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_Playable"
                .strClauseData(0) = g_udtGame.lngPlayable
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_Playable
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_Rarity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_Rarity.ToolTipText = "" Then
        strText = lbl_Info_Rarity.Caption
    Else
        strText = lbl_Info_Rarity.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected Rarity in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Rarity & " = *" & g_udtGame.strRarity & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Rarities"
                .strFieldName(0) = "Rarity"
                .strClauseData(0) = g_udtGame.strRarity
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Rarity)
    End If

End Sub



Private Sub lbl_Info_TitleScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected title screen value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_TitleScreen & " = " & lbl_Info_TitleScreen
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_TitleScreen"
                .strClauseData(0) = g_udtGame.lngTitleScreen
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_TitleScreen
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_VWebLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And g_udtGame.blnHasVWebLink Then
        Call g_xobjFunc.ShellExecuteFile(g_udtGame.strVWebLinkURL, "", False)
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_WebLink
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_WebLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And g_udtGame.blnHasWebLink Then
        Call g_xobjFunc.ShellExecuteFile(g_udtGame.strWebLinkURL, "", False)
    Else
        g_udtEditGameInfoFocus = e_InfoFocus_WebLink
        Call ShowEditGameInfo(False)
    End If

End Sub


Private Sub lbl_SFInfo_Photo_Click()

    If g_xobjFunc.DirExists(g_udtRegPathPhotos.strRealPath) = True Then
        Call ShowMusicianProps("Photo")
    Else
        MsgBox g_xobjFunc.GetINIMessage("Messages", "NoPhotoPath", "You must set your Photo path in 'Paths' first!"), vbInformation, gc_strAPPNAME
    End If

End Sub

Private Sub lbl_SInfo_Photo_Click()

    If g_xobjFunc.DirExists(g_udtRegPathPhotos.strRealPath) = True Then
        Call ShowMusicianProps("Photo")
    Else
        MsgBox g_xobjFunc.GetINIMessage("Messages", "NoPhotoPath", "You must set your Photo path in 'Paths' first!"), vbInformation, gc_strAPPNAME
    End If

End Sub




Private Sub lvLinks_DblClick()

    'run the selected quick link

    On Error GoTo Err_Handler

    Dim mbRes As VbMsgBoxResult
    Dim lngExtraIdx As Long
    Dim itmSelectedItem As ListItem
    Dim itmLink As ListItem
    Dim lngID As Long
    Dim lngLinkID As Long
    Dim strLinkType As String

    If lstView.ListIndex = gc_lngNOLISTINDEX Then Exit Sub
    
    'grab the selected item and ID
    Set itmSelectedItem = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
    If itmSelectedItem Is Nothing Then Exit Sub
    With itmSelectedItem
        lngID = Val(Right$(.Key, Len(.Key) - 1))
    End With

    'grab the selected quick link
    Set itmLink = lvLinks.SelectedItem
    If itmLink Is Nothing Then Exit Sub
    With itmLink
        strLinkType = Left$(.Key, 1)
        lngLinkID = Val(Right$(.Key, Len(.Key) - 1))
    End With

    If itmSelectedItem.Icon <> 5 And strLinkType <> "S" Then
        'game selected in main window
        If strLinkType = "G" Then
            If lngID <> lngLinkID Then Exit Sub
            If cmdPlay.Enabled Then
                cmdPlay.Value = True
            End If
        ElseIf strLinkType = "M" Then
            If lngID <> lngLinkID Then Exit Sub
            If cmdSid.Enabled Then
                cmdSid.Value = True
            End If
        ElseIf strLinkType = "W" Then
            If lngID <> lngLinkID Then Exit Sub
            If g_udtGame.blnHasWebLink = True Then
                Call g_xobjFunc.ShellExecuteFile(g_udtGame.strWebLinkURL, "", False)
            End If
        ElseIf strLinkType = "V" Then
            If lngID <> lngLinkID Then Exit Sub
            If g_udtGame.blnHasVWebLink = True Then
                Call g_xobjFunc.ShellExecuteFile(g_udtGame.strVWebLinkURL, "", False)
            End If
        ElseIf strLinkType = "E" Then
            Call RunExtra(lngLinkID)
        ElseIf strLinkType = "A" Then
            If cmdMemoSave.Enabled = True Then
                mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
                If mbRes = vbYes Then
                    cmdMemoSave.Value = True
                End If
            End If
            If g_udtGame.lngPrequel > 0 Then
                Call SetItemInList("G" & g_udtGame.lngPrequel, True, 0, "G" & g_udtGame.lngID)
                Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
            End If
        ElseIf strLinkType = "B" Then
            If cmdMemoSave.Enabled = True Then
                mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
                If mbRes = vbYes Then
                    cmdMemoSave.Value = True
                End If
            End If
            If g_udtGame.lngSequel > 0 Then
                Call SetItemInList("G" & g_udtGame.lngSequel, True, 0, "G" & g_udtGame.lngID)
                Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
            End If
        ElseIf strLinkType = "C" Then
            If cmdMemoSave.Enabled = True Then
                mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
                If mbRes = vbYes Then
                    cmdMemoSave.Value = True
                End If
            End If
            If g_udtGame.lngRelated > 0 Then
                Call SetItemInList("G" & g_udtGame.lngRelated, True, 0, "G" & g_udtGame.lngID)
                Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
            End If
        ElseIf strLinkType = "D" Then
            If cmdMemoSave.Enabled = True Then
                mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
                If mbRes = vbYes Then
                    cmdMemoSave.Value = True
                End If
            End If
            If g_udtGame.lngCloneOf > 0 Then
                Call SetItemInList("G" & g_udtGame.lngCloneOf, True, 0, "G" & g_udtGame.lngID)
                Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
            End If
        End If
        
        
    ElseIf strLinkType = "S" Then
        'music selected in main window
        If lngID <> lngLinkID Then Exit Sub
        If cmdSid.Enabled Then
            cmdSid.Value = True
        End If
    End If
    
    Exit Sub

Err_Handler:
    MsgBox "Unable to launch the selected Quick Link!" & vbCrLf & vbCrLf & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME

End Sub


Private Sub lvMain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim LItem As ListItem
    

    If Button = vbRightButton Then
        If mnuEdit.Enabled = True Then
            Set LItem = lvMain(Index).HitTest(X, Y)
            If Not (LItem Is Nothing) Then
                PopupMenu mnuEdit
            End If
        End If
    End If

End Sub

Private Sub mnuEditExtras_Click()

    Call ShowExtrasEditor

End Sub

Private Sub mnuEditFindIDGame_Click()

    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    With g_udtInputBox
        'set props
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 8
        .strCancelButtonText = g_strButtonCancelText
        .strCaption = GetINIMessage("Find ID Window", "GameCaption", "Find Game by ID")
        .strInvalidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,./<>?#';~@:][}{=+-_)(*&^%$£""!`¬\|"
        .strOKButtonText = GetINIMessage("Find ID Window", "FindButton", "&Find")
        .strPrompt = GetINIMessage("Find ID Window", "GamePrompt", "Game ID to find") & ":"
        .strReturnedText = ""
        .strStartingText = IIf((m_lngFindIDGame > 0), Trim$(m_lngFindIDGame), "")
        'show find window
        frmInputBox.Show vbModal, Me
        'find the text
        If .blnOKClicked And .strReturnedText <> "" Then
            If Val(.strReturnedText) > 0 Then
                m_lngFindIDGame = Val(.strReturnedText)
               Call FindByID(True)
            End If
        End If
    End With
    Call SetFocusToListview(lngListViewIdx)

End Sub


Private Sub mnuEditFindIDMusic_Click()

    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    With g_udtInputBox
        'set props
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 8
        .strCancelButtonText = g_strButtonCancelText
        .strCaption = GetINIMessage("Find ID Window", "MusicCaption", "Find Music by ID")
        .strInvalidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,./<>?#';~@:][}{=+-_)(*&^%$£""!`¬\|"
        .strOKButtonText = GetINIMessage("Find ID Window", "FindButton", "&Find")
        .strPrompt = GetINIMessage("Find ID Window", "MusicPrompt", "Music ID to find") & ":"
        .strReturnedText = ""
        .strStartingText = IIf((m_lngFindIDMusic > 0), Trim$(m_lngFindIDMusic), "")
        'show find window
        frmInputBox.Show vbModal, Me
        'find the text
        If .blnOKClicked And .strReturnedText <> "" Then
            If Val(.strReturnedText) > 0 Then
                m_lngFindIDMusic = Val(.strReturnedText)
               Call FindByID(False)
            End If
        End If
    End With
    Call SetFocusToListview(lngListViewIdx)

End Sub


Private Sub mnuEditGemus_Click()

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            g_UpdateKVTypeEnum = e_Save_Current
            g_strUpdateKVText = ""
            frmKeyValues.Show vbModal, Me
            'batch update?
            If g_UpdateKVTypeEnum <> e_Save_Current Then
                Load frmBatchEdit
                frmBatchEdit.m_blnUpdateKV = True
                frmBatchEdit.Show vbModal, Me
                Call ClickLstView
            End If
        End If
        Set liSelected = Nothing
    End If

End Sub

Private Sub mnuGB_Click(Index As Integer)

    'load a new instance of GameBase and close this one

    Call UnloadAllForms
    Me.Visible = False
    DoEvents
    Call Shell(g_strAppPathWithBackslash & App.EXEName & ".exe -v " & ma_strGameBaseNames(Index), vbNormalFocus)
    Unload Me

End Sub

Private Sub mnuGemusExport_Click()

    Screen.MousePointer = vbHourglass
    DoEvents
    Call g_xobjGEMUS.ExportEmulatorDefinitions
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuGemusWizard_Click()

    Dim lngListViewIdx As Long

    If g_xobjGEMUS.ShowWizard() = True Then
        Call LoadGemusMenu
        Call LoadEmulatorPopupMenu
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If Not (lvMain(lngListViewIdx).SelectedItem Is Nothing) Then
            Call SetItemInList(lvMain(lngListViewIdx).SelectedItem.Key, False)
        End If
    End If

End Sub






Private Sub mnuHelpDB_Click(Index As Integer)

    If Index = 0 Then
        If m_strDBHelpURL <> "" Then
            Call g_xobjFunc.ShellExecuteFile(m_strDBHelpFile, "", False)
        End If
    ElseIf Index = 1 Then
            If m_strDBHelpURL <> "" Then
            Call g_xobjFunc.ShellExecuteFile(m_strDBHelpURL, "", False)
        End If
    ElseIf Index = 2 Then
        If m_strDBHelpEmail <> "" Then
            Call g_xobjFunc.ShellExecuteFile("mailto:" & m_strDBHelpEmail, "", False)
        End If
    End If

End Sub


Private Sub mnuToolsDatabaseOptions_Click()

    Dim lngListViewIdx As Long

    'show the database options screen
    frmOptionsDB.Show vbModal, Me
    
    'update program
    If Not g_blnOptionsShowGameID Then
        Caption = g_strWINDOWCAPTION
    Else
        g_blnOptionsReClickGame = True
    End If
    mnuToolsVerify.Enabled = AllowVerifyFiles()
    If g_blnOptionsReLoadList = True Then
        Call UnCacheViews
        Call ClickLstView
    ElseIf g_blnOptionsReClickGame = True Then
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If lvMain(lngListViewIdx).SelectedItem Is Nothing Then
            Call SetControlsForView(False, lngListViewIdx)
        Else
            Call SetItemInList(lvMain(lngListViewIdx).SelectedItem.Key, False)
        End If
    End If

End Sub

Private Sub mnuToolsImportPersonal_Click()

    frmImportPersonal.Show vbModal, Me
    If g_blnImportPersonalInfoSuccess Then
        Call UnCacheViews
        Call ClickLstView
    End If

End Sub

Private Sub mnuToolsPaths_Click()

    Dim lngListViewIdx As Long

    'show the options screen
    frmPaths.Show vbModal, Me
    
    'update program
    mnuToolsVerify.Enabled = AllowVerifyFiles()
    If g_blnOptionsReLoadList = True Then
        Call UnCacheViews
        Call ClickLstView
    ElseIf g_blnOptionsReClickGame = True Then
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If lvMain(lngListViewIdx).SelectedItem Is Nothing Then
            Call SetControlsForView(False, lngListViewIdx)
        Else
            Call SetItemInList(lvMain(lngListViewIdx).SelectedItem.Key, False)
        End If
    End If

End Sub


Private Sub mnuViewExportToRomLister_Click()

    'export the current view to an xml file formatted for romlister
    'http://www.waste.org/~winkles/ROMLister/
    'generally used for importing game lists into arcade cabinet frontends

    Load frmRomLister
    frmRomLister.Show vbModal, Me

End Sub

Private Sub mnuViewExportToTextFile_Click()

    'export the current view to a text file

    Load frmBatchEdit
    frmBatchEdit.m_blnExportToTextFile = True
    frmBatchEdit.Show vbModal, Me

End Sub



Private Sub mnuViewExtraColumns_Click()

    Dim blnTemp As Boolean

    frmColumns.Show vbModal, Me
    If g_blnOKClicked Then
        ga_blnViewsLoaded(GetListViewIdxFromViewIdx(lstView.ListIndex)) = False
        Call SetExtraColumns(lstView.ListIndex)
        blnTemp = m_blnSaveViewPrefs
        m_blnSaveViewPrefs = False
        ClickLstView
        m_blnSaveViewPrefs = blnTemp
    End If

End Sub

Private Sub mnuViewFocusToList_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))

End Sub



Private Sub mnuViewSidebar_Click(Index As Integer)

    If m_lngRegSidebar = Index Then Exit Sub
    m_lngRegSidebar = Index
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "Sidebar", m_lngRegSidebar, False
    SetWindowSize m_lngRegScrnShots, m_blnRegDetails, Width, Height, m_lngRegSidebar
    If m_lngRegSidebar > 0 Then
        Call ShowQuickLinks
    End If

End Sub

Private Sub picSidebar_Resize()

    lvLinks.Top = 0
    lvLinks.Left = 0
    lvLinks.Width = picSidebar.ScaleWidth
    lvLinks.Height = picSidebar.ScaleHeight
    lvLinks.ColumnHeaders(1).Width = lvLinks.Width - 330

End Sub


Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_sngPreviousSplitterLeft = picSplitter.Left
    m_blnIsDraggingSidebarSplitter = True
    picSplitter.BackColor = lblSplitter.BackColor

End Sub


Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim sngNewX As Single

    If m_blnIsDraggingSidebarSplitter Then
        sngNewX = picSplitter.Left + X
        If sngNewX < mc_sngMINSIDEBARWIDTH Then
            sngNewX = mc_sngMINSIDEBARWIDTH
        ElseIf sngNewX > m_sngLastListViewSideBarAvailableWidth - mc_sngMINSIDEBARWIDTH Then
            sngNewX = m_sngLastListViewSideBarAvailableWidth - mc_sngMINSIDEBARWIDTH
        End If
        picSplitter.Left = sngNewX
    End If

End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_blnIsDraggingSidebarSplitter = False
    picSplitter.BackColor = BackColor
    If picSplitter.Left = m_sngPreviousSplitterLeft Then Exit Sub
    
    If m_lngRegSidebar = 1 Then
        'sidebar on left
        m_sngSidebarWidth = picSplitter.Left
    Else
        'sidebar on right
        m_sngSidebarWidth = m_sngLastListViewSideBarAvailableWidth - picSplitter.Left
    End If
    
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "SidebarWidth", m_sngSidebarWidth, False
    
    Call Form_Resize

End Sub


Private Sub picSS_Click()

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))

End Sub

Private Sub picSSMain_DblClick(Index As Integer)

    If m_blnScrnShotNone = False Then
        Call ShowTheBigPicture(Index, lblName)
    End If
                    
End Sub

Private Sub picSSMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetScreenshotArrows Index, X, Y, Shift

End Sub


Private Sub picSSMain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngLVIdx As Long

    Call ShowTip(gc_lngTIPS_SCREENSHOTS)
    If Button = vbLeftButton Then
        With picSSMain(Index)
            If .MousePointer = vbCustom Then
                If .MouseIcon = imgCurUp.MouseIcon Then
                    'move up
                    If m_lngScrnShotTop > 0 Then
                        m_lngScrnShotTop = m_lngScrnShotTop - 1
                        ShowScreenshots
                        SetScreenshotArrows Index - 1, X, Y, 0
                    End If
                ElseIf .MouseIcon = imgCurDown.MouseIcon Then
                    'move down
                    If picSSDown.Visible = True Then
                        m_lngScrnShotTop = m_lngScrnShotTop + 1
                        ShowScreenshots
                        SetScreenshotArrows Index + 1, X, Y, 0
                    End If
                End If
                lngLVIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
                If lvMain(lngLVIdx).Enabled = True Then
                    lvMain(lngLVIdx).SetFocus
                End If
            End If
        End With
    Else
        If Not (lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem Is Nothing) Then
            If g_eWhatIsSelected = e_SelType_IsGame Then
                If m_blnScrnShotNone = False Then
                    pmnuScrnShot(0).Enabled = True
                Else
                    pmnuScrnShot(0).Enabled = False
                End If
                m_longScrnShotZoomIdx = Index
                PopupMenu pmnuScrnShots
            End If
        End If
    End If
    
End Sub

Private Sub imgSSUp_Click()

    If m_lngScrnShotTop > 0 Then
        m_lngScrnShotTop = m_lngScrnShotTop - 1
    End If
    ShowScreenshots
    
End Sub

Private Sub lbl_Info_Comment_Click()

    g_udtEditGameInfoFocus = e_InfoFocus_Comment
    Call ShowEditGameInfo(False)

End Sub

Private Sub lbl_Info_Cracker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String
    
    If lbl_Info_Cracker.ToolTipText = "" Then
        strText = lbl_Info_Cracker.Caption
    Else
        strText = lbl_Info_Cracker.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected cracker in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Cracker & " = *" & strText & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Crackers"
                .strFieldName(0) = "Cracker"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Cracker)
    End If

End Sub

Private Sub lbl_Info_Difficulty_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected programmer in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Difficulty & " = *" & lbl_Info_Difficulty & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Difficulty"
                .strFieldName(0) = "Difficulty"
                .strClauseData(0) = lbl_Info_Difficulty
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditPersonalInfoFocus = e_PersonalFocus_Difficulty
        Call ShowEditPersonalInfo(False)
    End If

End Sub

Private Sub lbl_Info_Docs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected included docs value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_IncludedDocs & " = " & lbl_Info_Docs
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_IncludedDocs"
                .strClauseData(0) = g_udtGame.lngDocs
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_Docs
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_GameLength_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected game length in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_GameLength & " = " & lbl_Info_GameLength
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_Length"
                .strClauseData(0) = g_udtGame.lngGameLength
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField

                .strMusicFieldTable(1) = ""
                .strMusicFieldName(1) = ""
                .strFieldTable(1) = "Games"
                .strFieldName(1) = "V_LengthType"
                .strClauseData(1) = g_udtGame.lngGameLengthType
                .eOperator(1) = e_Op_Equal
                .eClauseType(1) = e_ClauseType_IsDatabaseField
                
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 2
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_Length
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_Genre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected genre in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Genre & " = " & lbl_Info_Genre
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Genres"
                .strFieldName(0) = "GE_Id"
                .strClauseData(0) = g_udtGame.lngGenreID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditGameInfoFocus = e_InfoFocus_Genre
        Call ShowEditGameInfo(False)
    End If

End Sub


Private Sub lbl_Info_HighScore_Click()
    g_udtEditPersonalInfoFocus = e_PersonalFocus_HighScore
    Call ShowEditPersonalInfo(False)
End Sub

Private Sub lbl_Info_HSS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected high score saver value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_HighScoreSaver & " = " & lbl_Info_HSS
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_HighScoreSaver"
                .strClauseData(0) = g_udtGame.lngHSS
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_HSS
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_Language_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected language in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Language & " = *" & lbl_Info_Language & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Languages"
                .strFieldName(0) = "Language"
                .strClauseData(0) = lbl_Info_Language
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Language)
    End If

End Sub


Private Sub lbl_Info_LoadingScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected loading screen value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_LoadingScreen & " = " & lbl_Info_LoadingScreen
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_LoadingScreen"
                .strClauseData(0) = g_udtGame.lngLoadingScreen
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_LoadingScreen
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_LP_Click()
    g_udtEditPersonalInfoFocus = e_PersonalFocus_LP
    Call ShowEditPersonalInfo(False)
End Sub

Private Sub lbl_Info_Musician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String
    
    If lbl_Info_Musician.ToolTipText = "" Then
        strText = lbl_Info_Musician.Caption
    Else
        strText = lbl_Info_Musician.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Musician & " = *" & strText & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Musician"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Musician"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Musician)
    End If

End Sub

Private Sub lbl_Info_NTSC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected PAL/NTSC info in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_PALNTSC & " = " & lbl_Info_NTSC
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_PalNTSC"
                .strClauseData(0) = g_udtGame.lngNTSC
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_NTSC
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_Players_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected no.players in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_NumPlayers & " = " & lbl_Info_Players
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "PlayersSim"
                .strClauseData(0) = g_udtGame.blnPlayersSim
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                If g_udtGame.lngPlayersFrom = g_udtGame.lngPlayersTo Then
                    .strMusicFieldTable(1) = ""
                    .strMusicFieldName(1) = ""
                    .strFieldTable(1) = "Games"
                    .strFieldName(1) = "PlayersFrom"
                    .strClauseData(1) = g_udtGame.lngPlayersFrom
                    .eOperator(1) = e_Op_Equal
                    .eClauseType(1) = e_ClauseType_IsDatabaseField
                    .lngNumFilters = 2
                Else
                    .strMusicFieldTable(1) = ""
                    .strMusicFieldName(1) = ""
                    .strFieldTable(1) = "Games"
                    .strFieldName(1) = "PlayersTo"
                    .strClauseData(1) = g_udtGame.lngPlayersTo
                    .eOperator(1) = e_Op_Equal
                    .eClauseType(1) = e_ClauseType_IsDatabaseField
                    .lngNumFilters = 2
                End If
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditGameInfoFocus = e_infoFocus_Players
        Call ShowEditGameInfo(False)
    End If

End Sub


Private Sub lbl_Info_Programmer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_Programmer.ToolTipText = "" Then
        strText = lbl_Info_Programmer.Caption
    Else
        strText = lbl_Info_Programmer.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected programmer in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Programmer & " = *" & strText & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Programmers"
                .strFieldName(0) = "Programmer"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Programmer)
    End If

End Sub


Private Sub lbl_Info_Publisher_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String

    If lbl_Info_Publisher.ToolTipText = "" Then
        strText = lbl_Info_Publisher.Caption
    Else
        strText = lbl_Info_Publisher.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected publisher in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Publisher & " = *" & strText & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Publishers"
                .strFieldName(0) = "Publisher"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Publisher)
    End If

End Sub

Private Sub lbl_Info_Rating_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected programmer in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Rating & " = " & lbl_Info_Rating
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                If g_udtGame.blnClassic = False Then
                    .strFieldName(0) = "Rating"
                    .strClauseData(0) = g_udtGame.lngRating
                    If g_udtGame.lngRating = 5 Then
                        .strFieldTable(1) = "Games"
                        .strFieldName(1) = "Classic"
                        .strClauseData(1) = False
                        .eOperator(1) = e_Op_Equal
                        .eClauseType(1) = e_ClauseType_IsDatabaseField
                        .lngNumFilters = 2
                    Else
                        .lngNumFilters = 1
                    End If
                Else
                    .strFieldName(0) = "Classic"
                    .strClauseData(0) = -1
                    .lngNumFilters = 1
                End If
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditPersonalInfoFocus = e_PersonalFocus_Rating
        Call ShowEditPersonalInfo(False)
    End If

End Sub


Private Sub lbl_Info_TDE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected true drive emulation value in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_TrueDriveEmulation & " = " & lbl_Info_TDE
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_TrueDriveEmu"
                .strClauseData(0) = g_udtGame.lngTrueDriveEmulation
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_TDE
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_TP_Click()
    g_udtEditPersonalInfoFocus = e_PersonalFocus_TP
    Call ShowEditPersonalInfo(False)
End Sub

Private Sub lbl_Info_Trainers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected num.trainers in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Trainers & " = " & lbl_Info_Trainers
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Games"
                .strFieldName(0) = "V_Trainers"
                .strClauseData(0) = g_udtGame.lngTrainers
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditVersionInfoFocus = e_VersionFocus_Trainers
        Call ShowEditVersionInfo(False)
    End If

End Sub


Private Sub lbl_Info_VComment_Click()

    g_udtEditVersionInfoFocus = e_InfoFocus_Comment
    Call ShowEditVersionInfo(False)

End Sub

Private Sub lbl_Info_Year_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngIdx As Long

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected year in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Year & " = *" & lbl_Info_Year & "*"
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Years"
                .strFieldName(0) = "YE_Id"
                .strClauseData(0) = 1
                For lngIdx = 1 To gc_lngYEARLOOKUPMAX
                    If lbl_Info_Year = ga_strYearLookupTable(lngIdx) Then
                        .strClauseData(0) = Trim$(lngIdx)
                        Exit For
                    End If
                Next lngIdx
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        g_udtEditGameInfoFocus = e_InfoFocus_Year
        Call ShowEditGameInfo(False)
    End If

End Sub

Private Sub lbl_SFInfo_Group_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician group in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_MusicianGroup & " = *" & lbl_SFInfo_Group & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Grp"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Grp"
                .strClauseData(0) = IIf((lbl_SFInfo_Group = g_strNone), "", lbl_SFInfo_Group)
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowMusicianProps("Group")
    End If

End Sub


Private Sub lbl_SFInfo_Musician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strText As String
    
    If lbl_SFInfo_Musician.ToolTipText = "" Then
        strText = lbl_SFInfo_Musician.Caption
    Else
        strText = lbl_SFInfo_Musician.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Musician & " = *" & strText & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Musician"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Musician"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowRecordSelector(e_RecordSelector_Musician)
    End If

End Sub

Private Sub lbl_SFInfo_Nick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician nickname in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_MusicianNick & " = *" & lbl_SFInfo_Nick & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Nick"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Nick"
                .strClauseData(0) = IIf((lbl_SFInfo_Nick = g_strNone), "", lbl_SFInfo_Nick)
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowMusicianProps("Nick")
    End If

End Sub


Private Sub lbl_SInfo_Group_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician group in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_MusicianGroup & " = *" & lbl_SInfo_Group & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Grp"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Grp"
                .strClauseData(0) = IIf((lbl_SInfo_Group = g_strNone), "", lbl_SInfo_Group)
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowMusicianProps("Group")
    End If

End Sub


Private Sub lbl_SInfo_Musician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long
    Dim strText As String

    If lbl_SInfo_Musician.ToolTipText = "" Then
        strText = lbl_SInfo_Musician.Caption
    Else
        strText = lbl_SInfo_Musician.ToolTipText
    End If

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_Musician & " = *" & strText & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Musician"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Musician"
                .strClauseData(0) = strText
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        'right-click - edit musician
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        Set liSelected = lvMain(lngListViewIdx).SelectedItem
        With g_udtRecordSelector
            .blnOKClicked = False
            .eRecordSelectorType = e_RecordSelector_Musician
            .lngID = g_udtSid.lngMU_Id
            .strSQL = "SELECT MU_Id as ID, Musician AS Name FROM Musicians ORDER BY Musician"
        End With
        Screen.MousePointer = vbHourglass
        DoEvents
        frmRecordSelector.Show vbModal, Me
        If g_udtRecordSelector.blnOKClicked Then
            'save field
            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MU_Id FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
            With g_rsGames
                If Not .EOF Then
                    .MoveFirst
                    .Edit
                    !MU_Id = g_udtRecordSelector.lngID
                    .Update
                End If
                .Close
            End With
            Set g_rsGames = Nothing
            're-click sid
            Call SetItemInList(liSelected.Key, False)
            Call RefreshExtraColumns(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem)
            Call UnCacheViews
        End If
        Screen.MousePointer = vbDefault
        Set liSelected = Nothing
    End If

End Sub

Private Sub lbl_SInfo_Nick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        If g_blnOpenFilter64 = False Then
            'left click, not using Filter64 so show all games of selected musician nickname in a quickview
            With g_udtQuickFilter
                .strTitle = g_strField_MusicianNick & " = *" & lbl_SInfo_Nick & "*"
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "Nick"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "Nick"
                .strClauseData(0) = IIf((lbl_SFInfo_Nick = g_strNone), "", lbl_SInfo_Nick)
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_FieldContainsText
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call SetQuickFilter(True)
        End If
    Else
        Call ShowMusicianProps("Nick")
    End If

End Sub



Private Sub lblName_Click()

    Dim strFormatted As String
    Dim mbRes As VbMsgBoxResult
    Dim strRecentID As String
    Dim lngIdx As Long

    With g_udtInputBox
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 228
        .strCancelButtonText = g_strButtonCancelText
        .strCaption = g_xobjFunc.GetINIMessage("Edit Name Window", "Caption", "Edit Name") & "..."
        .strInvalidChars = """"
        .strOKButtonText = g_strButtonOKText
        .strPrompt = g_xobjFunc.GetINIMessage("Edit Name Window", "Prompt", "Name") & ":"
        .strReturnedText = ""
        .strStartingText = lblName.Caption
        frmInputBox.Show vbModal, Me
        If .blnOKClicked Then
            If .strReturnedText <> lblName.Caption Then
                If g_eWhatIsSelected = e_SelType_IsGame Then
                    'check format
                    strFormatted = FormattedGameName(.strReturnedText)
                    If strFormatted <> .strReturnedText Then
                        mbRes = MsgBox("""" & .strReturnedText & """  =>  """ & strFormatted & """" & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Edit Name Window", "FormatQuestion", "Would you like to format the game name as shown above (recommended)?"), vbQuestion Or vbYesNoCancel, gc_strAPPNAME)
                        If mbRes = vbCancel Then
                            Exit Sub
                        ElseIf mbRes = vbYes Then
                            .strReturnedText = strFormatted
                        End If
                    End If
                    'save to games table
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Name FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .Edit
                            !Name = g_udtInputBox.strReturnedText
                            .Update
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    Call SetItemInList("G" & g_udtGame.lngID, False)
                    strRecentID = "G" & g_udtGame.lngID
                ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
                    'save to sids table
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Name FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .Edit
                            !Name = g_udtInputBox.strReturnedText
                            .Update
                        End If
                        .Close
                    End With
                    Call SetItemInList("M" & g_udtSid.lngID, False)
                    strRecentID = "M" & g_udtSid.lngID
                End If
                'amend the recent menu (if in there)
                For lngIdx = 0 To 9
                    If ma_udtRecent(lngIdx).strKey = strRecentID Then
                        ma_udtRecent(lngIdx).strName = g_udtInputBox.strReturnedText
                    End If
                Next lngIdx
                Call UpdateRecentMenu
                'set the new text in the list
                lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem.Text = .strReturnedText
                Call UnCacheViews
            End If
        End If
    End With

End Sub

Private Sub lstView_Click()

    '--------------------------------------------
    'Loads the selected view and selects the last
    'item that was selected in that view
    '--------------------------------------------

    Dim blnViewHasItems As Boolean      'the loaded view might be empty
    Dim lngListViewIdx As Long

    'save notes of current game if changed
    If cmdMemoSave.Enabled = True Then
        
    End If

    'You cannot unload imagebox controls from a form
    'from within a listview click() event.  For this
    'crappy reason we introduce a workaround where we
    'don't unload any surplus screenshots from the
    'context of this lstView_Click() event.
    m_blnScrnShotBugWorkaround = True

    'sometimes we need to set the view elsewhere in code
    If m_blnSetTheView = True Then
        With lstView
            If .ListIndex > gc_lngNOLISTINDEX Then
                'set the view
                blnViewHasItems = SetView(.ListIndex)
                'set the selecteditem (if any)
                If blnViewHasItems = True Then
                    m_blnSettingView = True 'for non stop music
                    Call SetItemInList(ga_udtViewData(.ListIndex).strSelectedItem, False)
                    m_blnSettingView = False 'for non stop music
                    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
                    If lvMain(lngListViewIdx).Enabled = True And lvMain(lngListViewIdx).Visible = True Then
                        lvMain(lngListViewIdx).SetFocus
                    End If
                Else
                    If m_lngRegSidebar > 0 Then
                        Call ShowQuickLinks
                    End If
                End If
            End If
        End With
    End If
    
    'end of event, now we can let the surplus
    'screenshot imageboxes  be unloaded again!
    'm_blnScrnShotBugWorkaround = False
    
End Sub

Private Sub lvMain_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Screen.MousePointer = vbHourglass
    DoEvents

    With lvMain(Index)
    
        If ColumnHeader.Index - 1 = .SortKey Then
            'sorted by current column
            If .SortOrder = lvwAscending Then
                ga_udtViewData(lstView.ListIndex).lngSortOrder = 1
                .SortOrder = lvwDescending
            Else
                ga_udtViewData(lstView.ListIndex).lngSortOrder = 0
                .SortOrder = lvwAscending
            End If
        Else
            'sorted by another column
                .Sorted = False
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
                ga_udtViewData(lstView.ListIndex).lngSortOrder = 0
                ga_udtViewData(lstView.ListIndex).lngSortColumn = .SortKey
                .Sorted = True
        End If
    
        .SelectedItem.EnsureVisible
    End With

    SetSortIcons Index
    Screen.MousePointer = vbDefault

End Sub

Private Sub lvMain_DblClick(Index As Integer)

    DoEvents

    'must have a selected game or sid
    If Not (lvMain(Index).SelectedItem Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsSid Then
            'play the sid tune
            If cmdSid.Enabled Then
                cmdSid.Value = True
            End If
        ElseIf g_eWhatIsSelected = e_SelType_IsGame Then
            If g_eDoubleClickAction = e_DoubleClick_PlaySid Then
                'play the sid tune
                If cmdSid.Enabled Then
                    cmdSid.Value = True
                End If
            Else
                'run the game
                If cmdPlay.Enabled Then
                    cmdPlay.Value = True
                End If
            End If
        End If
    End If

End Sub


Private Sub lvMain_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)

    Dim mbRes As VbMsgBoxResult

    If cmdMemoSave.Enabled = True Then
        mbRes = MsgBox("""" & lblName.Caption & """" & vbCrLf & vbCrLf & GetINIMessage("MainWindow", "NotesHaveChangedPrompt", "The Notes for this game have changed.  Do you wish to save them?"), vbQuestion Or vbYesNo, gc_strAPPNAME)
        If mbRes = vbYes Then
            cmdMemoSave.Value = True
        End If
        Call ShowItem
    Else
        Call ShowItem
    End If

End Sub





Private Sub lvMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim liSelected As ListItem
    Dim blnOldPromptFlag As Boolean
    Dim lngGPIdx As Long
    Dim strError As String

    If m_blnDragging = False Then
        Set liSelected = lvMain(Index).SelectedItem
        If Not (liSelected Is Nothing) Then
            If Button = vbLeftButton Then
                If g_eWhatIsSelected = e_SelType_IsGame Then
                    blnOldPromptFlag = g_blnPromptForGamePathDevice
                    g_blnPromptForGamePathDevice = False
                    strError = FindGameFile(lngGPIdx, g_udtGame.strGameFilename)
                    g_blnPromptForGamePathDevice = blnOldPromptFlag
                    If strError = "" Then
                        m_strDragFile = ga_udtRegPathGame(lngGPIdx).strWithBackSlash & g_udtGame.strGameFilename
                        m_blnDragging = True
                        lvMain(Index).OLEDrag
                    End If
                End If
            ElseIf Button = vbRightButton Then
                If g_eWhatIsSelected <> e_SelType_IsNone Then
                    If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, False, "") = True Then
                        m_strDragFile = g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename
                        m_blnDragging = True
                        lvMain(Index).OLEDrag
                    End If
                End If
            End If
        End If
        Set liSelected = Nothing
    End If

End Sub

Private Sub lvMain_OLECompleteDrag(Index As Integer, Effect As Long)
    
    m_blnDragging = False

End Sub

Private Sub lvMain_OLEStartDrag(Index As Integer, Data As MSComctlLib.DataObject, AllowedEffects As Long)

    AllowedEffects = vbDropEffectCopy
    Data.Clear
    Data.Files.Add m_strDragFile
    Data.SetData , vbCFFiles

End Sub

Private Sub mnuEditFav_Click()

    If lstView.ListIndex <> gc_lngFAVOURITESINDEX Then
        Call AddToFavourites
    Else
        Call RemoveFromFavourites
    End If

End Sub

Private Sub mnuEditFind_Click()

    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    With g_udtInputBox2
        'set props
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 64
        .lngCheckBoxStartingValue = m_lngFindCommentInList
        .strCheckBoxText = g_xobjFunc.GetINIMessage("Find Window", "IncludeComment", "For games, also scan &Comment field")
        .strCancelButtonText = g_strButtonCancelText
        .strCaption = g_xobjFunc.GetINIMessage("Find Window", "Caption", "Find in List")
        .strInvalidChars = ""
        .strOKButtonText = g_xobjFunc.GetINIMessage("Find Window", "FindButton", "&Find")
        .strPrompt = g_xobjFunc.GetINIMessage("Find Window", "Prompt", "&Text to find") & ":"
        .strReturnedText = ""
        .strStartingText = m_strFindInList
        'show find window
        frmInputBox2.Show vbModal, Me
        'find the text
        If .blnOKClicked And .strReturnedText <> "" Then
            m_strFindInList = .strReturnedText
            m_lngFindCommentInList = .lngCheckBoxReturnedValue
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "FindCommentInList", m_lngFindCommentInList, False
            Call FindInList(lngListViewIdx, 0)
        End If
    End With
    Call SetFocusToListview(lngListViewIdx)
    
End Sub

Private Sub mnuEditFindNext_Click()

    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    If lvMain(lngListViewIdx).SelectedItem Is Nothing Then
        mnuEditFindNext.Enabled = False
    Else
        Call FindInList(lngListViewIdx, lvMain(lngListViewIdx).SelectedItem.Index)
    End If
    Call SetFocusToListview(lngListViewIdx)

End Sub

Private Sub mnuEditGameAll_Click()

    If g_eWhatIsSelected = e_SelType_IsGame Then
        Call ShowEditGameInfo(True)
    ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
        Call ShowEditMusicInfo(True)
    End If

End Sub

Private Sub mnuEditGameCurrent_Click()
        
    If g_eWhatIsSelected = e_SelType_IsGame Then
        g_udtEditGameInfoFocus = e_InfoFocus_Year
        Call ShowEditGameInfo(False)
    ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
        g_udtEditMusicInfoFocus = e_MusicFocus_Musician
        Call ShowEditMusicInfo(False)
    End If
    
End Sub

Private Sub mnuEditGameFile_Click()

    Call ShowGameSelector(False)
    
End Sub

Private Sub mnuEditLinkClear_Click()

    Call ClearGameLink

End Sub

Private Sub mnuEditLinkSet_Click()

    Call SetGameLink

End Sub

Private Sub mnuEditMark_Click()

    '------------------------------------------------------------------------
    'Marks, or "flags" a game ready for being set as another game's prequel,
    'sequel or related game.  Shows in the 3rd panel in the status bar.
    '------------------------------------------------------------------------

    Dim liSelected As ListItem
    Dim lngListViewIdx As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    Set liSelected = lvMain(lngListViewIdx).SelectedItem
    If Not (liSelected Is Nothing) Then
        If (g_eWhatIsSelected = e_SelType_IsGame) And (g_udtGame.lngID > 0) Then
            Call ShowTip(gc_lngTIPS_GAMELINKS)
            m_udtMarkedGame.strName = g_udtGame.strName
            m_udtMarkedGame.lngID = g_udtGame.lngID
            sbStatus.Panels(4) = m_udtMarkedGame.strName
            sbStatus.Panels(4).ToolTipText = g_strMarked & ": " & m_udtMarkedGame.strName
            mnuEditLinkSet.Enabled = False
        End If
    End If
    Set liSelected = Nothing

End Sub





Private Sub mnuEditPersonalAll_Click()

    Call ShowEditPersonalInfo(True)

End Sub

Private Sub mnuEditPersonalCurrent_Click()
    
    g_udtEditPersonalInfoFocus = e_PersonalFocus_HighScore
    Call ShowEditPersonalInfo(False)

End Sub

Private Sub mnuEditScrnShots_Click()

    Call ShowScrnShotSelector

End Sub




Private Sub mnuEditSidFile_Click()

    Call ShowSidSelector(False)

End Sub

Private Sub mnuEditVersionAll_Click()

    Call ShowEditVersionInfo(True)

End Sub


Private Sub mnuEditVersionCurrent_Click()

    g_udtEditVersionInfoFocus = e_VersionFocus_Cracker
    Call ShowEditVersionInfo(False)

End Sub

Private Sub mnuFileAddGame_Click()

    Dim blnAdd As Boolean
    Dim strFormatted As String
    Dim mbRes As VbMsgBoxResult
    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim CRC As clsCRC
    
    Dim lngID As Long
    Dim strName As String
    Dim strFilename As String
    Dim strFileToRun As String
    Dim strGameName As String
    Dim lngFileIdx As Long
    Dim lngPathIdx As Long
    
    Dim Length, i, l As Integer
    Dim c As String
    Dim blnUpper As Boolean

    blnAdd = True
    strFilename = ""
    strFileToRun = ""
    lngFileIdx = 0
    strName = ""
    
    'get the game file
    If g_blnValidGamePathExists = True Then
        Call ShowGameSelector(True)
        blnAdd = g_blnGameSelectorOKPressed
        If g_blnGameSelectorOKPressed = True Then
            strFilename = g_strGameSelectorFilename
            strFileToRun = g_strGameSelectorDefaultFile
            lngFileIdx = g_strGameSelectorIndex
            
        End If
    End If
    'get the name
    If blnAdd = True Then
        With g_udtInputBox
            .blnAllowEmptyReturn = False
            .blnOKClicked = False
            .lngMaxLength = 228
            .strCancelButtonText = g_strButtonCancelText
            .strCaption = g_xobjFunc.GetINIMessage("Add Game Window", "Caption", "Add Game") & "..."
            .strInvalidChars = """"
            .strOKButtonText = g_strButtonOKText
            .strPrompt = g_xobjFunc.GetINIMessage("Add Game Window", "Prompt", "Name") & ":"
            .strReturnedText = ""
            
            strGameName = ""
            If g_strGameSelectorImageName <> "" Then
                strGameName = g_strGameSelectorImageName
            ElseIf strFileToRun <> "" Then
                strGameName = strFileToRun
            ElseIf strFilename <> "" Then
                strGameName = strFilename
            End If
            
            .strStartingText = ""
            If strGameName <> "" Then
                Length = Len(strGameName)
        
                'strip filename ending (.zip, .d64, etc.)
                If Length > 4 Then
                    For i = 0 To 4
                        If Mid$(strGameName, Length - i, 1) = "." Then
                            Length = Length - i - 1
                            Exit For
                        End If
                    Next i
                End If
        
                If Length > 0 Then
                    Length = Length - 1
                End If
        
                'first letter uppercase
                blnUpper = True
        
                For i = 0 To Length
                    c = Mid$(strGameName, i + 1, 1)
                    l = Asc(c)
                    
                    'upper/lower case
                    
                    If blnUpper = True And l <> 34 Then
                        If l > 96 And l < 123 Then
                            c = Chr$(l - 32)
                        End If
                        blnUpper = False
                    Else
                        If l = 32 Then        'uppercase after spaces
                            blnUpper = True
                        ElseIf l = 34 Then    'remove quotes
                            c = ""
                        ElseIf l > 64 And l < 91 Then
                            c = Chr$(l + 32)
                        End If
                    End If
            
                    .strStartingText = .strStartingText & c
                Next i
            End If
            
            frmInputBox.Show vbModal, Me
            If .blnOKClicked Then
                blnAdd = True
                'check format
                strFormatted = FormattedGameName(.strReturnedText)
                If strFormatted <> .strReturnedText Then
                    mbRes = MsgBox("""" & .strReturnedText & """  =>  """ & strFormatted & """" & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Add Game Window", "FormatQuestion", "Would you like to format the game name as shown above (recommended)?"), vbQuestion Or vbYesNoCancel, gc_strAPPNAME)
                    If mbRes = vbCancel Then
                        blnAdd = False
                    ElseIf mbRes = vbYes Then
                        .strReturnedText = strFormatted
                    End If
                End If
                strName = g_udtInputBox.strReturnedText
            Else
                blnAdd = False
            End If
        End With
    End If
    'add the game
    If blnAdd = True Then
        'save to database
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Games", dbOpenDynaset)
        With g_rsGames
            .AddNew
            lngID = !GA_Id
            !Name = strName
            !YE_ID = 1
            !Comment = ""
            !Filename = strFilename
            !FileToRun = strFileToRun
            !FilenameIndex = lngFileIdx
            !Version = g_lngGameSelectorVersion
            !ScrnShotFilename = ""
            !MU_Id = 38
            !GE_Id = 1
            !PU_Id = 37
            !DI_Id = 8
            !CR_Id = 8
            !SidFilename = ""
            !DateLastPlayed = ""
            !TimesPlayed = 0
            !CCode = IIf((strFilename = ""), 4, 2)
            !HighScore = ""
            !SA = False
            !FA = (strFilename <> "")
            !Fav = False
            !PR_Id = 1
            !LA_Id = 1
            !Extras = False
            !Classic = False
            !Rating = 0
            !V_LoadingScreen = False
            !V_HighScoreSaver = False
            !V_IncludedDocs = False
            !V_PalNTSC = 1
            !V_TrueDriveEmu = False
            !V_Length = 0
            !V_Trainers = -1
            !PlayersFrom = -1
            !PlayersTo = -1
            !PlayersSim = False
            !V_Comment = ""
            !Adult = False
            !MemoText = ""
            !Prequel = 0
            !Sequel = 0
            !Related = 0
            !Control = 0
            !CRC = ""
            !FileSize = 0
            .Update
            .Close
        End With
        Set g_rsGames = Nothing
        'uncache views
        Call UnCacheViews
        'add to list (including second column) and set controls for view
        If ga_udtViewData(lstView.ListIndex).eIncludeMusicTable <> e_Table_SidOnly Then
            lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
            Call SetControlsForView(True, lngListViewIdx)
            Set LItem = lvMain(lngListViewIdx).ListItems.Add(, "G" & lngID, strName, IIf((strFilename = ""), 4, 2), IIf((strFilename = ""), 4, 2))
            Call SetItemInList("G" & lngID, True)
            Call RefreshExtraColumns(lngListViewIdx, LItem)
            Set LItem = Nothing
        Else
            MsgBox g_xobjFunc.GetINIMessage("Messages", "GameAddedNoView", "The game was added, but this view is not set to show games!"), vbInformation, gc_strAPPNAME
        End If
    End If

End Sub


Private Sub mnuFileAddSid_Click()

    Dim blnAdd As Boolean
    Dim mbRes As VbMsgBoxResult
    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    
    Dim lngID As Long
    Dim strName As String
    Dim strFilename As String

    blnAdd = True
    strFilename = ""
    strName = ""
    
    'get the music file
    g_strSidSelectorFilename = g_strNone
    Call ShowSidSelector(True)
    blnAdd = (g_strSidSelectorFilename <> g_strNone)
    If g_strSidSelectorFilename <> "" Then
        strFilename = g_strSidSelectorFilename
        strName = g_strSidSelectorSidName
    End If
    'get the name
    If blnAdd = True Then
        With g_udtInputBox
            .blnAllowEmptyReturn = False
            .blnOKClicked = False
            .lngMaxLength = 228
            .strCancelButtonText = g_strButtonCancelText
            .strCaption = g_xobjFunc.GetINIMessage("Add Music Window", "Caption", "Add Tune") & "..."
            .strInvalidChars = """"
            .strOKButtonText = g_strButtonOKText
            .strPrompt = g_xobjFunc.GetINIMessage("Add Tune Window", "Prompt", "Name") & ":"
            .strReturnedText = ""
            .strStartingText = strName
            frmInputBox.Show vbModal, Me
            If .blnOKClicked Then
                blnAdd = True
                strName = g_udtInputBox.strReturnedText
            Else
                blnAdd = False
                strName = ""
            End If
        End With
    Else
        strName = ""
    End If
    'add the tune
    If blnAdd = True Then
        'save to database
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Music", dbOpenDynaset)
        With g_rsGames
            .AddNew
            lngID = !GA_Id
            !Name = strName
            !Filename = strFilename
            !MU_Id = 38
            !SA = (strFilename <> "")
            !SFav = False
            !Adult = False
            .Update
            .Close
        End With
        Set g_rsGames = Nothing
        'uncache views
        Call UnCacheViews
        'add to list (including second column) and set controls for view
        If ga_udtViewData(lstView.ListIndex).eIncludeMusicTable <> e_Table_GamesOnly Then
            lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
            Call SetControlsForView(True, lngListViewIdx)
            Set LItem = lvMain(lngListViewIdx).ListItems.Add(, "M" & lngID, strName, 5, 5)
            Call SetItemInList("M" & lngID, True)
            Call RefreshExtraColumns(lngListViewIdx, LItem)
            Set LItem = Nothing
        Else
            MsgBox g_xobjFunc.GetINIMessage("Messages", "TuneAddedNoView", "The tune was added, but this view is not set to show tunes!"), vbInformation, gc_strAPPNAME
        End If

    End If

End Sub


Private Sub mnuFileExit_Click()

    Unload Me
    
End Sub






Private Sub mnuFileRemove_Click()

    Dim lngListViewIdx As Long
    Dim strRecentKey As String
    Dim lngIdx As Long
    Dim strFields(2) As String
    Dim lngOldRelatedLink As Long

    lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
    If Not lvMain(lngListViewIdx).SelectedItem Is Nothing Then
        frmRemoveItem.Show vbModal, Me
        If g_blnRemoveItemConfirmed = True Then
            If g_eWhatIsSelected = e_SelType_IsGame Then
               
                'get related game from current game
                lngOldRelatedLink = 0
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        lngOldRelatedLink = !Related
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                                
                'remove current game from linked list by setting any games related to current game to its related link
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Related FROM Games WHERE Related = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            .Edit
                            !Related = lngOldRelatedLink
                            .Update
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing

                'get sequel of current game
                lngOldRelatedLink = 0
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        lngOldRelatedLink = !Sequel
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                
                'find any games that list current game as sequel and set their sequel to the sequel of current game
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Sequel FROM Games WHERE Sequel = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            .Edit
                            !Sequel = lngOldRelatedLink
                            .Update
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                
                'get prequel of current game
                lngOldRelatedLink = 0
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        lngOldRelatedLink = !Prequel
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                
                'find any games that list current game as prequel and set their prequel to the prequel of current game
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Prequel FROM Games WHERE Prequel = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            .Edit
                            !Prequel = lngOldRelatedLink
                            .Update
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                
                'delete current game from database
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Games WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .Delete
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                
                strRecentKey = "G" & g_udtGame.lngID
                'extras?
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Extras WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            .Delete
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
            ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
                'delete from database
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Music WHERE GA_Id = " & g_udtSid.lngID, dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        .Delete
                    End If
                    .Close
                End With
                Set g_rsGames = Nothing
                'recent menu?
                strRecentKey = "M" & g_udtSid.lngID
            End If
            'delete from recent menu
            For lngIdx = 0 To 9
                If ma_udtRecent(lngIdx).strKey = strRecentKey Then
                    Call RemoveRecentItem(lngIdx)
                    Exit For
                End If
            Next lngIdx
            'remove from list
            lvMain(lngListViewIdx).ListItems.Remove lvMain(lngListViewIdx).SelectedItem.Index
            ga_blnViewsLoaded(0) = (lngListViewIdx = 0)
            ga_blnViewsLoaded(1) = (lngListViewIdx = 1)
            ga_blnViewsLoaded(2) = (lngListViewIdx = 2)
            ga_blnViewsLoaded(3) = (lngListViewIdx = 3)
            ga_blnViewsLoaded(4) = (lngListViewIdx = 4)
            Call ClickLstView
        End If
    End If

End Sub

Private Sub mnuFileRemoveAll_Click()

    MsgBox "To do!"

End Sub

Private Sub mnuGemusEmu_Click(Index As Integer)

    If g_xobjGEMUS.EditEmulator(Index) = True Then
        Call LoadGemusMenu
        Call LoadEmulatorPopupMenu
    End If

End Sub

Private Sub mnuGemusFields_Click()

    Dim udtFields As ManualGEMUSFields

    Call g_xobjGEMUS.GetManualFieldsFromUser(udtFields, True, g_strGameBaseName)

End Sub

Private Sub mnuGemusFileTypes_Click()

    Call g_xobjGEMUS.ShowFileTypeDialog

End Sub

Private Sub mnuGemusManage_Click()

    Dim lngListViewIdx As Long

    If g_xobjGEMUS.ManageEmulators() = True Then
        Call LoadGemusMenu
        Call LoadEmulatorPopupMenu
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If Not (lvMain(lngListViewIdx).SelectedItem Is Nothing) Then
            Call SetItemInList(lvMain(lngListViewIdx).SelectedItem.Key, False)
        End If
    End If

End Sub


Private Sub mnuHelpAbout_Click()

    'show the about screen
    frmAbout.Show vbModal, Me
    
End Sub


Private Sub mnuHelpEmail_Click()

    Call g_xobjFunc.ShellExecuteFile("mailto:gamebase.frontend@leesti.com", "", False)

End Sub

Private Sub mnuHelpGamebase64_Click()

    Call g_xobjFunc.ShellExecuteFile("http://www.bu22.com/wiki/", "", False)

End Sub



Private Sub mnuHelpHome_Click()

    Call g_xobjFunc.ShellExecuteFile("http://www.bu22.com", "", False)

End Sub

Private Sub mnuRec_Click(Index As Integer)

    Dim LItem As ListItem
    Dim blnExists As Boolean
    Dim strSQL As String

    'find and select the recent game/sid in the game list
    With ma_udtRecent(Index)
        If Left$(.strKey, 1) = "G" Or Left$(.strKey, 1) = "M" Then
            If Left$(.strKey, 1) = "G" Then
                strSQL = "SELECT Fav FROM Games WHERE GA_Id = " & Right$(.strKey, Len(.strKey) - 1)
            Else
                strSQL = "SELECT SFav FROM Music WHERE GA_Id = " & Right$(.strKey, Len(.strKey) - 1)
            End If
            Set g_rsGames = g_dbGamebase.OpenRecordset(strSQL, dbOpenSnapshot)
            With g_rsGames
                blnExists = (g_rsGames.EOF = False)
                .Close
            End With
            Set g_rsGames = Nothing
            If blnExists = True Then
                Set LItem = lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem
                If LItem Is Nothing Then
                    Call SetItemInList(.strKey, True, IIf((Left$(.strKey, 1) = "G"), 0, 1))
                Else
                    Call SetItemInList(.strKey, True, IIf((Left$(.strKey, 1) = "G"), 0, 1), LItem.Key)
                End If
                Set LItem = Nothing
            Else
                'remove item as it's invalid
                Call RemoveRecentItem(Index)
                Call UpdateRecentMenu
            End If
        End If
    End With

End Sub

Private Sub mnuToolsAdultPassword_Click()

    Dim strPassword As String       'the password from the registry
    Dim strUserInfo As String       'message to user
    
    'do we have a password or not?
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\OregonASFD", "Check", strPassword, True, ""
    If strPassword <> "" Then
        'user must type the password
        frmGetPassword.Show vbModal, Me
        If g_strAdultPassword = "" Then
            'cancel pressed
            Exit Sub
        ElseIf UnMuddle(g_strAdultPassword) <> UnMuddle(strPassword) Then
            'wrong password
            strUserInfo = g_xobjFunc.GetINIMessage("Password Window", "WrongPassword", "The password you typed was not correct!")
            MsgBox strUserInfo, vbOKOnly Or vbInformation, gc_strAPPNAME
            Exit Sub
        Else
            'reset the password
            g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\OregonASFD", "Check", "", True
        End If
    End If
    strUserInfo = g_xobjFunc.GetINIMessage("Password Window", "UserInfoReset", "The password has been reset!")
    MsgBox strUserInfo, vbInformation, gc_strAPPNAME

End Sub
Private Sub mnuToolsHideAdult_Click()

    '-----------------------------------------------------
    'checks to see if the adult password is already set
    'and either prompts for it then changes the view state
    'or just asks for the already set password to be typed
    'before setting the Adult filter (or removing it)
    '-----------------------------------------------------
    
    Dim strPassword As String       'the password from the registry
    Dim strUserInfo As String       'message to user
    
    'only ask for password if showing adult games
  '  If g_blnRegHideAdultItems = True Then
        'do we have a password or not?
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\OregonASFD", "Check", strPassword, True, ""
        If strPassword = "" Then
            'get a new password and save it
            frmNewPassword.Show vbModal, Me
            If g_strAdultPassword = "" Then
                'cancel pressed
                Exit Sub
            End If
        Else
            'user must type the password
            frmGetPassword.Show vbModal, Me
            If g_strAdultPassword = "" Then
                'cancel pressed
                Exit Sub
            ElseIf UnMuddle(g_strAdultPassword) <> UnMuddle(strPassword) Then
                'wrong password
                strUserInfo = g_xobjFunc.GetINIMessage("Password Window", "WrongPassword", "The password you typed was not correct!")
                MsgBox strUserInfo, vbOKOnly Or vbInformation, gc_strAPPNAME
                Exit Sub
            End If
        End If
   ' End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'if we get here the password has been created or authorised
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'reverse the state of the filter and save to the registry
    g_blnRegHideAdultItems = Not g_blnRegHideAdultItems
    mnuToolsHideAdult.Checked = g_blnRegHideAdultItems
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\WinPlant3", "Seed", g_blnRegHideAdultItems, False
    
    're-read the current view in
    ga_blnViewsLoaded(GetListViewIdxFromViewIdx(lstView.ListIndex)) = False
    Call UnCacheViews
    Call ClickLstView
    
    'show a message to the user
    If g_blnRegHideAdultItems Then
        strUserInfo = g_xobjFunc.GetINIMessage("Password Window", "UserInfoSet", "The Adult Filter is now ON.")
    Else
        strUserInfo = g_xobjFunc.GetINIMessage("Password Window", "UserInfoUnSet", "The Adult Filter is now OFF.")
    End If
    MsgBox strUserInfo, vbOKOnly Or vbInformation, gc_strAPPNAME

End Sub

Private Sub mnuToolsNSM_Click()

    'reverse the state of the flag
    m_blnRegNonStopMusic = Not m_blnRegNonStopMusic
    mnuToolsNSM.Checked = m_blnRegNonStopMusic
    mnuToolsSounds.Enabled = Not m_blnRegNonStopMusic
    sbStatus.Panels(6).Text = IIf((m_blnRegNonStopMusic), "N", "")
    'play the selected sid if exists
    If cmdSid.Enabled And m_blnRegNonStopMusic Then
        cmdSid.Value = True
    End If

End Sub

Private Sub mnuToolsOptions_Click()

    Dim lngListViewIdx As Long

    'show the options screen
    frmOptions.Show vbModal, Me
    
    'update program
    If Not g_blnOptionsShowGameID Then
        Caption = g_strWINDOWCAPTION
    Else
        g_blnOptionsReClickGame = True
    End If
    mnuToolsVerify.Enabled = AllowVerifyFiles()
    If g_blnOptionsReLoadList = True Then
        Call UnCacheViews
        Call ClickLstView
    ElseIf g_blnOptionsReClickGame = True Then
        lngListViewIdx = GetListViewIdxFromViewIdx(lstView.ListIndex)
        If lvMain(lngListViewIdx).SelectedItem Is Nothing Then
            Call SetControlsForView(False, lngListViewIdx)
        Else
            Call SetItemInList(lvMain(lngListViewIdx).SelectedItem.Key, False)
        End If
    End If

End Sub

Private Sub mnuToolsSounds_Click()

    'reverse the state of the filter and save to the registry
    m_blnRegSounds = Not m_blnRegSounds
    mnuToolsSounds.Checked = m_blnRegSounds
    sbStatus.Panels(5).Text = IIf((m_blnRegSounds), "S", "")
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "SoundsOn", m_blnRegSounds, False

End Sub


Private Sub PlayRandomSound()

    '---------------------------------------------
    'Plays a random sound from the strSounds Array
    '
    'Remarks:   Called from ShowItem if a game is
    '           marked as a "classic game"
    '---------------------------------------------

    Dim lngSoundIdx As Integer      'array index
    Dim strSoundFile As String      'full path\filename of chosen sound
    Dim lngLowerBound As Long       'for Rnd() select
    Dim lngUpperBound As Long       'for Rnd() select
    Dim vResID As Variant
    Dim strResType As String
    
    'set an error handler just incase the API produces an error
    On Error GoTo Err_Handler
    
    If m_lngMaxSounds = 0 Then
        'no sounds to play in gamebase folder
        vResID = 101
        strResType = "CUSTOM"
        SoundArray = LoadResData(vResID, strResType)
        Call sndPlaySoundMem(SoundArray(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    Else
        'pick a random sound to play
        lngLowerBound = 1
        lngUpperBound = m_lngMaxSounds
        Randomize
        lngSoundIdx = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)
        strSoundFile = g_strAppPathWithBackslash & g_strGameBaseName & "\Sounds\" & ma_strSounds(lngSoundIdx)
        'try to play it
        Call sndPlaySound(strSoundFile, SND_NODEFAULT Or SND_ASYNC)
    End If
    
    'release the handler
    On Error GoTo 0
    
Err_Handler:
    
End Sub

Private Sub mnuToolsVerify_Click()

    frmVerifyFiles.Show vbModal, Me
    If g_blnVerifyDone Then
        Call UnCacheViews
        Call ClickLstView
    End If

End Sub

Private Sub mnuViewAddView_Click()

    Call GetCustomView(True)
    Call SetSortCustomViewsEnabledState

End Sub

Private Sub mnuViewCategories_Click()

    frmCategories.Show vbModeless, Me
    If frmCategories.WindowState = vbMinimized Then
        frmCategories.WindowState = vbNormal
    End If

End Sub

Private Sub mnuViewEditView_Click()

    Call GetCustomView(False)

End Sub


Private Sub mnuViewFontCols_Click()

    frmListFontCols.Show vbModal, Me

End Sub

Private Sub mnuViewInfo_Click()

    If mnuViewInfo.Checked = True Then
        m_blnRegDetails = False
    Else
        m_blnRegDetails = True
    End If

    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "ShowDetails", m_blnRegDetails, False
    
    SetWindowSize m_lngRegScrnShots, m_blnRegDetails, Width, Height, m_lngRegSidebar
    If mnuViewSSSnap.Enabled = True Then
        mnuViewSSSnap_Click
    End If
    
End Sub

Private Sub mnuViewListDetails_Click()

    mnuViewListList.Checked = False
    mnuViewListDetails.Checked = True
    
    ga_udtViewData(lstView.ListIndex).lngListViewType = 3
    lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).View = lvwReport

End Sub


Private Sub mnuViewListList_Click()

    mnuViewListList.Checked = True
    mnuViewListDetails.Checked = False
    
    ga_udtViewData(lstView.ListIndex).lngListViewType = 2
    'ga_udtViewData(lstView.ListIndex).lngSecondColSubIdx = -1
    'ga_udtViewData(lstView.ListIndex).lngSecondColIdx = -1
    lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).View = lvwList
    
End Sub




Private Sub mnuViewQuickFilter_Click()

    Call GetQuickFilter

End Sub

Private Sub mnuViewRemoveView_Click()

    With lstView
        If DeleteCustomView(.ListIndex) = True Then
            Call DeleteCustomViewFromSelector(.ListIndex)
            Call SetSortCustomViewsEnabledState
            'goto full main view (0)
            m_blnSaveViewPrefs = False
            .ListIndex = -1
            .ListIndex = 0
            m_blnSaveViewPrefs = True
        End If
    End With
            
End Sub







Private Sub mnuViewScreenshots_Click(Index As Integer)

    If m_lngRegScrnShots = Index Then Exit Sub
    m_lngRegScrnShots = Index
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "ScreenshotsType", m_lngRegScrnShots, False

    mnuViewSSSnap.Enabled = (m_lngRegScrnShots > 0)
  
    SetWindowSize m_lngRegScrnShots, m_blnRegDetails, Width, Height, m_lngRegSidebar
    If mnuViewSSSnap.Enabled = True Then
        mnuViewSSSnap_Click
    End If

End Sub



Private Sub mnuViewSSSnap_Click()

    'snap the vertical size of the window to the nearest screenshot height
    
    Dim sngH As Single
    Dim sngRes As Single
    Dim sngMid As Single
    Dim lngHi As Long
    Dim lngLow As Long
    Dim lngSnap As Long
    Dim lngSnap2 As Long
    Dim sngSnap As Single
    
    'find height of screenshot area
    sngH = lvMain(0).Height
    lngSnap = sngH / Screen.TwipsPerPixelY
    lngSnap2 = m_sngScrnShotHeight / Screen.TwipsPerPixelY
    If (lngSnap Mod lngSnap2) <> 0 Then
        'not currently snapped
        'find nearest snap below
        sngMid = sngH / m_sngScrnShotHeight
        lngLow = Int(sngMid)
        lngHi = lngLow + 1
        If (sngMid - CSng(lngLow)) > (CSng(lngHi) - sngMid) Then
            'snap to next screenshot
            With m_udtMinSizes
                If m_blnRegDetails = True Then
                    sngSnap = .sngScaleHeightLessSSDetails_Twips + (m_sngScrnShotHeight * lngHi) + (Height - ScaleHeight)
                    If sngSnap > Screen.Height Then
                        Height = .sngScaleHeightLessSSDetails_Twips + (m_sngScrnShotHeight * lngLow) + (Height - ScaleHeight)
                    Else
                        Height = sngSnap
                    End If
                Else
                    sngSnap = .sngScaleHeightLessSSNoDetails_Twips + (m_sngScrnShotHeight * lngHi) + (Height - ScaleHeight)
                    If sngSnap > Screen.Height Then
                        Height = .sngScaleHeightLessSSNoDetails_Twips + (m_sngScrnShotHeight * lngLow) + (Height - ScaleHeight)
                    Else
                        Height = sngSnap
                    End If
                End If
            End With
        Else
            'snap to previous screenshot
            With m_udtMinSizes
                If m_blnRegDetails = True Then
                    Height = .sngScaleHeightLessSSDetails_Twips + (m_sngScrnShotHeight * lngLow) + (Height - ScaleHeight)
                Else
                    Height = .sngScaleHeightLessSSNoDetails_Twips + (m_sngScrnShotHeight * lngLow) + (Height - ScaleHeight)
                End If
            End With
        End If
    End If

End Sub

Private Sub pmnuEmulator_Click(Index As Integer)

    Call RunGame(Index)

End Sub

Private Sub pmnuExtra_Click(Index As Integer)

    Call SetFocusToListview(GetListViewIdxFromViewIdx(lstView.ListIndex))
    Call RunExtra(pmnuExtra(Index).Tag)

End Sub

Private Sub pmnuScrnShot_Click(Index As Integer)

    Dim strFilename As String

    If Not (lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SelectedItem Is Nothing) Then
        If g_eWhatIsSelected = e_SelType_IsGame Then
            If Index = 0 Then
                If m_blnScrnShotNone = False Then
                    Call ShowTheBigPicture(m_longScrnShotZoomIdx, lblName)
                End If
            Else
                Call ShowScrnShotSelector
            End If
        End If
    End If

End Sub

Private Sub pmnuSidPlayer_Click(Index As Integer)

    'shell the sidplayer
    Call PlaySidWithSpecificPlayer(Index)

    'set focus back to the listview
    lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).SetFocus
    
End Sub

Private Sub pmnuView_Click(Index As Integer)

    Dim lngCurrentLstViewIdx As Long

    Select Case Index
        
        Case 0  'Add new view
            Call GetCustomView(True)
            Call SetSortCustomViewsEnabledState
        
        Case 1  'Edit Current View
            Call GetCustomView(False)
        
        Case 2  'Remove Custom View
            With lstView
                If DeleteCustomView(.ListIndex) = True Then
                    Call DeleteCustomViewFromSelector(.ListIndex)
                    Call SetSortCustomViewsEnabledState
                    'goto full main view (0)
                    m_blnSaveViewPrefs = False
                    .ListIndex = -1
                    .ListIndex = 0
                    m_blnSaveViewPrefs = True
                End If
            End With
            
        Case 3  'Separator Bar
   
        Case 4  'Quick Filter
            Call GetQuickFilter
        
        Case 5  'Separator Bar
        
        Case 6
            g_blnViewsWereSorted = False
            frmCVSort.Show vbModal, Me
            If g_blnViewsWereSorted = True Then
                'refresh the views
                lngCurrentLstViewIdx = lstView.ListIndex
                Call LoadDefaultViews
                Call LoadCustomViewsFromDB
                Call LoadQuickFilterView
                Call SetSortCustomViewsEnabledState
                Call UnCacheViews
                Call FillViewSelector
            
                'set the view (force with gc_lngNOLISTINDEX)
                With lstView
                    If lngCurrentLstViewIdx > .ListCount - 1 Then
                        lngCurrentLstViewIdx = .ListCount - 1
                    End If
                    m_lngLastViewIdx = lngCurrentLstViewIdx
                    .ListIndex = gc_lngNOLISTINDEX
                    .ListIndex = lngCurrentLstViewIdx
                End With
            End If

    End Select

End Sub

Private Sub sbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)

    Select Case Panel.Index
        Case 1 'game file selector
            Call ShowGameSelector(False)
        Case 2 'sid file selector
            Call ShowSidSelector(False)
        Case 3 'Extras Selector
            Call ShowExtrasEditor
        Case 5 'sounds
            mnuToolsSounds_Click
        Case 6 'non stop music
            mnuToolsNSM_Click
    End Select

End Sub

Private Sub tabDetails_Click(PreviousTab As Integer)

    If tabDetails.Enabled = False Then Exit Sub

    If m_blnDoneTabDetails = True Then
        Call ShowTip(gc_lngTIPS_DETAILSTABS)
    Else
        m_blnDoneTabDetails = True
    End If
    Select Case tabDetails.Tab
        Case 0
            If g_udtGame.blnLoaded_0 = False Then
                Show0
            End If
            If m_ShowingPhoto Then
                imgPhoto.Visible = False
                m_blnPicExists = False
                If Not g_blnMamePicStyle Then
                    Call LoadScreenshots(g_udtGame.strScreenshotFilename)
                Else
                    Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
                End If
                Call ShowScreenshots
                m_ShowingPhoto = False
            End If
        Case 1
            If g_udtGame.blnLoaded_1 = False Then
                Show1
            End If
            If m_ShowingPhoto Then
                imgPhoto.Visible = False
                m_blnPicExists = False
                If Not g_blnMamePicStyle Then
                    Call LoadScreenshots(g_udtGame.strScreenshotFilename)
                Else
                    Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
                End If
                Call ShowScreenshots
                m_ShowingPhoto = False
            End If
        Case 2
            If g_udtGame.blnLoaded_2 = False Then
                Show2
            End If
            If m_ShowingPhoto Then
                imgPhoto.Visible = False
                m_blnPicExists = False
                If Not g_blnMamePicStyle Then
                    Call LoadScreenshots(g_udtGame.strScreenshotFilename)
                Else
                    Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
                End If
                Call ShowScreenshots
                m_ShowingPhoto = False
            End If
        Case 3
            If g_udtGame.blnLoaded_3 = False Then
                Show3
            End If
            If m_ShowingPhoto Then
                imgPhoto.Visible = False
                m_blnPicExists = False
                If Not g_blnMamePicStyle Then
                    Call LoadScreenshots(g_udtGame.strScreenshotFilename)
                Else
                    Call LoadScreenshotsMAME(g_udtGame.strScreenshotFilename)
                End If
                Call ShowScreenshots
                m_ShowingPhoto = False
            End If
        Case 4
            If g_udtSid.blnGotInfo = False Then
                Show4
            End If
            If Not m_ShowingPhoto Then
                Call ClearScreenshots
                picSS.Visible = False
                picSSMain(0).Visible = False
                picSSUp.Visible = False
                picSSDown.Visible = False
                With imgPhoto
                    .Stretch = False
                    .Visible = False
                    If g_udtSid.strPhoto = g_strNone Then
                        'no photo
                        If m_blnLastPhotoWasNone = False Then
                            .Picture = imgNoPhoto.Picture
                            m_blnLastPhotoWasNone = True
                        End If
                    ElseIf g_xobjFunc.FileExists(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto, False, gc_strAPPNAME) = False Then
                        'photo referenced in DB but not found in photo path
                        If m_blnLastPhotoWasNone = False Then
                            .Picture = imgNoPhoto.Picture
                            m_blnLastPhotoWasNone = True
                        End If
                    Else
                        'photo found!
                        .Picture = LoadPicture(g_udtRegPathPhotos.strWithBackSlash & g_udtSid.strPhoto)
                        m_blnLastPhotoWasNone = False
                    End If
                    m_sngCachedPhotoHeight = .Height
                    m_sngCachedPhotoWidth = .Width
                    Call ShowPhoto
                End With
                m_ShowingPhoto = True
            End If
    End Select
    
End Sub

Private Sub txtMemo_Change()
    
    cmdMemoSave.Enabled = True
    cmdMemoClear.Enabled = (txtMemo <> "")
    If txtMemo <> "" Then
        tabDetails.TabPicture(3) = imgTab3Y.Picture
    Else
        tabDetails.TabPicture(3) = imgTab3N.Picture
    End If

End Sub


Private Sub txtMemo_GotFocus()

    mnuFileRemove.Enabled = False
    
    m_blnNotesCut = mnuEditLinkClear.Enabled
    m_blnNotesCopy = mnuEditMark.Enabled
    m_blnNotesPaste = mnuEditLinkSet.Enabled
    
    mnuEditLinkClear.Enabled = False
    mnuEditLinkSet.Enabled = False
    mnuEditMark.Enabled = False

End Sub


Private Sub txtMemo_LostFocus()

    If lvMain(GetListViewIdxFromViewIdx(lstView.ListIndex)).ListItems.Count > 0 Then
        mnuFileRemove.Enabled = True
        mnuEditLinkClear.Enabled = m_blnNotesCut
        mnuEditLinkSet.Enabled = m_blnNotesPaste
        mnuEditMark.Enabled = m_blnNotesCopy
    End If

End Sub


