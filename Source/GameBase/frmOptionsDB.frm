VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptionsDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Options"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmOptionsDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   77
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   76
      Top             =   7200
      Width           =   1095
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "G&UI"
      TabPicture(0)   =   "frmOptionsDB.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraUI"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSSP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tabDetails"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Music Players"
      TabPicture(1)   =   "frmOptionsDB.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSidPlayers"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Games"
      TabPicture(2)   =   "frmOptionsDB.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraZip"
      Tab(2).Control(1)=   "fraGP"
      Tab(2).ControlCount=   2
      Begin TabDlg.SSTab tabDetails 
         Height          =   2775
         Left            =   360
         TabIndex        =   30
         Top             =   3840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Game"
         TabPicture(0)   =   "frmOptionsDB.frx":091E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstField(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstField(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lstField(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lstField(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lstField(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lstField(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lstField(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lstField(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lstField(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lstField(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lstField(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lstField(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Version"
         TabPicture(1)   =   "frmOptionsDB.frx":093A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstField(23)"
         Tab(1).Control(1)=   "lstField(22)"
         Tab(1).Control(2)=   "lstField(21)"
         Tab(1).Control(3)=   "lstField(20)"
         Tab(1).Control(4)=   "lstField(19)"
         Tab(1).Control(5)=   "lstField(18)"
         Tab(1).Control(6)=   "lstField(17)"
         Tab(1).Control(7)=   "lstField(16)"
         Tab(1).Control(8)=   "lstField(15)"
         Tab(1).Control(9)=   "lstField(14)"
         Tab(1).Control(10)=   "lstField(13)"
         Tab(1).Control(11)=   "lstField(12)"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "Personal"
         TabPicture(2)   =   "frmOptionsDB.frx":0956
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstField(28)"
         Tab(2).Control(1)=   "lstField(27)"
         Tab(2).Control(2)=   "lstField(25)"
         Tab(2).Control(3)=   "lstField(26)"
         Tab(2).Control(4)=   "lstField(24)"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Music"
         TabPicture(3)   =   "frmOptionsDB.frx":0972
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstField(32)"
         Tab(3).Control(1)=   "lstField(31)"
         Tab(3).Control(2)=   "lstField(29)"
         Tab(3).Control(3)=   "lstField(30)"
         Tab(3).ControlCount=   4
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   32
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   31
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   29
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   30
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   28
            Left            =   -73080
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2220
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   27
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1500
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   25
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1500
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   26
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   660
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   24
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   660
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   23
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   22
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1920
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   21
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   20
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   19
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   18
            Left            =   -71520
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   17
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   16
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1920
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   15
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   14
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   13
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   12
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   11
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   10
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1920
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   9
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   8
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   7
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   6
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   5
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   4
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1920
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   3
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox lstField 
            Height          =   315
            Index           =   0
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.Frame fraSSP 
         Caption         =   "Pictures"
         Height          =   1695
         Left            =   360
         TabIndex        =   8
         Top             =   1980
         Width           =   6735
         Begin VB.CheckBox chkImageExtras 
            Caption         =   "Also show Picture E&xtras"
            Height          =   255
            Left            =   3960
            TabIndex        =   29
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtZH 
            Height          =   285
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   28
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSH 
            Height          =   285
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   23
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtNH 
            Height          =   285
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtZW 
            Height          =   285
            Left            =   5160
            MaxLength       =   5
            TabIndex        =   26
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSW 
            Height          =   285
            Left            =   5160
            MaxLength       =   5
            TabIndex        =   21
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtNW 
            Height          =   285
            Left            =   5160
            MaxLength       =   5
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chkZoomFront 
            Caption         =   "Always bring &zoomed pictures to front"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   3015
         End
         Begin VB.OptionButton optPictures 
            Caption         =   "GameBase Sty&le"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton optPictures 
            Caption         =   "MAME &Style"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblZoomed 
            Caption         =   "Zoomed"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblSmall 
            Caption         =   "Small"
            Height          =   255
            Left            =   3960
            TabIndex        =   19
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblNormal 
            Caption         =   "Normal"
            Height          =   255
            Left            =   3960
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblH 
            Caption         =   "H:"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   27
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblH 
            Caption         =   "H:"
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   22
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblH 
            Caption         =   "H:"
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblW 
            Caption         =   "W:"
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   25
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblW 
            Caption         =   "W:"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   20
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblW 
            Caption         =   "W:"
            Height          =   255
            Index           =   0
            Left            =   4920
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblPicBack 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            MouseIcon       =   "frmOptionsDB.frx":098E
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Pictures BackColor"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   12
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame fraSidPlayers 
         Caption         =   "Music Players"
         Height          =   6015
         Left            =   -74640
         TabIndex        =   64
         Top             =   540
         Width           =   6735
         Begin GameBase.ctlExeSelector ctlSidPlayers 
            Height          =   5460
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Width           =   6255
            _extentx        =   11033
            _extenty        =   9631
         End
      End
      Begin VB.Frame fraGP 
         Caption         =   "Game Paths"
         Height          =   855
         Left            =   -74640
         TabIndex        =   66
         Top             =   540
         Width           =   6735
         Begin VB.CheckBox chkGamePathPrompt 
            Caption         =   "Prompt for &insertion of CD/Removable drives if not present"
            Height          =   495
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.Frame fraUI 
         Caption         =   "General"
         Height          =   1275
         Left            =   360
         TabIndex        =   1
         Top             =   540
         Width           =   6735
         Begin VB.OptionButton optDoubleClick 
            Caption         =   "Play game in &default emulator"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   6
            Top             =   540
            Width           =   2655
         End
         Begin VB.OptionButton optDoubleClick 
            Caption         =   "Play t&une for selected game"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   7
            Top             =   840
            Width           =   2655
         End
         Begin VB.CheckBox chkBoldTab2 
            Caption         =   "Bold Comment on &Tab 2"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   540
            Width           =   2175
         End
         Begin VB.CheckBox chkBoldTab1 
            Caption         =   "Bold Comm&ent on Tab 1"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chkShowID 
            Caption         =   "S&how Game ID in Window Title Bar"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblDoubleClickPrompt 
            Caption         =   "Double Click Action:"
            Height          =   255
            Left            =   3720
            TabIndex        =   5
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame fraZip 
         Caption         =   "Re-Zipping (supported archives: .7z .zip .gz .gzip .tgz .bz2 .bzip2 .tbz2 .tbz .tar)"
         Height          =   2175
         Left            =   -74640
         TabIndex        =   68
         Top             =   1680
         Width           =   6735
         Begin VB.CheckBox chkReZip 
            Caption         =   "R&e-Zip un-zipped games after playing (if changed)"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   360
            Width           =   6255
         End
         Begin VB.CheckBox chkRezipAsk 
            Caption         =   "&Ask before rezipping"
            Height          =   255
            Left            =   480
            TabIndex        =   70
            Top             =   720
            Width           =   6015
         End
         Begin VB.OptionButton optRezipIgnore 
            Caption         =   "&Ignore them"
            Height          =   255
            Left            =   720
            TabIndex        =   73
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton optRezipZip 
            Caption         =   "Re&zip them"
            Height          =   255
            Left            =   2160
            TabIndex        =   74
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton optRezipAsk 
            Caption         =   "As&k me!"
            Height          =   255
            Left            =   3600
            TabIndex        =   75
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkRezipNotify 
            Caption         =   "&Notify me on a successful Rezip"
            Height          =   255
            Left            =   480
            TabIndex        =   71
            Top             =   1080
            Width           =   6015
         End
         Begin VB.Label lblRezipPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "When there are more files to Rezip than were unzipped:"
            Height          =   255
            Left            =   480
            TabIndex        =   72
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   6015
         End
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   120
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOptionsDB"
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

Private m_blnOKPressed As Boolean
Private m_lngOldNumPlayers As Long
Private m_strOldPlayers() As String
Private m_blnSetOnce As Boolean
Private m_strHide As String
Private Sub SaveOptions()

    '-------------------------------------------------
    'Saves all the edited options from memory/controls
    '-------------------------------------------------

    Dim lngIdx As Long              'loop index
    Dim strSidExe As String         'exe file for sid players
    Dim strSidName As String        'name for sid players
    Dim strSidCLP As String         'command line for sid players
    Dim blnUseShort As Boolean      'short filename flag for emulators and sid players
    Dim strScript As String         'for emulator script
    Dim strFiles() As String        'for emulator files
    Dim lngNumFiles As Long         'for emulator file count
    Dim lngFileIdx As Long          'loop index
    Dim strName As String           'for emulator name
    Dim strExe As String            'for emulator path\filename
    Dim a_strTypes() As String
    Dim lngTotalTypes As Long
    Dim strTypes As String
    Dim lngFTIdx As Long
    Dim blnWritten As Boolean
    Dim lngTotalPlayers As Long
    Dim blnFound As Boolean
    Dim lngPlayerIdx As Long
    Dim strTemp As String
    Dim lngNW As Long
    Dim lngNH As Long
    Dim lngSW As Long
    Dim lngSH As Long
    Dim lngZW As Long
    Dim lngZH As Long

    g_blnOptionsReClickGame = True

    '===========
    'general tab
    '===========

    'double click action
    For lngIdx = e_DoubleClick_PlayGame To e_DoubleClick_PlaySid
        If optDoubleClick(lngIdx).Value = True Then
            g_eDoubleClickAction = lngIdx
            Exit For
        End If
    Next lngIdx
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "DoubleClickAction", g_eDoubleClickAction, False

    'game id in window title bar
    blnWritten = g_xobjFunc.INI_WriteSetting("UI", "ShowGameID", chkShowID.Value, g_strSettingsFile)
    If Not blnWritten Then
        MsgBox "Could not save the UI details to Settings.ini!", vbExclamation Or vbOKOnly, gc_strAPPNAME
    End If
    g_blnOptionsShowGameID = (chkShowID.Value = vbChecked)

    'bold comments
    If g_xobjFunc.INI_WriteSetting("UI", "Tab1BoldComment", Trim$(chkBoldTab1.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'Tab1BoldComment' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    If g_xobjFunc.INI_WriteSetting("UI", "Tab2BoldComment", Trim$(chkBoldTab2.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'Tab1BoldComment' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    frmMain.lbl_Info_Comment.FontBold = (chkBoldTab1.Value = vbChecked)
    frmMain.lbl_Info_VComment.FontBold = (chkBoldTab2.Value = vbChecked)

    'pictures backcolor
    g_lngPictureBackColor = lblPicBack.BackColor
    If g_xobjFunc.INI_WriteSetting("UI", "PicturesBackColor", Trim$(g_lngPictureBackColor), g_strSettingsFile) = False Then
        MsgBox "Could not save 'PicturesBackColor' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    frmMain.picSSMain(0).BackColor = g_lngPictureBackColor
    frmMain.picFull.BackColor = g_lngPictureBackColor

    'zoom pic to front
    g_blnSetPicZoomForeground = (chkZoomFront = 1)
    If g_xobjFunc.INI_WriteSetting("General", "BringZoomedPicToFront", Trim$(chkZoomFront.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'BringZoomedPicToFront' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If

    'also load image extras
    g_blnLoadImageExtrasAsScreenshots = (chkImageExtras = 1)
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "LoadImageExtrasAsScreenshots", g_blnLoadImageExtrasAsScreenshots, False

    'picture style
    g_blnMamePicStyle = (optPictures(1).Value = True)
    If g_xobjFunc.INI_WriteSetting("General", "MamePictureStyle", Trim$(Abs(g_blnMamePicStyle)), g_strSettingsFile) = False Then
        MsgBox "Could not save 'MamePictureStyle' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    g_blnMamePicStyle = (g_xobjFunc.INI_GetSetting("General", "MamePictureStyle", "0", g_strSettingsFile) = "1")

    'picture sizes
    lngNW = CLng(txtNW.Text)
    lngNH = CLng(txtNH.Text)
    lngSW = CLng(txtSW.Text)
    lngSH = CLng(txtSH.Text)
    lngZW = CLng(txtZW.Text)
    lngZH = CLng(txtZH.Text)
    If lngNW < 1 Then lngNW = 1
    If lngNH < 1 Then lngNH = 1
    If lngSW < 1 Then lngSW = 1
    If lngSH < 1 Then lngSH = 1
    If lngZW < 1 Then lngZW = 1
    If lngZH < 1 Then lngZH = 1
    If _
        lngNW <> g_lngFullSizePictureWidthPixels Or _
        lngNH <> g_lngFullSizePictureHeightPixels Or _
        lngSW <> g_lngSmallPictureWidthPixels Or _
        lngSH <> g_lngSmallPictureHeightPixels Or _
        lngZW <> g_lngZoomedPictureWidthPixels Or _
        lngZH <> g_lngZoomedPictureHeightPixels _
    Then
        MsgBox "You must restart " & gc_strAPPNAME & " for the new Picture sizes to take effect.", vbInformation, gc_strAPPNAME
    End If
    g_lngFullSizePictureWidthPixels = lngNW
    g_lngFullSizePictureHeightPixels = lngNH
    g_lngSmallPictureWidthPixels = lngSW
    g_lngSmallPictureHeightPixels = lngSH
    g_lngZoomedPictureWidthPixels = lngZW
    g_lngZoomedPictureHeightPixels = lngZH
    With g_xobjFunc
        .INI_WriteSetting "UI", "FullSizePictureWidthPixels", Trim$(g_lngFullSizePictureWidthPixels), g_strSettingsFile
        .INI_WriteSetting "UI", "FullSizePictureHeightPixels", Trim$(g_lngFullSizePictureHeightPixels), g_strSettingsFile
        .INI_WriteSetting "UI", "SmallPictureWidthPixels", Trim$(g_lngSmallPictureWidthPixels), g_strSettingsFile
        .INI_WriteSetting "UI", "SmallPictureHeightPixels", Trim$(g_lngSmallPictureHeightPixels), g_strSettingsFile
        .INI_WriteSetting "UI", "ZoomedWidthPixels", Trim$(g_lngZoomedPictureWidthPixels), g_strSettingsFile
        .INI_WriteSetting "UI", "ZoomedPictureHeightPixels", Trim$(g_lngZoomedPictureHeightPixels), g_strSettingsFile
    End With

    'details tabs label positions
    For lngIdx = 0 To 32
        With lstField(lngIdx)
            If g_xobjFunc.INI_WriteSetting("UI", "DetailsTabField" & lngIdx & "Pos", Trim$(.ItemData(.ListIndex)), g_strSettingsFile) = False Then
                MsgBox "Could not save '""DetailsTabField" & lngIdx & "Pos"" to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
            End If
        End With
    Next lngIdx
    Call frmMain.LoadDetailsTabPositions

    '================
    'musicplayers tab
    '================
    
    blnWritten = True
    With ctlSidPlayers
        g_lngTotalSidPlayers = .ListCount
        ReDim ga_udtSidPlayers(0)
        If g_lngTotalSidPlayers > 0 Then
            ReDim ga_udtSidPlayers(g_lngTotalSidPlayers - 1)
        End If
        For lngIdx = 0 To (g_lngTotalSidPlayers - 1)
            'get item and save to array
            Call .GetItem(lngIdx, strSidName, strSidExe, strSidCLP, blnUseShort, a_strTypes(), lngTotalTypes)
            Call GetSidPlayerInfo(strSidCLP, blnUseShort, strSidName, strSidExe, ga_udtSidPlayers(lngIdx), a_strTypes(), lngTotalTypes)
            'write to the players list
            blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting("Players", Trim$(lngIdx + 1), strSidName, g_strMusicFile)
            'write the player info
            blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting(strSidName, "UseShortFilenames", IIf((blnUseShort = True), "1", "0"), g_strMusicFile)
            blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting(strSidName, "PlayerPathFile", strSidExe, g_strMusicFile)
            blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting(strSidName, "CommandLineParams", strSidCLP, g_strMusicFile)
            strTypes = ""
            For lngFTIdx = 0 To (lngTotalTypes - 1)
                strTypes = strTypes & ga_udtSidPlayers(lngIdx).a_strTypes(lngFTIdx) & ";"
            Next lngFTIdx
            blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting(strSidName, "SupportedTypes", strTypes, g_strMusicFile)
        Next lngIdx
        'delete extra unwanted emulators
        blnWritten = blnWritten And g_xobjFunc.INI_DeleteNumberedValues("Players", g_lngTotalSidPlayers + 1, g_strMusicFile)
        If blnWritten = True Then
            blnFound = False
            For lngIdx = 0 To (m_lngOldNumPlayers - 1)
                For lngPlayerIdx = 0 To (.ListCount - 1)
                    If UCase$(m_strOldPlayers(lngIdx)) = UCase$(.List(lngPlayerIdx)) Then
                        blnFound = True
                        Exit For
                    End If
                Next lngPlayerIdx
                If blnFound = False Then
                    g_xobjFunc.INI_DeleteSection m_strOldPlayers(lngIdx), g_strMusicFile
                End If
            Next lngIdx
        Else
            MsgBox "There was an error saving the Music Player info to " & g_strMusicFile, vbExclamation, gc_strAPPNAME
        End If
    End With

    Call frmMain.LoadSidPlayerPopupMenu

    '=========
    'GAMES TAB
    '=========

    'game path device prompt
    g_blnPromptForGamePathDevice = (chkGamePathPrompt = 1)
    If g_xobjFunc.INI_WriteSetting("General", "RequestGameDeviceInsert", Trim$(chkGamePathPrompt.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'RequestGameDeviceInsert' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If

    'game re-zipping
    g_blnReZip = (chkReZip = 1)
    If g_xobjFunc.INI_WriteSetting("General", "ReZipGames", Trim$(chkReZip.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'ReZipGames' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    g_blnRezipAsk = (chkRezipAsk = 1)
    If g_xobjFunc.INI_WriteSetting("General", "ReZipAsk", Trim$(chkRezipAsk.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'ReZipAsk' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    g_blnZipNotify = (chkRezipNotify = 1)
    If g_xobjFunc.INI_WriteSetting("General", "ReZipNotify", Trim$(chkRezipNotify.Value), g_strSettingsFile) = False Then
        MsgBox "Could not save 'ReZipNotify' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If
    If optRezipIgnore.Value = True Then
        ge_RezipAdditional = eRezipAdditionalIgnore
        strTemp = "0"
    ElseIf optRezipZip.Value = True Then
        ge_RezipAdditional = eRezipAdditionalZip
        strTemp = "1"
    Else
        ge_RezipAdditional = eRezipAdditionalAsk
        strTemp = "2"
    End If
    If g_xobjFunc.INI_WriteSetting("General", "ReZipAdditional", strTemp, g_strSettingsFile) = False Then
        MsgBox "Could not save 'ReZipAdditional' to " & g_strSettingsFile, vbExclamation, gc_strAPPNAME
    End If

End Sub

Private Sub chkReZip_Click()

    Call SetRezipControls

End Sub


Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdOK_Click()

    Dim lngIdx As Long
    Dim a_blnValues(32) As Long
    Dim lngValue As Long

    'check that there are no duplicate details tab label positions
    For lngIdx = 0 To 32
        a_blnValues(lngIdx) = False
    Next lngIdx
    For lngIdx = 0 To 32
        With lstField(lngIdx)
            lngValue = .ItemData(.ListIndex)
            If lngValue > -1 Then
                If a_blnValues(lngValue) = True Then
                    MsgBox g_xobjFunc.GetINIMessage( _
                        "Database Options Window", _
                        "DetailsTabsError", _
                        "You can only show each field on the details tabs once. Please check" _
                    ) & " """ & .List(.ListIndex) & """", vbExclamation, gc_strAPPNAME
                    Exit Sub
                End If
                a_blnValues(lngValue) = True
            End If
        End With
    Next lngIdx

    'save the changed options and close window
    Call SaveOptions
    m_blnOKPressed = True
    Unload Me

End Sub


Private Sub Form_Load()


    Dim lngIdx As Long                          'loop index
    Dim lngLabelIdx As Long
    Const DesignScaleHeight As Single = 7695    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 7695     'The required scalewidth for the form
    Dim strSupportedArchives As String
    Dim strReZip As String

    m_blnSetOnce = False

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        'with shortcuts
        tabOptions.TabCaption(0) = .GetINIMessage("Database Options Window", "Tab_Title_GUI", "G&UI")
        tabOptions.TabCaption(1) = .GetINIMessage("Database Options Window", "Tab_Title_Musicplayers", "&Music Players")
        tabOptions.TabCaption(2) = .GetINIMessage("Database Options Window", "Tab_Title_Games", "&Games")
        optDoubleClick(0).Caption = .GetINIMessage("Database Options Window", "DoubleClickEmulator", "Play game in &default emulator")
        optDoubleClick(1).Caption = .GetINIMessage("Database Options Window", "DoubleClickMusicplayer", "Play t&une for selected game")
        chkZoomFront.Caption = .GetINIMessage("Database Options Window", "BringZoomedPicToFront", "Always bring &zoomed pictures to front")
        chkBoldTab1.Caption = .GetINIMessage("Database Options Window", "BoldCommentTab1", "Bold Comm&ent on Tab 1")
        chkBoldTab2.Caption = .GetINIMessage("Database Options Window", "BoldCommentTab2", "Bold Comment on &Tab 2")
        optPictures(0).Caption = .GetINIMessage("Database Options Window", "GameBaseStylePics", "GameBase &Style")
        optPictures(1).Caption = .GetINIMessage("Database Options Window", "MameStylePics", "MAME Sty&le")
        chkGamePathPrompt.Caption = .GetINIMessage("Database Options Window", "InsertPrompt", "Prompt for &insertion of CD/Removable drives if not present")
        chkReZip.Caption = .GetINIMessage("Database Options Window", "Repack", "R&epack unpacked games after playing (if changed)")
        chkRezipAsk.Caption = .GetINIMessage("Database Options Window", "RepackAsk", "&Ask before repacking")
        optRezipIgnore.Caption = .GetINIMessage("Database Options Window", "RepackRadioIgnore", "&Ignore them")
        optRezipZip.Caption = .GetINIMessage("Database Options Window", "RepackRadioZip", "Re&pack them")
        optRezipAsk.Caption = .GetINIMessage("Database Options Window", "RepackRadioAsk", "As&k me!")
        chkRezipNotify.Caption = .GetINIMessage("Database Options Window", "RepackNotify", "&Notify me on a successful repack")
        chkShowID.Caption = .GetINIMessage("Database Options Window", "ShowGameID", "S&how Game ID in Window Title Bar")

        'without shortcuts
        Caption = g_strGameBaseName & " " & .GetINIMessage("Database Options Window", "Caption", "Options")
        lblPrompt(2).Caption = .GetINIMessage("Database Options Window", "PicturesBackColor", "Pictures BackColor")
        lblDoubleClickPrompt.Caption = .GetINIMessage("Database Options Window", "DoubleClickAction", "Double Click Action")
        fraUI.Caption = .GetINIMessage("Database Options Window", "General", "General")
        fraGP.Caption = .GetINIMessage("Database Options Window", "GamePaths", "Game Paths")
        fraSidPlayers.Caption = .GetINIMessage("Database Options Window", "MusicPlayers", "Music Players")
        strReZip = .GetINIMessage("Database Options Window", "RepackTitle", "Repacking")
        fraSSP.Caption = .GetINIMessage("Database Options Window", "PicTitle", "Pictures")
        lblRezipPrompt = .GetINIMessage("Database Options Window", "RepackPrompt", "When there are more files to repack than were unpacked") & ":"
        lblNormal.Caption = .GetINIMessage("Database Options Window", "Normal", "Normal")
        lblSmall.Caption = .GetINIMessage("Database Options Window", "Small", "Small")
        lblZoomed.Caption = .GetINIMessage("Database Options Window", "Zoomed", "Zoomed")
        lblW(0).Caption = .GetINIMessage("Database Options Window", "Width", "W") & ":"
        lblH(0).Caption = .GetINIMessage("Database Options Window", "Height", "H") & ":"
        lblW(1).Caption = lblW(0).Caption
        lblW(2).Caption = lblW(0).Caption
        lblH(1).Caption = lblH(0).Caption
        lblH(2).Caption = lblH(0).Caption
        tabDetails.TabCaption(0) = " " & GetINIMessage("MainWindow", "Tab_Game_Title", "Game")
        tabDetails.TabCaption(1) = " " & GetINIMessage("MainWindow", "Tab_Version_Title", "Version")
        tabDetails.TabCaption(2) = " " & GetINIMessage("MainWindow", "Tab_Personal_Title", "Personal")
        tabDetails.TabCaption(3) = " " & GetINIMessage("MainWindow", "Tab_Music_Title", "Music")
        m_strHide = "(" & .GetINIMessage("Database Options Window", "Hide", "Hide") & ")"
        strSupportedArchives = .GetINIMessage("Database Options Window", "SupportedArchives", "supported archives")

        fraZip.Caption = strReZip & " (" & strSupportedArchives & ": .7z .zip .gz .gzip .tgz .bz2 .bzip2 .tbz2 .tbz .tar)"

    End With

    'lstField listboxes
    For lngIdx = 0 To 5
        With lstField(lngIdx)
            For lngLabelIdx = 0 To 5
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(6) = -1
        End With
    Next lngIdx
    For lngIdx = 6 To 11
        With lstField(lngIdx)
            For lngLabelIdx = 6 To 11
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx - 6) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(6) = -1
        End With
    Next lngIdx
    For lngIdx = 12 To 17
        With lstField(lngIdx)
            For lngLabelIdx = 12 To 17
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx - 12) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(6) = -1
        End With
    Next lngIdx
    For lngIdx = 18 To 23
        With lstField(lngIdx)
            For lngLabelIdx = 18 To 23
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx - 18) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(6) = -1
        End With
    Next lngIdx
    For lngIdx = 24 To 28
        With lstField(lngIdx)
            For lngLabelIdx = 24 To 28
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx - 24) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(5) = -1
        End With
    Next lngIdx
    For lngIdx = 29 To 32
        With lstField(lngIdx)
            For lngLabelIdx = 29 To 32
                .AddItem frmMain.lblPrompt(lngLabelIdx)
                .ItemData(lngLabelIdx - 29) = lngLabelIdx
            Next lngLabelIdx
            .AddItem m_strHide
            .ItemData(4) = -1
        End With
    Next lngIdx

    'load the options
    Call LoadOptions
    
    'set the rezip controls
    Call SetRezipControls

    'nothing changed yet
    g_blnOptionsReClickGame = False
    g_blnOptionsReLoadList = False
    g_blnChangedSidplayerSettings = False
    g_blnChangedEmulatorSettings = False
    m_blnOKPressed = False

End Sub

Private Sub SetRezipControls()

    chkRezipAsk.Enabled = (chkReZip.Value = vbChecked)
    chkRezipNotify.Enabled = chkRezipAsk.Enabled
    optRezipIgnore.Enabled = chkRezipAsk.Enabled
    optRezipZip.Enabled = chkRezipAsk.Enabled
    optRezipAsk.Enabled = chkRezipAsk.Enabled
    lblRezipPrompt.Enabled = chkRezipAsk.Enabled

End Sub



Private Sub LoadOptions()

    '------------------------------------------------------
    'Loads all the options into memory/controls for editing
    '------------------------------------------------------

    Dim lngPathIdx As Long          'loop index
    Dim strFileTypes As String      'native emulator file types
    Dim lngStars As Long            'stars in about screen
    Dim lngStarIdx As Long          'loop index
    Dim strSeconds As String
    Dim lngIdx As Long
    Dim lngLabelIdx As Long

    '======
    'UI tab
    '======

    'double click action
    If (g_eDoubleClickAction < e_DoubleClick_PlayGame) Or (g_eDoubleClickAction > e_DoubleClick_PlaySid) Then
        g_eDoubleClickAction = e_DoubleClick_PlayGame
    End If
    optDoubleClick(g_eDoubleClickAction).Value = True

    'bold comments
    chkBoldTab1.Value = Abs(frmMain.lbl_Info_Comment.FontBold = True)
    chkBoldTab2.Value = Abs(frmMain.lbl_Info_VComment.FontBold = True)

    'show game id
    chkShowID.Value = IIf((g_blnOptionsShowGameID), vbChecked, vbUnchecked)

    'pictures backcolor
    lblPicBack.BackColor = g_lngPictureBackColor

    'zoom pic to front
    chkZoomFront = Abs(g_blnSetPicZoomForeground)

    'also load image extras
    chkImageExtras = Abs(g_blnLoadImageExtrasAsScreenshots)

    'picture style
    optPictures(Abs(g_blnMamePicStyle)).Value = True

    'picture sizes
    txtNW.Text = g_lngFullSizePictureWidthPixels
    txtNH.Text = g_lngFullSizePictureHeightPixels
    txtSW.Text = g_lngSmallPictureWidthPixels
    txtSH.Text = g_lngSmallPictureHeightPixels
    txtZW.Text = g_lngZoomedPictureWidthPixels
    txtZH.Text = g_lngZoomedPictureHeightPixels

    'details tab labels
    For lngIdx = 0 To 5
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx
    For lngIdx = 6 To 11
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx - 6
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx
    For lngIdx = 12 To 17
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx - 12
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx
    For lngIdx = 18 To 23
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx - 18
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx
    For lngIdx = 24 To 28
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx - 24
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx
    For lngIdx = 29 To 32
        lngLabelIdx = ga_lngDetailsTabPositions(lngIdx)
        With lstField(lngIdx)
            If lngLabelIdx > -1 Then
                .ListIndex = lngLabelIdx - 29
            Else
                .ListIndex = .ListCount - 1
            End If
        End With
    Next lngIdx

    '================
    'musicplayers tab
    '================

    'sidplayers
    m_lngOldNumPlayers = g_lngTotalSidPlayers
    ReDim m_strOldPlayers(m_lngOldNumPlayers)
    For lngPathIdx = 0 To (g_lngTotalSidPlayers - 1)
        With ga_udtSidPlayers(lngPathIdx)
            ctlSidPlayers.AddItem .strName, .strFullPathAndFilename, .strCommandLine, .blnUseShortFilenames, .a_strTypes, .lngTotalTypes
            m_strOldPlayers(lngPathIdx) = .strName
        End With
    Next lngPathIdx

    '=========
    'games tab
    '=========
    
    'game path device prompt
     chkGamePathPrompt = Abs(g_blnPromptForGamePathDevice)
     
    'rezip games
    chkReZip = Abs(g_blnReZip)
    chkRezipAsk = Abs(g_blnRezipAsk)
    chkRezipNotify = Abs(g_blnZipNotify)
    If ge_RezipAdditional = eRezipAdditionalIgnore Then
        optRezipIgnore.Value = True
    ElseIf ge_RezipAdditional = eRezipAdditionalZip Then
        optRezipZip.Value = True
    Else
        optRezipAsk.Value = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If m_blnOKPressed = False Then
        If g_blnChangedSidplayerSettings = True Then
            If MsgBox(g_xobjFunc.GetINIMessage("Database Options Window", "MusicChangedWarning", "You appear to have changed the properties of at least one Music player.") & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Database Options Window", "CancelAndLoseChanges", "Are you sure you want to cancel and lose these changes?"), vbQuestion Or vbYesNo, gc_strAPPNAME) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If

End Sub


Private Sub lblPicBack_Click()

    On Error GoTo CancelPressed

    With cDialog
        .Color = g_lngPictureBackColor
        .DialogTitle = lblPrompt(2).Caption
        .Flags = cdlCCRGBInit Or cdlCCFullOpen
        .ShowColor
        g_lngPictureBackColor = .Color
        lblPicBack.BackColor = .Color
    End With

    Exit Sub

CancelPressed:
    
End Sub


Private Sub optPictures_Click(Index As Integer)

    If m_blnSetOnce = False Then
        m_blnSetOnce = True
    Else
        Call ShowTip(gc_lngTIPS_PICTURESTYLE)
    End If

End Sub

Private Sub txtNH_GotFocus()
    With txtNH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtNW_GotFocus()
    With txtNW
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNW_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtSH_GotFocus()
    With txtSH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtSW_GotFocus()
    With txtSW
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSW_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtZH_GotFocus()
    With txtZH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtZH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtZW_GotFocus()
    With txtZW
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtZW_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


