VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGameInfoProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Info Properties..."
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "frmGameInfoProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDummy 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   66
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   67
      Top             =   6600
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProps 
      Height          =   6255
      Left            =   120
      TabIndex        =   68
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Include"
      TabPicture(0)   =   "frmGameInfoProps.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraUpdate"
      Tab(0).Control(1)=   "cmdNone"
      Tab(0).Control(2)=   "cmdAll"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fields"
      TabPicture(1)   =   "frmGameInfoProps.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPrompt(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPrompt(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPrompt(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblPrompt(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPrompt(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPrompt(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblPrompt(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblPrompt(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblPrompt(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblPrompt(11)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblPublisher"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblProgrammer"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblMusician"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblLanguage"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblArtist"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblPrompt(8)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblDeveloper"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblPrompt(10)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lblLicense"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lblPrompt(12)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lblRarity"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "lblPrompt(13)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lblPrompt(14)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lblPrompt(15)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lblPrompt(16)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lblReviewRating"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "sldReviewRating"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "lstYear"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmdPublisher"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdProgrammer"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cmdMusician"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lstGenre"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lstControl"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmdLanguage"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtComment"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "optPlayersOnly"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtPlayers"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "optPlayersTo"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txtPlayersTo"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtPlayersFrom"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "chkSim"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "cmdArtist"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "cmdDeveloper"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "cmdLicense"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "cmdRarity"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "txtWebLinkName"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "txtWebLinkURL"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "chkReviewRating"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).ControlCount=   48
      Begin VB.CheckBox chkReviewRating 
         Caption         =   "(Unknown)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtWebLinkURL 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   5040
         Width           =   3855
      End
      Begin VB.TextBox txtWebLinkName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   4320
         Width           =   3855
      End
      Begin VB.CommandButton cmdRarity 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   7800
         TabIndex        =   43
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdLicense 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton cmdDeveloper 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdArtist 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   3600
         Width           =   375
      End
      Begin VB.CheckBox chkSim 
         Caption         =   "&Simultaneous Play"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6480
         TabIndex        =   35
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtPlayersFrom 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   32
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPlayersTo 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   34
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton optPlayersTo 
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtPlayers 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   30
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton optPlayersOnly 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtComment 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   5730
         Width           =   7935
      End
      Begin VB.CommandButton cmdLanguage 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   7800
         TabIndex        =   38
         Top             =   2160
         Width           =   375
      End
      Begin VB.ComboBox lstControl 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2880
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox lstGenre 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdMusician 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdProgrammer 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdPublisher 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox lstYear 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   -74760
         TabIndex        =   64
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   -73200
         TabIndex        =   65
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Fields to Update"
         Height          =   3615
         Left            =   -74760
         TabIndex        =   48
         Top             =   480
         Width           =   7935
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Review Rating"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   56
            Top             =   2880
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Developer"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   51
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "More In&fo"
            Height          =   255
            Index           =   13
            Left            =   4080
            TabIndex        =   63
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Rar&ity"
            Height          =   255
            Index           =   12
            Left            =   4080
            TabIndex        =   62
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "L&icense"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   55
            Top             =   2520
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Artist"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   54
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Con&trol"
            Height          =   255
            Index           =   7
            Left            =   4080
            TabIndex        =   61
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "C&omment"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   57
            Top             =   3240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Language"
            Height          =   255
            Index           =   6
            Left            =   4080
            TabIndex        =   60
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "N&umber of Players"
            Height          =   255
            Index           =   5
            Left            =   4080
            TabIndex        =   59
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Genre"
            Height          =   255
            Index           =   4
            Left            =   4080
            TabIndex        =   58
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Musician"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   53
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "P&rogrammer"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   52
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Publisher"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Year"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin MSComctlLib.Slider sldReviewRating 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5055
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Max             =   100
      End
      Begin VB.Label lblReviewRating 
         Alignment       =   1  'Right Justify
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Revie&w Rating:"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   19
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "More Info &URL:"
         Height          =   255
         Index           =   15
         Left            =   4320
         TabIndex        =   46
         Top             =   4800
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "More Info &Name:"
         Height          =   255
         Index           =   14
         Left            =   4320
         TabIndex        =   44
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Rarit&y:"
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   41
         Top             =   3360
         Width           =   3855
      End
      Begin VB.Label lblRarity 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   42
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "L&icense:"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label lblLicense 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4320
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Developer:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblDeveloper 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artist:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label lblArtist 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblLanguage 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblMusician 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblProgrammer 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblPublisher 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Index           =   11
         Left            =   5760
         TabIndex        =   33
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Comm&ent:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   5520
         Width           =   7935
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Pla&yers:"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   28
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Language:"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   36
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Con&trol:"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   39
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Genre:"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   26
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Musician:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "P&rogrammer:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Publisher:"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Year:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label lblDummy 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   1200
      TabIndex        =   70
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmGameInfoProps"
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

Private m_blnSeenOptions As Boolean
Private m_blnLoading As Boolean
Private m_blnStartup As Boolean
Private m_lngPubID As Long
Private m_lngProID As Long
Private m_lngMusID As Long
Private m_lngLanID As Long
Private m_lngArtID As Long
Private m_lngDevID As Long
Private m_lngLicID As Long
Private m_lngRarID As Long
Private m_blnFirstTime As Boolean
Private Function PlayersOK() As Boolean

    If optPlayersOnly.Value = True Then
        PlayersOK = ((txtPlayers.Text = "?") Or (Val(txtPlayers.Text) > -1))
    Else
        If (txtPlayersFrom.Text = "?" Or Val(txtPlayersFrom.Text) > -1) And (txtPlayersTo.Text = "?" Or Val(txtPlayersTo.Text) > -1) Then
            If txtPlayersFrom.Text = "?" Then
                PlayersOK = True
            ElseIf txtPlayersTo.Text = "?" Then
                PlayersOK = True
            Else
                PlayersOK = Val(txtPlayersTo.Text) >= Val(txtPlayersFrom.Text)
            End If
        Else
            PlayersOK = False
        End If
    End If

End Function

Private Sub SetOKButton()

    If m_blnSeenOptions = False Then
        cmdOK.Enabled = False
    Else
        If (lstYear.ListIndex > -1) And _
           (lstGenre.ListIndex > -1) And _
           (lstControl.ListIndex > -1) And _
           (PlayersOK = True) Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End If

End Sub

Private Sub SetPlayers(ByVal lngFrom_in As Long, ByVal lngTo_in As Long)

    If lngFrom_in = lngTo_in Then
        txtPlayers.Text = IIf((lngFrom_in = gc_lngPLAYERSUNKNOWN), "?", lngFrom_in)
        optPlayersOnly.Value = True
    Else
        txtPlayersFrom.Text = IIf((lngFrom_in = gc_lngPLAYERSUNKNOWN), "?", lngFrom_in)
        txtPlayersTo.Text = IIf((lngTo_in = gc_lngPLAYERSUNKNOWN), "?", lngTo_in)
        optPlayersTo.Value = True
    End If

End Sub

Private Sub chkReviewRating_Click()

    If chkReviewRating.Value = vbChecked Then
        'unknown
        sldReviewRating.Enabled = False
        sldReviewRating.Visible = False
        lblReviewRating.Visible = False
    Else
        'not unknown
        sldReviewRating.Enabled = True
        sldReviewRating.Visible = True
        lblReviewRating.Visible = True
        lblReviewRating.Caption = sldReviewRating.Value & "%"
    End If

End Sub

Private Sub chkSim_Click()

    If chkSim.Value = vbChecked Then
        If optPlayersOnly.Value = True Then
            If txtPlayers = "1" Or txtPlayers = "0" Then
                chkSim.Value = vbUnchecked
            End If
        Else
            If txtPlayersFrom = "1" And txtPlayersTo = "1" Then
                chkSim.Value = vbUnchecked
            End If
        End If
    End If

End Sub

Private Sub chkUpdate_Click(Index As Integer)

    Dim lngIdx As Long
    Dim blnOK As Boolean

    If g_udtEditGameInfo.blnEditAll = True Then
        If Index = 0 Then
            If chkUpdate(0).Value = vbChecked Then
                lstYear.Enabled = True
                lstYear.BackColor = lstDummy.BackColor
            Else
                lstYear.Enabled = False
                lstYear.BackColor = BackColor
            End If
        End If
        If Index = 1 Then
            If chkUpdate(1).Value = vbChecked Then
                lblPublisher.Enabled = True
                cmdPublisher.Enabled = True
                lblPublisher.BackColor = lblDummy.BackColor
            Else
                lblPublisher.Enabled = False
                cmdPublisher.Enabled = False
                lblPublisher.BackColor = BackColor
            End If
        End If
        If Index = 9 Then
            If chkUpdate(9).Value = vbChecked Then
                lblDeveloper.Enabled = True
                cmdDeveloper.Enabled = True
                lblDeveloper.BackColor = lblDummy.BackColor
            Else
                lblDeveloper.Enabled = False
                cmdDeveloper.Enabled = False
                lblDeveloper.BackColor = BackColor
            End If
        End If
        If Index = 10 Then
            If chkUpdate(10).Value = vbChecked Then
                lblArtist.Enabled = True
                cmdArtist.Enabled = True
                lblArtist.BackColor = lblDummy.BackColor
            Else
                lblArtist.Enabled = False
                cmdArtist.Enabled = False
                lblArtist.BackColor = BackColor
            End If
        End If
        If Index = 11 Then
            If chkUpdate(11).Value = vbChecked Then
                lblLicense.Enabled = True
                cmdLicense.Enabled = True
                lblLicense.BackColor = lblDummy.BackColor
            Else
                lblLicense.Enabled = False
                cmdLicense.Enabled = False
                lblLicense.BackColor = BackColor
            End If
        End If
        If Index = 12 Then
            If chkUpdate(12).Value = vbChecked Then
                lblRarity.Enabled = True
                cmdRarity.Enabled = True
                lblRarity.BackColor = lblDummy.BackColor
            Else
                lblRarity.Enabled = False
                cmdRarity.Enabled = False
                lblRarity.BackColor = BackColor
            End If
        End If
        If Index = 2 Then
            If chkUpdate(2).Value = vbChecked Then
                lblProgrammer.Enabled = True
                cmdProgrammer.Enabled = True
                lblProgrammer.BackColor = lblDummy.BackColor
            Else
                lblProgrammer.Enabled = False
                cmdProgrammer.Enabled = False
                lblProgrammer.BackColor = BackColor
            End If
        End If
        If Index = 3 Then
            If chkUpdate(3).Value = vbChecked Then
                lblMusician.Enabled = True
                cmdMusician.Enabled = True
                lblMusician.BackColor = lblDummy.BackColor
            Else
                lblMusician.Enabled = False
                cmdMusician.Enabled = False
                lblMusician.BackColor = BackColor
            End If
        End If
        If Index = 4 Then
            If chkUpdate(4).Value = vbChecked Then
                lstGenre.Enabled = True
                lstGenre.BackColor = lstDummy.BackColor
            Else
                lstGenre.Enabled = False
                lstGenre.BackColor = BackColor
            End If
        End If
        If Index = 5 Then
            If chkUpdate(5).Value = vbChecked Then
                optPlayersOnly.Enabled = True
                optPlayersTo.Enabled = True
                chkSim.Enabled = True
                If optPlayersOnly.Value = True Then
                    txtPlayers.Enabled = True
                    txtPlayersFrom.Enabled = False
                    txtPlayersTo.Enabled = False
                    txtPlayers.BackColor = lstDummy.BackColor
                    txtPlayersFrom.BackColor = BackColor
                    txtPlayersTo.BackColor = BackColor
                Else
                    txtPlayers.Enabled = False
                    txtPlayersFrom.Enabled = True
                    txtPlayersTo.Enabled = True
                    txtPlayers.BackColor = BackColor
                    txtPlayersFrom.BackColor = lstDummy.BackColor
                    txtPlayersTo.BackColor = lstDummy.BackColor
                End If
            Else
                optPlayersOnly.Enabled = False
                optPlayersTo.Enabled = False
                txtPlayers.Enabled = False
                txtPlayersFrom.Enabled = False
                txtPlayersTo.Enabled = False
                txtPlayers.BackColor = BackColor
                txtPlayersFrom.BackColor = BackColor
                txtPlayersTo.BackColor = BackColor
                chkSim.Enabled = False
            End If
        End If
        If Index = 6 Then
            If chkUpdate(6).Value = vbChecked Then
                lblLanguage.Enabled = True
                cmdLanguage.Enabled = True
                lblLanguage.BackColor = lblDummy.BackColor
            Else
                lblLanguage.Enabled = False
                cmdLanguage.Enabled = False
                lblLanguage.BackColor = BackColor
            End If
        End If
        If Index = 7 Then
            If chkUpdate(7).Value = vbChecked Then
                lstControl.Enabled = True
                lstControl.BackColor = lstDummy.BackColor
            Else
                lstControl.Enabled = False
                lstControl.BackColor = BackColor
            End If
        End If
        If Index = 8 Then
            If chkUpdate(8).Value = vbChecked Then
                txtComment.Enabled = True
                txtComment.BackColor = lstDummy.BackColor
            Else
                txtComment.Enabled = False
                txtComment.BackColor = BackColor
            End If
        End If
        If Index = 13 Then
            If chkUpdate(13).Value = vbChecked Then
                txtWebLinkName.Enabled = True
                txtWebLinkName.BackColor = lstDummy.BackColor
                txtWebLinkURL.Enabled = True
                txtWebLinkURL.BackColor = lstDummy.BackColor
            Else
                txtWebLinkName.Enabled = False
                txtWebLinkName.BackColor = BackColor
                txtWebLinkURL.Enabled = False
                txtWebLinkURL.BackColor = BackColor
            End If
        End If
        If Index = 14 Then
            If chkUpdate(14).Value = vbChecked Then
                chkReviewRating.Enabled = True
                sldReviewRating.Enabled = (chkReviewRating.Value = vbUnchecked)
                sldReviewRating.Visible = sldReviewRating.Enabled
            Else
                chkReviewRating.Enabled = False
                sldReviewRating.Enabled = False
                sldReviewRating.Visible = False
            End If
        End If
        
        'any selected?
        blnOK = False
        For lngIdx = 0 To 14
            If chkUpdate(lngIdx).Value = vbChecked Then
                blnOK = True
                Exit For
            End If
        Next lngIdx
        If Not blnOK Then
            cmdOK.Enabled = False
            tabProps.TabEnabled(1) = False
        Else
            tabProps.TabEnabled(1) = True
            Call SetOKButton
        End If
    End If

End Sub

Private Sub cmdAll_Click()

    Dim lngIdx As Long
    
    For lngIdx = 0 To 14
        chkUpdate(lngIdx).Value = vbChecked
    Next lngIdx

End Sub

Private Sub cmdArtist_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Artist
        .lngID = m_lngArtID
        .strSQL = "SELECT Artist AS Name, AR_Id as ID FROM Artists ORDER BY Artist"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngArtID = .lngID
            lblArtist.Caption = .strReturnedName
        End If
    End With
    cmdArtist.SetFocus

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDeveloper_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Developer
        .lngID = m_lngDevID
        .strSQL = "SELECT Developer AS Name, DE_Id as ID FROM Developers ORDER BY Developer"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngDevID = .lngID
            lblDeveloper.Caption = .strReturnedName
        End If
    End With
    cmdDeveloper.SetFocus

End Sub

Private Sub cmdLanguage_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Language
        .lngID = m_lngLanID
        .strSQL = "SELECT Language AS Name, LA_Id as ID FROM Languages ORDER BY Language"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngLanID = .lngID
            lblLanguage.Caption = .strReturnedName
        End If
    End With
    cmdLanguage.SetFocus

End Sub

Private Sub cmdLicense_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_License
        .lngID = m_lngLicID
        .strSQL = "SELECT License AS Name, LI_Id as ID FROM Licenses ORDER BY License"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngLicID = .lngID
            lblLicense.Caption = .strReturnedName
        End If
    End With
    cmdLicense.SetFocus

End Sub

Private Sub cmdMusician_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Musician
        .lngID = m_lngMusID
        .strSQL = "SELECT Musician AS Name, MU_Id as ID FROM Musicians ORDER BY Musician"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngMusID = .lngID
            lblMusician.Caption = .strReturnedName
        End If
    End With
    cmdMusician.SetFocus

End Sub


Private Sub cmdNone_Click()

    Dim lngIdx As Long
    
    For lngIdx = 0 To 14
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx
    
End Sub

Private Sub cmdOK_Click()

    'confirmation
    If g_udtEditGameInfo.blnEditAll Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSure", "Are you sure you want to update ALL GAMES in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If
    'save all info
    With g_udtEditGameInfo
        .blnOKClicked = True
        If .blnEditAll = True Then
            .blnUpdateComment = (chkUpdate(8).Value = vbChecked)
            .blnUpdateControl = (chkUpdate(7).Value = vbChecked)
            .blnUpdateGenre = (chkUpdate(4).Value = vbChecked)
            .blnUpdateLanguage = (chkUpdate(6).Value = vbChecked)
            .blnUpdateMusician = (chkUpdate(3).Value = vbChecked)
            .blnUpdatePlayers = (chkUpdate(5).Value = vbChecked)
            .blnUpdateProgrammer = (chkUpdate(2).Value = vbChecked)
            .blnUpdatePublisher = (chkUpdate(1).Value = vbChecked)
            .blnUpdateYear = (chkUpdate(0).Value = vbChecked)
            .blnUpdateArtist = (chkUpdate(10).Value = vbChecked)
            .blnUpdateDeveloper = (chkUpdate(9).Value = vbChecked)
            .blnUpdateLicense = (chkUpdate(11).Value = vbChecked)
            .blnUpdateRarity = (chkUpdate(12).Value = vbChecked)
            .blnUpdateWebLink = (chkUpdate(13).Value = vbChecked)
            .blnUpdateReviewRating = (chkUpdate(14).Value = vbChecked)
        End If
        .lngControlID = lstControl.ItemData(lstControl.ListIndex)
        .lngGenreID = lstGenre.ItemData(lstGenre.ListIndex)
        .lngLanguageID = m_lngLanID
        .lngMusicianID = m_lngMusID
        .lngArtistID = m_lngArtID
        .lngDeveloperID = m_lngDevID
        .lngLicenseID = m_lngLicID
        .lngRarityID = m_lngRarID
        If optPlayersOnly.Value = True Then
            .lngPlayersFrom = IIf((txtPlayers.Text = "?"), gc_lngPLAYERSUNKNOWN, Val(txtPlayers.Text))
            .lngPlayersTo = .lngPlayersFrom
            If chkSim.Value = vbChecked Then
                .blnSimPlay = (.lngPlayersFrom < 0 Or .lngPlayersFrom > 1)
            Else
                .blnSimPlay = False
            End If
        Else
            .lngPlayersFrom = IIf((txtPlayersFrom.Text = "?"), gc_lngPLAYERSUNKNOWN, Val(txtPlayersFrom.Text))
            .lngPlayersTo = IIf((txtPlayersTo.Text = "?"), gc_lngPLAYERSUNKNOWN, Val(txtPlayersTo.Text))
            .blnSimPlay = (chkSim.Value = vbChecked)
        End If
        .lngProgrammerID = m_lngProID
        .lngPublisherID = m_lngPubID
        .lngYearID = lstYear.ItemData(lstYear.ListIndex)
        .strComment = txtComment.Text
        .strWebLinkName = txtWebLinkName.Text
        .strWebLinkURL = txtWebLinkURL.Text
        If chkReviewRating.Value = vbChecked Then
            .lngReviewRating = -1
        Else
            .lngReviewRating = sldReviewRating.Value
        End If
    End With

    Unload Me

End Sub

Private Sub cmdProgrammer_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Programmer
        .lngID = m_lngProID
        .strSQL = "SELECT Programmer AS Name, PR_Id as ID FROM Programmers ORDER BY Programmer"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngProID = .lngID
            lblProgrammer.Caption = .strReturnedName
        End If
    End With
    cmdProgrammer.SetFocus

End Sub

Private Sub cmdPublisher_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Publisher
        .lngID = m_lngPubID
        .strSQL = "SELECT Publisher AS Name, PU_Id as ID FROM Publishers ORDER BY Publisher"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngPubID = .lngID
            lblPublisher.Caption = .strReturnedName
        End If
    End With
    cmdPublisher.SetFocus

End Sub


Private Sub cmdRarity_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Rarity
        .lngID = m_lngRarID
        .strSQL = "SELECT Rarity AS Name, RA_Id as ID FROM Rarities ORDER BY Rarity"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditGameInfo.blnTableEdited = True
            m_lngRarID = .lngID
            lblRarity.Caption = .strReturnedName
        End If
    End With
    cmdRarity.SetFocus

End Sub

Private Sub Form_Activate()

    'each time the form is shown
    If m_blnFirstTime = True Then
        m_blnFirstTime = False
    Else
        Exit Sub
    End If

    Dim lngIdx As Long

    m_blnStartup = True
    m_blnLoading = False
    
    'clear all checkmarks
    For lngIdx = 0 To 14
        chkUpdate(lngIdx).Value = vbChecked
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx

    'are we changing all the games?
    If g_udtEditGameInfo.blnEditAll = False Then
        lstYear.BackColor = lstDummy.BackColor
        lstYear.Enabled = True
    End If
    If g_udtEditGameInfo.blnEditAll = False Then
        lstGenre.BackColor = lstDummy.BackColor
        lstGenre.Enabled = True
    End If
    If g_udtEditGameInfo.blnEditAll = False Then
        lstControl.BackColor = lstDummy.BackColor
        lstControl.Enabled = True
    End If
    If g_udtEditGameInfo.blnEditAll = False Then
        chkReviewRating.Enabled = True
        sldReviewRating.Enabled = (chkReviewRating.Value = vbUnchecked)
        sldReviewRating.Visible = sldReviewRating.Enabled
    End If
    If g_udtEditGameInfo.blnEditAll = False Then
        optPlayersOnly.Enabled = True
        optPlayersTo.Enabled = True
        cmdPublisher.Enabled = True
        cmdArtist.Enabled = True
        cmdDeveloper.Enabled = True
        cmdLicense.Enabled = True
        cmdRarity.Enabled = True
        cmdMusician.Enabled = True
        cmdProgrammer.Enabled = True
        cmdLanguage.Enabled = True
        lblPublisher.Enabled = True
        lblArtist.Enabled = True
        lblDeveloper.Enabled = True
        lblLicense.Enabled = True
        lblRarity.Enabled = True
        lblMusician.Enabled = True
        lblProgrammer.Enabled = True
        lblLanguage.Enabled = True
        txtComment.BackColor = lstDummy.BackColor
        txtComment.Enabled = True
        txtWebLinkName.BackColor = lstDummy.BackColor
        txtWebLinkName.Enabled = True
        txtWebLinkURL.BackColor = lstDummy.BackColor
        txtWebLinkURL.Enabled = True
        chkSim.Enabled = True
        lblPublisher.BackColor = lblDummy.BackColor
        lblProgrammer.BackColor = lblDummy.BackColor
        lblMusician.BackColor = lblDummy.BackColor
        lblLanguage.BackColor = lblDummy.BackColor
        lblArtist.BackColor = lblDummy.BackColor
        lblDeveloper.BackColor = lblDummy.BackColor
        lblLicense.BackColor = lblDummy.BackColor
        lblRarity.BackColor = lblDummy.BackColor
    End If

    If g_udtEditGameInfo.blnEditAll = True Then
        tabProps.TabEnabled(0) = True
        tabProps.TabEnabled(1) = False
        tabProps.Tab = 0
    Else
        tabProps.TabEnabled(0) = False
        tabProps.TabEnabled(1) = True
        tabProps.Tab = 1
    End If
    
    'init props
    lstYear.ListIndex = 0
    lstControl.ListIndex = 0
    lstGenre.ListIndex = 0
    chkReviewRating.Value = vbChecked
    sldReviewRating.Enabled = (chkReviewRating.Value = vbUnchecked)
    sldReviewRating.Visible = sldReviewRating.Enabled
    txtPlayers.Text = "?"
    optPlayersOnly.Value = True
    chkSim.Value = vbUnchecked
    txtComment.Text = ""
    txtWebLinkName.Text = ""
    txtWebLinkURL.Text = ""

    'fill props
    With lstYear
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditGameInfo.lngYearID Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    With lstGenre
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditGameInfo.lngGenreID Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    With lstControl
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditGameInfo.lngControlID Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    
    If g_udtEditGameInfo.lngReviewRating = -1 Then
        chkReviewRating.Value = vbChecked
        sldReviewRating.Enabled = False
        sldReviewRating.Visible = False
        sldReviewRating.Value = 0
        lblReviewRating.Visible = False
    Else
        chkReviewRating.Value = vbUnchecked
        sldReviewRating.Enabled = True
        sldReviewRating.Visible = True
        sldReviewRating.Value = g_udtEditGameInfo.lngReviewRating
        lblReviewRating.Visible = True
        lblReviewRating.Caption = sldReviewRating.Value & "%"
    End If

    Call SetPlayers(g_udtEditGameInfo.lngPlayersFrom, g_udtEditGameInfo.lngPlayersTo)
    chkSim.Value = Abs(g_udtEditGameInfo.blnSimPlay)
    txtComment.Text = g_udtEditGameInfo.strComment
    txtWebLinkName.Text = g_udtEditGameInfo.strWebLinkName
    txtWebLinkURL.Text = g_udtEditGameInfo.strWebLinkURL
    If tabProps.Tab = 1 Then
        If g_udtEditGameInfoFocus = e_InfoFocus_Comment Then
            txtComment.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Control Then
            lstControl.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Genre Then
            lstGenre.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Language Then
            cmdLanguage.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Musician Then
            cmdMusician.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_infoFocus_Players Then
            If optPlayersOnly.Value = True Then
                txtPlayers.SetFocus
            Else
                txtPlayersFrom.SetFocus
            End If
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Programmer Then
            cmdProgrammer.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Publisher Then
            cmdPublisher.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Artist Then
            cmdArtist.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Developer Then
            cmdDeveloper.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_License Then
            cmdLicense.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_Rarity Then
            cmdRarity.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_WebLink Then
            txtWebLinkName.SetFocus
        ElseIf g_udtEditGameInfoFocus = e_InfoFocus_ReviewRating Then
            If sldReviewRating.Enabled = True And sldReviewRating.Visible = True Then
                sldReviewRating.SetFocus
            Else
                chkReviewRating.SetFocus
            End If
        Else 'year
            lstYear.SetFocus
        End If
    Else
        chkUpdate(0).SetFocus
    End If

    'set ok button and finish off
    m_blnSeenOptions = Not g_udtEditGameInfo.blnEditAll
    Call SetOKButton
    g_udtEditGameInfo.blnOKClicked = False
    g_udtEditGameInfo.blnTableEdited = True
    m_blnStartup = False
    Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 7095    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 8655     'The required scalewidth for the form
    Dim lngIdx As Long
    Dim fldName As Field
    Dim fldID As Field
    Dim fldName2 As Field
    Dim strSQL(7) As String
    Dim lblBox(7) As Label

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    m_blnLoading = True
    With g_udtEditGameInfo
        m_lngPubID = .lngPublisherID
        m_lngProID = .lngProgrammerID
        m_lngMusID = .lngMusicianID
        m_lngLanID = .lngLanguageID
        m_lngArtID = .lngArtistID
        m_lngDevID = .lngDeveloperID
        m_lngLicID = .lngLicenseID
        m_lngRarID = .lngRarityID
    End With
    
    'read in the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Game Info Properties Window", "Caption", "Game Info Properties") & "..."
        tabProps.TabCaption(0) = .GetINIMessage("Game Info Properties Window", "IncludeTab", "&Include")
        tabProps.TabCaption(1) = .GetINIMessage("Game Info Properties Window", "FieldsTab", "&Fields")
        cmdAll.Caption = .GetINIMessage("Game Info Properties Window", "SelectAllButton", "Select &All")
        cmdNone.Caption = .GetINIMessage("Game Info Properties Window", "SelectNoneButton", "Select &None")
        fraUpdate.Caption = .GetINIMessage("Game Info Properties Window", "FieldsToUpdate", "Fields to Update")
        chkUpdate(0).Caption = GetINIMessage("Game Info Properties Window", "UpdateYearBox", "Year")
        chkUpdate(1).Caption = GetINIMessage("Game Info Properties Window", "UpdatePublisherBox", "Publisher")
        chkUpdate(2).Caption = GetINIMessage("Game Info Properties Window", "UpdateProgrammerBox", "Programmer")
        chkUpdate(3).Caption = GetINIMessage("Game Info Properties Window", "UpdateMusicianBox", "Musician")
        chkUpdate(6).Caption = GetINIMessage("Game Info Properties Window", "UpdateLanguageBox", "Language")
        chkUpdate(10).Caption = GetINIMessage("Game Info Properties Window", "UpdateArtistBox", "Artist")
        chkUpdate(9).Caption = GetINIMessage("Game Info Properties Window", "UpdateDeveloperBox", "Developer")
        chkUpdate(11).Caption = GetINIMessage("Game Info Properties Window", "UpdateLicenseBox", "License")
        chkUpdate(12).Caption = GetINIMessage("Game Info Properties Window", "UpdateRarityBox", "Rarity")
        chkUpdate(4).Caption = GetINIMessage("Game Info Properties Window", "UpdateGenreBox", "Genre")
        chkUpdate(5).Caption = GetINIMessage("Game Info Properties Window", "UpdatePlayersBox", "Number of Players")
        chkUpdate(7).Caption = GetINIMessage("Game Info Properties Window", "UpdateControlBox", "Control")
        chkUpdate(8).Caption = GetINIMessage("Game Info Properties Window", "UpdateCommentBox", "Comment")
        chkUpdate(13).Caption = GetINIMessage("Game Info Properties Window", "UpdateWebLinkBox", "More Info")
        chkUpdate(14).Caption = GetINIMessage("Game Info Properties Window", "UpdateReviewRatingBox", "Review Rating")
        lblPrompt(0).Caption = GetINIMessage("Game Info Properties Window", "YearPrompt", "&Year") & ":"
        lblPrompt(1).Caption = GetINIMessage("Game Info Properties Window", "PublisherPrompt", "&Publisher") & ":"
        lblPrompt(2).Caption = GetINIMessage("Game Info Properties Window", "ProgrammerPrompt", "P&rogrammer") & ":"
        lblPrompt(3).Caption = GetINIMessage("Game Info Properties Window", "MusicianPrompt", "&Musician") & ":"
        lblPrompt(6).Caption = GetINIMessage("Game Info Properties Window", "LanguagePrompt", "&Language") & ":"

        lblPrompt(8).Caption = GetINIMessage("Game Info Properties Window", "ArtistPrompt", "&Artist") & ":"
        lblPrompt(10).Caption = GetINIMessage("Game Info Properties Window", "DeveloperPrompt", "&Developer") & ":"
        lblPrompt(12).Caption = GetINIMessage("Game Info Properties Window", "LicensePrompt", "L&icense") & ":"
        lblPrompt(13).Caption = GetINIMessage("Game Info Properties Window", "RarityPrompt", "&Rarity") & ":"
        lblPrompt(14).Caption = GetINIMessage("Game Info Properties Window", "WebLinkNamePrompt", "More Info &Name") & ":"
        lblPrompt(15).Caption = GetINIMessage("Game Info Properties Window", "WebLinkURLPrompt", "More Info &URL") & ":"
        lblPrompt(16).Caption = GetINIMessage("Game Info Properties Window", "ReviewRatingPrompt", "Revie&w Rating") & ":"

        lblPrompt(4).Caption = GetINIMessage("Game Info Properties Window", "GenrePrompt", "&Genre") & ":"
        lblPrompt(5).Caption = GetINIMessage("Game Info Properties Window", "ControlPrompt", "Con&trol") & ":"
        lblPrompt(7).Caption = GetINIMessage("Game Info Properties Window", "PlayersPrompt", "No. of Pla&yers") & ":"
        lblPrompt(9).Caption = GetINIMessage("Game Info Properties Window", "CommentPrompt", "Comm&ent") & ":"
        chkSim.Caption = GetINIMessage("Game Info Properties Window", "SimultaneousPlayPrompt", "&Simultaneous Play")
        chkReviewRating.Caption = "(" & g_strUnknown & ")"
    End With
    
    'read in the years
    With lstYear
        For lngIdx = 1 To gc_lngYEARLOOKUPMAX
            .AddItem ga_strYearLookupTable(lngIdx)
            .ItemData(.NewIndex) = lngIdx
        Next lngIdx
        .Visible = True
    End With

    'read in the genres
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Genre, ParentGenre, GE_ID FROM PGenres INNER JOIN Genres ON PGenres.PG_Id = Genres.PG_Id ORDER BY PGenres.ParentGenre, Genres.Genre", dbOpenSnapshot)
    Set fldName = g_rsGames!Genre
    Set fldName2 = g_rsGames!ParentGenre
    Set fldID = g_rsGames!GE_Id
    With lstGenre
        g_rsGames.MoveFirst
        Do Until g_rsGames.EOF
            If Len(fldName2) > 0 Then
                .AddItem fldName2 & " - " & fldName
            Else
                .AddItem fldName
            End If
            .ItemData(.NewIndex) = fldID
            g_rsGames.MoveNext
        Loop
        g_rsGames.Close
    End With
    Set fldName = Nothing
    Set fldName2 = Nothing
    Set fldID = Nothing
    Set g_rsGames = Nothing
    lstGenre.Visible = True

    'read in the controls
    With lstControl
        For lngIdx = 0 To gc_lngCONTROLLOOKUPMAX
            .AddItem ga_strControlLookupTable(lngIdx)
            .ItemData(.NewIndex) = lngIdx
        Next lngIdx
        .Visible = True
    End With

    'read in the table data
    strSQL(0) = "SELECT PU_Id as ID, Publisher as Name FROM Publishers WHERE PU_Id = " & m_lngPubID
    strSQL(1) = "SELECT PR_Id as ID, Programmer as Name FROM Programmers WHERE PR_Id = " & m_lngProID
    strSQL(2) = "SELECT MU_Id as ID, Musician as Name FROM Musicians WHERE MU_Id = " & m_lngMusID
    strSQL(3) = "SELECT LA_Id as ID, Language as Name FROM Languages WHERE LA_Id = " & m_lngLanID
    strSQL(4) = "SELECT AR_Id as ID, Artist as Name FROM Artists WHERE AR_Id = " & m_lngArtID
    strSQL(5) = "SELECT DE_Id as ID, Developer as Name FROM Developers WHERE DE_Id = " & m_lngDevID
    strSQL(6) = "SELECT LI_Id as ID, License as Name FROM Licenses WHERE LI_Id = " & m_lngLicID
    strSQL(7) = "SELECT RA_Id as ID, Rarity as Name FROM Rarities WHERE RA_Id = " & m_lngRarID
    Set lblBox(0) = lblPublisher
    Set lblBox(1) = lblProgrammer
    Set lblBox(2) = lblMusician
    Set lblBox(3) = lblLanguage
    Set lblBox(4) = lblArtist
    Set lblBox(5) = lblDeveloper
    Set lblBox(6) = lblLicense
    Set lblBox(7) = lblRarity
    For lngIdx = 0 To 7
        Set g_rsGames = g_dbGamebase.OpenRecordset(strSQL(lngIdx), dbOpenSnapshot)
        With g_rsGames
            If Not .EOF Then
                lblBox(lngIdx).Caption = !Name
            End If
            .Close
        End With
        Set g_rsGames = Nothing
    Next lngIdx
    Set lblBox(0) = Nothing
    Set lblBox(1) = Nothing
    Set lblBox(2) = Nothing
    Set lblBox(3) = Nothing
    Set lblBox(4) = Nothing
    Set lblBox(5) = Nothing
    Set lblBox(6) = Nothing
    Set lblBox(7) = Nothing

    m_blnFirstTime = True

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If

End Sub

Private Sub lstControl_Click()

    Call SetOKButton

End Sub


Private Sub lstGenre_Click()

    Call SetOKButton

End Sub




Private Sub lstYear_Click()

    Call SetOKButton

End Sub


Private Sub optPlayersOnly_Click()

    If Not ((g_udtEditGameInfo.blnEditAll = True) And (m_blnStartup = True)) Then
        txtPlayers.Enabled = True
        txtPlayers.BackColor = lstDummy.BackColor
        txtPlayersFrom.Enabled = False
        txtPlayersTo.Enabled = False
        txtPlayersFrom.BackColor = BackColor
        txtPlayersTo.BackColor = BackColor
        txtPlayers.SetFocus
        Call SetOKButton
    End If

End Sub


Private Sub optPlayersTo_Click()

    If Not ((g_udtEditGameInfo.blnEditAll = True) And (m_blnStartup = True)) Then
        txtPlayersTo.Enabled = True
        txtPlayersFrom.Enabled = True
        txtPlayersFrom.BackColor = lstDummy.BackColor
        txtPlayersTo.BackColor = lstDummy.BackColor
        txtPlayers.Enabled = False
        txtPlayers.BackColor = BackColor
        txtPlayersFrom.SetFocus
        Call SetOKButton
    End If

End Sub


Private Sub sldReviewRating_Change()

    lblReviewRating.Caption = sldReviewRating.Value & "%"

End Sub

Private Sub sldReviewRating_Scroll()

    lblReviewRating.Caption = sldReviewRating.Value & "%"

End Sub

Private Sub tabProps_Click(PreviousTab As Integer)
    
    If tabProps.Tab <> PreviousTab Then
        If tabProps.Tab = 1 Then
            If Not m_blnLoading Then
                If lstYear.Enabled = True Then
                    lstYear.SetFocus
                ElseIf cmdPublisher.Enabled = True Then
                    cmdPublisher.SetFocus
                ElseIf cmdDeveloper.Enabled = True Then
                    cmdDeveloper.SetFocus
                ElseIf cmdProgrammer.Enabled = True Then
                    cmdProgrammer.SetFocus
                ElseIf cmdMusician.Enabled = True Then
                    cmdMusician.SetFocus
                ElseIf cmdArtist.Enabled = True Then
                    cmdArtist.SetFocus
                ElseIf cmdLicense.Enabled = True Then
                    cmdLicense.SetFocus
                ElseIf chkReviewRating.Enabled = True Then
                    chkReviewRating.SetFocus
                ElseIf sldReviewRating.Enabled = True And sldReviewRating.Visible = True Then
                    sldReviewRating.SetFocus
                ElseIf txtComment.Enabled = True Then
                    txtComment.SetFocus
                ElseIf lstGenre.Enabled = True Then
                    lstGenre.SetFocus
                ElseIf txtPlayers.Enabled = True Then
                    txtPlayers.SetFocus
                ElseIf txtPlayersFrom.Enabled = True Then
                    txtPlayersFrom.SetFocus
                ElseIf txtPlayersTo.Enabled = True Then
                    txtPlayersTo.SetFocus
                ElseIf chkSim.Enabled = True Then
                    chkSim.SetFocus
                ElseIf cmdLanguage.Enabled = True Then
                    cmdLanguage.SetFocus
                ElseIf lstControl.Enabled = True Then
                    lstControl.SetFocus
                ElseIf cmdRarity.Enabled = True Then
                    cmdRarity.SetFocus
                ElseIf txtWebLinkName.Enabled = True Then
                    txtWebLinkName.SetFocus
                ElseIf txtWebLinkURL.Enabled = True Then
                    txtWebLinkURL.SetFocus
                End If
            End If
        Else
            If chkUpdate(0).Enabled = True And (Not m_blnLoading) Then
                chkUpdate(0).SetFocus
            End If
        End If
    End If
    If tabProps.Tab = 1 Then
        m_blnSeenOptions = True
        Call SetOKButton
    End If
    
End Sub

Private Sub txtComment_GotFocus()

    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtPlayers_Change()
    Call SetOKButton
End Sub


Private Sub txtPlayers_GotFocus()
    With txtPlayers
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtPlayersFrom_Change()
    Call SetOKButton
End Sub


Private Sub txtPlayersFrom_GotFocus()
    With txtPlayersFrom
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtPlayersTo_Change()
    Call SetOKButton
End Sub


Private Sub txtPlayersTo_GotFocus()
    With txtPlayersTo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWebLinkName_GotFocus()

    With txtWebLinkName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtWebLinkURL_GotFocus()

    With txtWebLinkURL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
