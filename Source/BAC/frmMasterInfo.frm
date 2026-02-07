VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMasterInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER GAME INFO"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmMasterInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   4665
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "&Game"
         TabPicture(0)   =   "frmMasterInfo.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label2(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label2(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label2(7)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label2(8)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label2(5)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label2(13)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label2(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label2(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label2(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label2(12)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label2(19)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "&Version"
         TabPicture(1)   =   "frmMasterInfo.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(28)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(18)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label2(17)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label2(16)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label2(15)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label2(14)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label2(27)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label2(26)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label2(25)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label2(24)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label2(23)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label2(22)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label2(21)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label2(20)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label2(6)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "&Notes"
         TabPicture(2)   =   "frmMasterInfo.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Text1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   3495
            Left            =   -74880
            Locked          =   -1  'True
            MaxLength       =   65526
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "MusicFilename:"
            Height          =   255
            Index           =   28
            Left            =   -74880
            TabIndex        =   34
            Top             =   3780
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CRC-32:"
            Height          =   255
            Index           =   18
            Left            =   -74880
            TabIndex        =   33
            Top             =   2820
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "PictureFilename:"
            Height          =   255
            Index           =   17
            Left            =   -74880
            TabIndex        =   32
            Top             =   3540
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "FilenameIndex:"
            Height          =   255
            Index           =   16
            Left            =   -74880
            TabIndex        =   31
            Top             =   3300
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "FileToRun:"
            Height          =   255
            Index           =   15
            Left            =   -74880
            TabIndex        =   30
            Top             =   3060
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
            Height          =   255
            Index           =   14
            Left            =   -74880
            TabIndex        =   29
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "VersionComment:"
            Height          =   255
            Index           =   27
            Left            =   -74880
            TabIndex        =   28
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Game Length:"
            Height          =   255
            Index           =   26
            Left            =   -74880
            TabIndex        =   27
            Top             =   2100
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "True Drive Emulation:"
            Height          =   255
            Index           =   25
            Left            =   -74880
            TabIndex        =   26
            Top             =   1860
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "PAL/NTSC:"
            Height          =   255
            Index           =   24
            Left            =   -74880
            TabIndex        =   25
            Top             =   1620
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Included Docs:"
            Height          =   255
            Index           =   23
            Left            =   -74880
            TabIndex        =   24
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "High Score Saver:"
            Height          =   255
            Index           =   22
            Left            =   -74880
            TabIndex        =   23
            Top             =   1140
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Loading Screen:"
            Height          =   255
            Index           =   21
            Left            =   -74880
            TabIndex        =   22
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Adult:"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   20
            Top             =   3660
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Num Trainers:"
            Height          =   255
            Index           =   20
            Left            =   -74880
            TabIndex        =   19
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Related:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   18
            Top             =   3420
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sequel:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   17
            Top             =   3180
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Prequel:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   16
            Top             =   2940
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cracker:"
            Height          =   255
            Index           =   6
            Left            =   -74880
            TabIndex        =   15
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Control:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   14
            Top             =   2460
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Comment:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   13
            Top             =   2700
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Language:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   2220
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Players:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   1980
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Genre:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   10
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Musician:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Programmer:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   1020
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Publisher:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   540
            Width           =   375
         End
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME GOES HERE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4560
      End
   End
   Begin VB.PictureBox picNoInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NO GAME SELECTED OR CHOSEN DIFF GAME ID DOES NOT EXIST IN THE MASTER DATABASE!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   720
         TabIndex        =   1
         Top             =   1560
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmMasterInfo"
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
'       Borg Assimilation Console (BAC)
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

Public Sub ShowGame(ByVal lngID_in As Long)

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    picInfo.Enabled = True
    picInfo.ZOrder 0

End Sub

Public Sub ShowNoGame()

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    picInfo.Enabled = False
    picNoInfo.ZOrder 0

End Sub


