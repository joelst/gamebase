VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emulator Setup Wizard"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   2280
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilstWizard 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":08CA
            Key             =   "info"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "&Abort"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   14
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Top             =   600
         Width           =   6495
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Left            =   4800
            TabIndex        =   4
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   1320
            MaxLength       =   255
            TabIndex        =   3
            Top             =   1440
            Width           =   3495
         End
         Begin MSComctlLib.ProgressBar pbSearch 
            Height          =   255
            Left            =   1320
            TabIndex        =   32
            Top             =   1440
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblSearchFolder 
            Height          =   255
            Left            =   1320
            TabIndex        =   43
            Top             =   1725
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblSearch 
            Caption         =   "Searching..."
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Top             =   1200
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Search &Path:"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   2
            Top             =   1200
            Width           =   3855
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Search for Emulators"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5205
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   1
         Left            =   0
         Picture         =   "frmWizard.frx":0978
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   15
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   3
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   6495
         Begin VB.CommandButton cmdWebSite 
            Caption         =   "&Web Site"
            Height          =   375
            Left            =   4800
            TabIndex        =   10
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelectNone 
            Caption         =   "Select N&one"
            Height          =   375
            Left            =   1800
            TabIndex        =   8
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Selec&t All"
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelectFile 
            Caption         =   "Select F&ile"
            Height          =   375
            Left            =   3360
            TabIndex        =   9
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   615
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   65526
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   2760
            Width           =   5775
         End
         Begin MSComctlLib.ListView lvSelect 
            Height          =   1455
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ilstWizard"
            SmallIcons      =   "ilstWizard"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   4789
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Filename"
               Object.Width           =   4789
            EndProperty
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Select the &Emulators you wish to setup:"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   5775
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Emulators"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1200
         TabIndex        =   18
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5205
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   2
         Left            =   0
         Picture         =   "frmWizard.frx":0AFF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   3
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   16
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   4
         Left            =   0
         TabIndex        =   35
         Top             =   600
         Width           =   6495
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":0C0D
            Height          =   975
            Index           =   4
            Left            =   1320
            TabIndex        =   36
            Top             =   1320
            UseMnemonic     =   0   'False
            Width           =   3855
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Emulators to Setup!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1200
         TabIndex        =   19
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5295
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   3
         Left            =   0
         Picture         =   "frmWizard.frx":0C98
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   5
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   20
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   5
         Left            =   0
         TabIndex        =   37
         Top             =   600
         Width           =   6495
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Did the Wizard not have the emulator you wanted?"
            Height          =   255
            Index           =   8
            Left            =   480
            TabIndex        =   41
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   5535
         End
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "You can download the latest GEMUS.DAT file from"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   2040
            TabIndex        =   40
            Top             =   2160
            UseMnemonic     =   0   'False
            Width           =   2415
         End
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.bu22.com/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2250
            MouseIcon       =   "frmWizard.frx":0D9E
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   2760
            UseMnemonic     =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "The Wizard has finished updating your GEMUS setup!"
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
            Index           =   7
            Left            =   480
            TabIndex        =   38
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   5535
         End
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   5
         Left            =   0
         Picture         =   "frmWizard.frx":10A8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Finished!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   1320
         TabIndex        =   21
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5085
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   6
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   24
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   600
         Width           =   6495
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":11DE
            Height          =   1155
            Index           =   9
            Left            =   1560
            TabIndex        =   29
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   3375
         End
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.bu22.com/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2250
            MouseIcon       =   "frmWizard.frx":1285
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   2760
            UseMnemonic     =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "You can download the latest GEMUS.DAT file from"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   2040
            TabIndex        =   27
            Top             =   2160
            UseMnemonic     =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   6
         Left            =   0
         Picture         =   "frmWizard.frx":158F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Emulator Setup Wizard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   1200
         TabIndex        =   25
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5205
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   22
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraBorder 
         Height          =   3615
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   600
         Width           =   6495
         Begin VB.ComboBox lstSystem 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1440
            Width           =   3975
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Which &System would you like to setup?"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   0
            Top             =   1200
            Width           =   2775
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1080
         TabIndex        =   23
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   5325
      End
      Begin VB.Image imgPic 
         Height          =   600
         Index           =   0
         Left            =   0
         Picture         =   "frmWizard.frx":1A08
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmWizard"
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
'     GEMUS - Gamebase EMUlator Script
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

Private m_xobjFunc As GBFuncs.clsFunctions
Private m_xobjRegistry As GBFuncs.clsRegistry

Private m_blnByCode As Boolean

Private m_strNotFound As String
Private m_strSelectFile As String
Private m_strChangeFile As String
Private m_strSelectEmulatorFile As String

Private m_strLastFileDir As String

Private Type WizardScreenType
    strAbortButtonText As String
    strBackButtonText As String
    strNextButtonText As String
    blnAbortEnable As Boolean
    blnBackEnable As Boolean
    blnNextEnable As Boolean
End Type
Private ma_udtScreens(6) As WizardScreenType

Private m_strAbortButton As String
Private m_strBackButton As String
Private m_strNextButton As String
Private m_strCloseButton As String
Private m_strFinishButton As String
Private m_strStartButton As String

Private Type WizardScreenEnumType
    e_SYSTEM_SELECT As Long
    e_EMU_SEARCH As Long
    e_EMU_SELECT As Long
    e_NONE_FOUND As Long
    e_FINISHED As Long
    e_INTRO As Long
End Type
Private eWS As WizardScreenEnumType
Private m_lngCurrentScreen As Long

Private ma_strSearchFileTypes() As String
Private m_lngTotalSearchFileTypes As Long

Private m_blnSearching As Boolean

Private Type FoundFileType
    strPathFile As String
    strFile As String
    strCRC As String
End Type
Private ma_udtFiles() As FoundFileType
Private m_lngTotalFiles As Long

Private m_lngTotalFoundEmulators As Long
Private m_lngNextSaveEmuIdx As Long

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'for Finding Files
Private Const MAX_PATH As Long = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Sub ClearFoundFilesForSystem(ByVal lngSystemIdx_in As Long)

    'clear all found files for the currently selected system

    Dim lngIdx As Long

    With g_udtWizard.a_udtSections(lngSystemIdx_in)
        For lngIdx = 0 To (.lngTotalEmulators - 1)
            .a_udtEmulators(lngIdx).strFound = ""
        Next lngIdx
    End With

End Sub

Private Sub CreateFile(strFile_in As String)

    Dim intFN As Integer

    On Error GoTo Eek
    
    intFN = FreeFile()
    Open strFile_in For Output As #intFN
    Print #intFN, ""
    Close #intFN
    
Eek:
    'who cares?

End Sub

Private Function CreateNestedFolder(strPath_in As String) As Boolean

    Dim a_strFolders() As String
    Dim strFolder As String
    Dim lngIdx As Long

    On Error GoTo Err_Handler

    CreateNestedFolder = False
    If InStr(strPath_in, "\") = 0 Then Exit Function
    If (Mid$(strPath_in, 2, 1) <> ":") Or (Mid$(strPath_in, 3, 1) <> "\") Then Exit Function
    a_strFolders = Split(strPath_in, "\")
    If UBound(a_strFolders) = 0 Then Exit Function
    CreateNestedFolder = True
    If m_xobjFunc.DirExists(strPath_in) = True Then Exit Function

    strFolder = ""
    For lngIdx = 0 To UBound(a_strFolders)
        strFolder = strFolder & a_strFolders(lngIdx)
        If Len(strFolder) = 2 Then strFolder = strFolder & "\"
        If m_xobjFunc.DirExists(strFolder) = False Then
            MkDir strFolder
        End If
        If Right$(strFolder, 1) <> "\" Then
            strFolder = strFolder & "\"
        End If
    Next lngIdx

    Exit Function

Err_Handler:
    CreateNestedFolder = False

End Function

Private Function EmulatorSearch(strSearchPath_in As String) As String

    'returns true if at least one emulator was found, false otherwise
    Dim lngIdx As Long
    Dim lngFindIdx As Long
    Dim blnFound As Boolean
    Dim strExt As String
    Dim lngPos As Long
    Dim lngLen As Long
    Dim strError As String
    Dim lngTotalAlreadySetup As Long
    Dim lngTemp As Long

    'setup the user interface for the search
    cmdNext.Enabled = False
    cmdAbort.Enabled = False
    cmdBack.Enabled = False
    m_blnSearching = True
    txtSearch.Visible = False
    cmdBrowse.Visible = False
    txtSearch.Enabled = False
    cmdBrowse.Enabled = False
    lblPrompt(1).Visible = False
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "Initialising", "Initialising", g_udtWizard.strMessagesFile) & "..."
    lblSearch.Visible = True
    pbSearch.Min = 0
    pbSearch.Max = 100
    pbSearch.Value = 0
    pbSearch.Visible = True
    DoEvents

    'clear all found files for the currently selected system
    Call ClearFoundFilesForSystem(lstSystem.ItemData(lstSystem.ListIndex))

    'get the search filetypes
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "PreparingForScan", "Preparing for Scan", g_udtWizard.strMessagesFile) & "..."
    lngTemp = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).lngTotalEmulators - 1
    If lngTemp < 1 Then lngTemp = 1
    pbSearch.Max = lngTemp
    DoEvents
    ReDim ma_strSearchFileTypes(0)
    m_lngTotalSearchFileTypes = 0
    With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
        For lngIdx = 0 To (.lngTotalEmulators - 1)
            lngPos = InStrRev(.a_udtEmulators(lngIdx).strFilename, ".")
            If lngPos > 1 Then
                lngLen = Len(.a_udtEmulators(lngIdx).strFilename)
                If lngPos < lngLen Then
                    strExt = LCase$(Right$(.a_udtEmulators(lngIdx).strFilename, lngLen - lngPos + 1))
                    blnFound = False
                    For lngFindIdx = 0 To (m_lngTotalSearchFileTypes - 1)
                        If strExt = ma_strSearchFileTypes(lngFindIdx) Then
                            blnFound = True
                            Exit For
                        End If
                    Next lngFindIdx
                    If blnFound = False Then
                        ReDim Preserve ma_strSearchFileTypes(m_lngTotalSearchFileTypes)
                        ma_strSearchFileTypes(m_lngTotalSearchFileTypes) = strExt
                        m_lngTotalSearchFileTypes = m_lngTotalSearchFileTypes + 1
                    End If
                End If
            End If
            pbSearch.Value = lngIdx
            DoEvents
        Next lngIdx
    End With
    If m_lngTotalSearchFileTypes = 0 Then
        EmulatorSearch = "No Filetypes found in Emulator filenames in GEMUS.DAT!"
        m_blnSearching = False
        Exit Function
    End If
    'get a list of files to CRC check
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "ScanningSearchPath", "Scanning Search Path (may take a while)", g_udtWizard.strMessagesFile) & "..."
    lblSearchFolder.Caption = strSearchPath_in
    lblSearchFolder.Visible = True
    pbSearch.Value = 0
    DoEvents
    strError = GetFiles(strSearchPath_in)
    lblSearchFolder.Caption = ""
    lblSearchFolder.Visible = False
    If Len(strError) > 0 Then
        EmulatorSearch = strError
        m_blnSearching = False
        Exit Function
    End If
    'get the CRCs
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "Obtaining File CRCs", "Obtaining File CRCs", g_udtWizard.strMessagesFile) & "..."
    pbSearch.Min = 0
    pbSearch.Value = 0
    pbSearch.Max = m_lngTotalFiles + 1
    DoEvents
    strError = GetCRCs()
    If Len(strError) > 0 Then
        EmulatorSearch = strError
        m_blnSearching = False
        Exit Function
    End If
    'do the crc check
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "CheckingForMatchingEmulators", "Checking for Matching Emulators", g_udtWizard.strMessagesFile) & "..."
    pbSearch.Min = 0
    pbSearch.Value = 0
    lngTemp = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).lngTotalEmulators - 1
    If lngTemp < 1 Then lngTemp = 1
    pbSearch.Max = lngTemp
    DoEvents
    strError = FindEmulators()
    If Len(strError) > 0 Then
        EmulatorSearch = strError
        m_blnSearching = False
        Exit Function
    End If
    'remove duplicates from emulators.ini file
    lblSearch.Caption = GetINIMessageFromFile("Wizard", "RemovingDuplicates", "Removing Duplicates from Current GEMUS Setup", g_udtWizard.strMessagesFile) & "..."
    pbSearch.Min = 0
    pbSearch.Value = 0
    DoEvents
    strError = RemoveDuplicates(lngTotalAlreadySetup)
    If Len(strError) > 0 Then
        EmulatorSearch = strError
        m_blnSearching = False
        Exit Function
    End If
    If lngTotalAlreadySetup = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).lngTotalEmulators Then
        EmulatorSearch = "~~NONE~~"
        m_blnSearching = False
        Exit Function
    End If

    'finished!
    EmulatorSearch = ""
    m_blnSearching = False

End Function
Private Sub FillSelectorLV()

    Dim LItem As ListItem
    Dim lngIdx As Long

    m_blnByCode = True
    With lvSelect
        .Enabled = False
        .Visible = False
        .ListItems.Clear
        With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
            For lngIdx = 0 To (.lngTotalEmulators - 1)
                With .a_udtEmulators(lngIdx)
                    If .strFound <> "|~~IGNORE~~|" Then
                        Set LItem = lvSelect.ListItems.Add(, "E" & lngIdx, .strEmulatorName)
                        If Len(.strFound) > 0 Then
                            LItem.SubItems(1) = .strFound
                            LItem.Checked = True
                            LItem.Bold = True
                        Else
                            LItem.SubItems(1) = m_strNotFound
                            LItem.Checked = False
                            LItem.Bold = False
                        End If
                        If Len(.strInfo) > 0 Then
                            LItem.Tag = .strInfo
                            LItem.SmallIcon = "info"
                            LItem.Icon = "info"
                        Else
                            LItem.Tag = ""
                        End If
                        Set LItem = Nothing
                    End If
                End With
            Next lngIdx
        End With
        .Visible = True
        .Enabled = True
    End With
    m_blnByCode = False

End Sub

Private Function FindEmulators() As String

    'finds emulators (no shit)

    Dim lngEmuIdx As Long
    Dim lngFileIdx As Long
    Dim lngCRCIdx As Long
    Dim blnCRCMatch As Boolean

    m_lngTotalFoundEmulators = 0
    With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
        For lngEmuIdx = 0 To (.lngTotalEmulators - 1)
            'search for this emulator!
            With .a_udtEmulators(lngEmuIdx)
                If .lngTotalCRCs = 1 And .a_strCRC32(0) = "*" Then
                    'just filename search
                    For lngFileIdx = 0 To (m_lngTotalFiles - 1)
                        If UCase$(ma_udtFiles(lngFileIdx).strFile) = UCase$(.strFilename) Then
                            .strFound = ma_udtFiles(lngFileIdx).strPathFile
                            m_lngTotalFoundEmulators = m_lngTotalFoundEmulators + 1
                            Exit For
                        End If
                    Next lngFileIdx
                Else
                    'crc search
                    For lngFileIdx = 0 To (m_lngTotalFiles - 1)
                        If UCase$(ma_udtFiles(lngFileIdx).strFile) = UCase$(.strFilename) Then
                            'test each crc
                            blnCRCMatch = False
                            For lngCRCIdx = 0 To (.lngTotalCRCs - 1)
                                If ma_udtFiles(lngFileIdx).strCRC = UCase$(.a_strCRC32(lngCRCIdx)) Then
                                    .strFound = ma_udtFiles(lngFileIdx).strPathFile
                                    m_lngTotalFoundEmulators = m_lngTotalFoundEmulators + 1
                                    blnCRCMatch = True
                                    Exit For
                                End If
                            Next lngCRCIdx
                            If blnCRCMatch = True Then
                                Exit For
                            End If
                        End If
                    Next lngFileIdx
                End If
            End With
            'update the progress bar
            pbSearch.Value = lngEmuIdx
            DoEvents
        Next lngEmuIdx
    End With

End Function
Private Sub FindFilesMatch(strPath_in As String)
    
    Dim strFilename As String           'Working filename variable
    Dim strDirName As String            'Working SubDirectory variable
    Dim a_strDirNames() As String       'Buffer for directory name entries
    Dim intSubDirCount As Integer       'Number of directories in current path
    Dim intSubDirIdx As Integer         'loop index
    Dim hSearch As Long                 'Search Handle
    Dim udtWFD As WIN32_FIND_DATA       'structure for API
    Dim intContinue As Integer          'continue searching flag
    Dim lngFTIdx As Long
    
    'get folder with backslash
    If Right(strPath_in, 1) <> "\" Then strPath_in = strPath_in & "\"
    
    ' Search for subdirectories.
    intSubDirCount = 0
    ReDim a_strDirNames(intSubDirCount)
    intContinue = True
    hSearch = FindFirstFile(strPath_in & "*", udtWFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While intContinue
            strDirName = StripNulls(udtWFD.cFileName)
            'Ignore the current and encompassing directories.
            If (strDirName <> ".") And (strDirName <> "..") Then
                'Check for directory with bitwise comparison.
                If GetFileAttributes(strPath_in & strDirName) And FILE_ATTRIBUTE_DIRECTORY Then
                    a_strDirNames(intSubDirCount) = strDirName
                    intSubDirCount = intSubDirCount + 1
                    ReDim Preserve a_strDirNames(intSubDirCount)
                End If
            End If
            'Get next subdirectory
            intContinue = FindNextFile(hSearch, udtWFD)
            DoEvents
        Loop
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Walk through this directory
    For lngFTIdx = 0 To (m_lngTotalSearchFileTypes - 1)
        hSearch = FindFirstFile(strPath_in & "*" & ma_strSearchFileTypes(lngFTIdx), udtWFD)
        intContinue = True
        If hSearch <> INVALID_HANDLE_VALUE Then
            While intContinue
                strFilename = StripNulls(udtWFD.cFileName)
                If (strFilename <> ".") And (strFilename <> "..") Then
                    If Not (GetFileAttributes(strPath_in & strFilename) And FILE_ATTRIBUTE_DIRECTORY) Then
                        'add to files
                        ReDim Preserve ma_udtFiles(m_lngTotalFiles)
                        ma_udtFiles(m_lngTotalFiles).strPathFile = strPath_in & strFilename
                        ma_udtFiles(m_lngTotalFiles).strFile = strFilename
                        m_lngTotalFiles = m_lngTotalFiles + 1
                    End If
                End If
                'Get next file
                intContinue = FindNextFile(hSearch, udtWFD)
                DoEvents
            Wend
            'tidy up
            intContinue = FindClose(hSearch)
        End If
        DoEvents
    Next lngFTIdx
    
    
    'Any sub-directories?
    If intSubDirCount > 0 Then
        'Recursively process them
        For intSubDirIdx = 0 To intSubDirCount - 1
            lblSearchFolder.Caption = strPath_in & a_strDirNames(intSubDirIdx)
            DoEvents
            Call FindFilesMatch(strPath_in & a_strDirNames(intSubDirIdx))
            DoEvents
        Next intSubDirIdx
    End If

End Sub

Private Function GetCRCs() As String

    Dim lngIdx As Long
    Dim objCRC As clsCRC

    On Error GoTo Err_Handler

    Set objCRC = New clsCRC

    For lngIdx = 0 To (m_lngTotalFiles - 1)
        With ma_udtFiles(lngIdx)
            .strCRC = Hex(objCRC.CalculateFile(.strPathFile))
        End With
        pbSearch.Value = lngIdx
        DoEvents
    Next lngIdx

    Set objCRC = Nothing
    Exit Function

Err_Handler:
    GetCRCs = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
    Set objCRC = Nothing

End Function

Private Sub GetFilenameForEmulator()

    Dim strFile As String
    Dim strPath As String
    Dim LItem As ListItem

    On Error GoTo CancelPressed

    Set LItem = lvSelect.SelectedItem
    If LItem Is Nothing Then Exit Sub

    If LItem.SubItems(1) = m_strNotFound Then
        strFile = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).a_udtEmulators(Val(Right$(LItem.Key, Len(LItem.Key) - 1))).strFilename
        If m_xobjFunc.DirExists(m_strLastFileDir) = True Then
            strPath = m_strLastFileDir
        Else
            strPath = App.Path
        End If
    Else
        Call m_xobjFunc.GetFileAndPath(LItem.SubItems(1), strPath, strFile)
        If m_xobjFunc.FileExists(LItem.SubItems(1), False, "") = False Then
            strFile = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).a_udtEmulators(Val(Right$(LItem.Key, Len(LItem.Key) - 1))).strFilename
            If m_xobjFunc.DirExists(strPath) = False Then
                If m_xobjFunc.DirExists(m_strLastFileDir) = True Then
                    strPath = m_strLastFileDir
                Else
                    strPath = App.Path
                End If
            End If
        End If
    End If
    
    With cDialog
        .CancelError = True
        .DialogTitle = m_strSelectEmulatorFile
        .InitDir = strPath
        .Filename = strFile
        .Filter = "*.*|*.*|*.exe|*.exe|*.com|*.com|*.bat|*.bat|*.cmd|*.cmd"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
        .ShowOpen
        'new filename to set
        g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).a_udtEmulators(Val(Right$(LItem.Key, Len(LItem.Key) - 1))).strFound = .Filename
        LItem.SubItems(1) = .Filename
        cmdSelectFile.Caption = m_strChangeFile
        'update last traversed path
        Call m_xobjFunc.GetFileAndPath(.Filename, strPath, strFile)
        m_strLastFileDir = strPath
        m_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Wizard", "EmulatorFilePath", m_strLastFileDir, True
    End With

CancelPressed:

    Set LItem = Nothing
    lvSelect.SetFocus

End Sub

Private Function RemoveDuplicates(lngTotalAlreadySetup_out As Long) As String

    Dim lngGemusIdx As Long
    Dim a_strGemusEmus() As String
    Dim lngTotalGemusEmus As Long
    Dim strEmuName As String
    Dim lngEmuIdx As Long

    'get the emulator names from the emulators.ini file
    lngTotalGemusEmus = 0
    ReDim a_strGemusEmus(0)
    lngGemusIdx = 1
    strEmuName = m_xobjFunc.INI_GetSetting("Emulators", Trim$(lngGemusIdx), "", g_udtWizard.strEmuFile)
    Do Until strEmuName = ""
        ReDim Preserve a_strGemusEmus(lngTotalGemusEmus)
        a_strGemusEmus(lngTotalGemusEmus) = UCase$(strEmuName)
        lngTotalGemusEmus = lngTotalGemusEmus + 1
        lngGemusIdx = lngGemusIdx + 1
        strEmuName = m_xobjFunc.INI_GetSetting("Emulators", Trim$(lngGemusIdx), "", g_udtWizard.strEmuFile)
    Loop

    m_lngNextSaveEmuIdx = lngGemusIdx
    If lngTotalGemusEmus = 0 Then
        RemoveDuplicates = ""
    Else
        pbSearch.Min = 0
        pbSearch.Value = 0
        pbSearch.Max = lngTotalGemusEmus
        DoEvents
        lngTotalAlreadySetup_out = 0
        For lngGemusIdx = 0 To (lngTotalGemusEmus - 1)
            With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
                For lngEmuIdx = 0 To (.lngTotalEmulators - 1)
                    With .a_udtEmulators(lngEmuIdx)
                        If Len(.strFound) > 0 Then
                            If UCase$(.strEmulatorName) = a_strGemusEmus(lngGemusIdx) Then
                                .strFound = "|~~IGNORE~~|"
                                m_lngTotalFoundEmulators = m_lngTotalFoundEmulators - 1
                                lngTotalAlreadySetup_out = lngTotalAlreadySetup_out + 1
                            End If
                        End If
                    End With
                Next lngEmuIdx
            End With
            pbSearch.Value = lngGemusIdx
            DoEvents
        Next lngGemusIdx
    End If
    
End Function

Private Function SaveEmulators() As String

    'save emulators to the emulators.ini file

    Dim lngEmuIdx As Long
    Dim strDestScriptFolder As String
    Dim strSource As String
    Dim strDest As String
    Dim a_strDone() As String
    Dim lngTotalDone As Long
    Dim lngDoneIdx As Long
    Dim blnDone As Boolean
    Dim mbRes As VbMsgBoxResult
    Dim blnSkip As Boolean
    Dim blnSuccess As Boolean
    Dim lngFTIdx As Long
    Dim strFileTypes As String
    Dim strPath As String
    Dim strFile As String
    Dim strAssFile As String
    Dim strMBQuestion As String
    Dim strMBTitle As String
    Dim strYes As String
    Dim strNo As String
    Dim strCancel As String

    On Error GoTo Err_Handler

    strMBQuestion = GetINIMessageFromFile("Wizard", "OverwriteScriptQuestion", "The Script for this emulator already exists. Do you want to overwrite it with the one pointed to in the GEMUS.DAT file?", g_udtWizard.strMessagesFile)
    strMBTitle = GetINIMessageFromFile("Wizard", "OverwriteScriptTitle", "Overwrite the GEMUS Script for", g_udtWizard.strMessagesFile)
    
    strYes = GetINIMessageFromFile("Wizard", "OverwriteYes", "Click 'YES' to use the new GEMUS.DAT script.", g_udtWizard.strMessagesFile)
    strNo = GetINIMessageFromFile("Wizard", "OverwriteNo", "Click 'NO' to keep your existing script.", g_udtWizard.strMessagesFile)
    strCancel = GetINIMessageFromFile("Wizard", "OverwriteCancel", "Click 'CANCEL' to skip adding this emulator.", g_udtWizard.strMessagesFile)
    
    lngTotalDone = 0
    ReDim a_strDone(0)
    With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
        For lngEmuIdx = 0 To (.lngTotalEmulators - 1)
            blnSkip = False
            With .a_udtEmulators(lngEmuIdx)
                If (Len(.strFound) > 0) And (.strFound <> "|~~IGNORE~~|") Then
                    'copy across the emulator script file?
                    If g_udtWizard.blnGlobalDAT = True Then
                        strDestScriptFolder = g_udtWizard.strGBFolder & "\Scripts"
                        If m_xobjFunc.DirExists(strDestScriptFolder) = False Then
                            Call MkDir(strDestScriptFolder)
                        End If
                        'does it already exist in the dest folder?
                        strSource = g_udtWizard.strScriptFolder & "\" & .strScriptFile
                        strDest = strDestScriptFolder & "\" & .strScriptFile
                        If m_xobjFunc.FileExists(strDest, False, "") = True Then
                            'SCRIPT FILE ALREADY EXISTS!
                            blnDone = False
                            If lngTotalDone > 0 Then
                                For lngDoneIdx = 0 To (lngTotalDone - 1)
                                    If .strScriptFile = a_strDone(lngDoneIdx) Then
                                        blnDone = True
                                        Exit For
                                    End If
                                Next lngDoneIdx
                            End If
                            If Not blnDone Then
                                mbRes = MsgBox(.strEmulatorName & "  (" & .strScriptFile & ")" & vbCrLf & vbCrLf & strMBQuestion & vbCrLf & vbCrLf & strYes & vbCrLf & strNo & vbCrLf & strCancel, vbYesNoCancel, strMBTitle & " """ & .strEmulatorName & """?")
                                If mbRes = vbYes Then
                                    'overwrite the script file
                                    If CopyFile(strSource, strDest, 0&) = 0 Then
                                        SaveEmulators = "Failed copying script file: -" & vbCrLf & vbCrLf & strSource & "   ->   " & strDest
                                        Exit Function
                                    End If
                                    'add the script file to the done array for next time
                                    ReDim Preserve a_strDone(lngTotalDone)
                                    a_strDone(lngTotalDone) = .strScriptFile
                                    lngTotalDone = lngTotalDone + 1
                                ElseIf mbRes = vbCancel Then
                                    'don't add this emulator
                                    blnSkip = True
                                End If
                            Else
                                'already copied across from source scripts folder
                                'so we don't need to copy across or ask the user
                            End If
                        Else
                            'SCRIPT FILE DOESN'T EXIST - JUST COPY ACROSS
                            If CopyFile(strSource, strDest, 1&) = 0 Then
                                SaveEmulators = "Failed copying script file: -" & vbCrLf & vbCrLf & strSource & "   ->   " & strDest
                                Exit Function
                            End If
                        End If
                    End If
                        
                    'are we still ok?
                    If blnSkip = False Then
                    
                        'ADD EMULATORS TO THE EMULATORS.INI FILE
                        
                        'emulators section
                        blnSuccess = m_xobjFunc.INI_WriteSetting("Emulators", Trim$(m_lngNextSaveEmuIdx), .strEmulatorName, g_udtWizard.strEmuFile)
                        'use short filenames
                        blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting(.strEmulatorName, "UseShortFilenames", IIf((.blnUseShort), "1", "0"), g_udtWizard.strEmuFile)
                        'supported file types
                        strFileTypes = ""
                        For lngFTIdx = 0 To (.lngTotalFileTypes - 1)
                            strFileTypes = strFileTypes & .a_strFileTypes(lngFTIdx) & ";"
                        Next lngFTIdx
                        If Right$(strFileTypes, 1) = ";" Then
                            strFileTypes = Left$(strFileTypes, Len(strFileTypes) - 1)
                        End If
                        blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting(.strEmulatorName, "SupportedTypes", strFileTypes, g_udtWizard.strEmuFile)
                        'emulator path\file
                        blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting(.strEmulatorName, "EmulatorPathFile", .strFound, g_udtWizard.strEmuFile)
                        'script file
                        blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting(.strEmulatorName, "ScriptFile", .strScriptFile, g_udtWizard.strEmuFile)
                        'associated files
                        For lngFTIdx = 0 To (.lngTotalAssFiles - 1)
                            Call m_xobjFunc.GetFileAndPath(.strFound, strPath, strFile)
                            strAssFile = strPath
                            If Right$(strAssFile, 1) <> "\" Then
                                strAssFile = strAssFile & "\"
                            End If
                            strAssFile = strAssFile & .a_strAssFiles(lngFTIdx)
                            Call m_xobjFunc.GetFileAndPath(strAssFile, strPath, strFile)
                            blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting(.strEmulatorName, Trim$(lngFTIdx + 1), strAssFile, g_udtWizard.strEmuFile)
                            If m_xobjFunc.FileExists(strAssFile, False, "") = False Then
                                'attempt to create it, don't worry if it can't
                                If InStr(.a_strAssFiles(lngFTIdx), "\") > 0 Then
                                    If CreateNestedFolder(strPath) = True Then
                                        Call CreateFile(strAssFile)
                                    End If
                                Else
                                    Call CreateFile(strAssFile)
                                End If
                            End If
                        Next lngFTIdx
                        
                        'not needed now that we auto calculate the valid filetypes when running emulators and selecting game files
                        
                        'global file types
'                        strFileTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", "", g_udtWizard.strEmuFile)
'                        strFileTypes = Trim$(UCase$(strFileTypes))
'                        If Left$(strFileTypes, 1) <> ";" Then
'                            strFileTypes = ";" & strFileTypes
'                        End If
'                        If Right$(strFileTypes, 1) <> ";" Then
'                            strFileTypes = strFileTypes & ";"
'                        End If
'                        If Len(strFileTypes) = 2 Then
'                            strFileTypes = ";"
'                        End If
'                        For lngFTIdx = 0 To (.lngTotalFileTypes - 1)
'                            If InStr(strFileTypes, ";" & UCase$(.a_strFileTypes(lngFTIdx)) & ";") = 0 Then
'                                strFileTypes = strFileTypes & UCase$(.a_strFileTypes(lngFTIdx)) & ";"
'                            End If
'                        Next lngFTIdx
'                        If Right$(strFileTypes, 1) = ";" Then
'                            strFileTypes = Left$(strFileTypes, Len(strFileTypes) - 1)
'                        End If
'                        If Left$(strFileTypes, 1) = ";" Then
'                            strFileTypes = Right$(strFileTypes, Len(strFileTypes) - 1)
'                        End If
'                        blnSuccess = blnSuccess Or m_xobjFunc.INI_WriteSetting("File Types", "SupportTypes", strFileTypes, g_udtWizard.strEmuFile)
                        
                        'all ok?
                        If Not blnSuccess Then
                            SaveEmulators = "Failed writing emulator settings to " & g_udtWizard.strEmuFile
                            Exit Function
                        End If
                        
                        'add one to the emulator counter for the ini file
                        m_lngNextSaveEmuIdx = m_lngNextSaveEmuIdx + 1
                    End If
                        
                End If
            End With
        Next lngEmuIdx
    End With

    SaveEmulators = ""
    Exit Function

Err_Handler:
    SaveEmulators = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Function StripNulls(strSource_inout As String) As String
 
    '------------------------------------------------------------------------------
    'Used by FindFilesMatch() to strip trailing NULL characters off the passed in string
    '
    'Modify:    strSource_inout:    The source string to strip the NULLs from.
    '                               Returned with no trailing NULLs
    '
    'Returns:   copy of the source string but with trailing NULLs
    '------------------------------------------------------------------------------
    
    'test for NULLs and remove
    If (InStr(strSource_inout, Chr(0)) > 0) Then
        strSource_inout = Left(strSource_inout, InStr(strSource_inout, Chr(0)) - 1)
    End If
    
    'copy to return string
    StripNulls = strSource_inout
    
End Function
Private Function GetFiles(strSearchPath_in As String) As String

    'gets a list of files to search in the scan

    Dim strRootPath As String

    On Error GoTo Err_Handler

    'init vars
    strRootPath = strSearchPath_in
    GetFiles = ""
    m_lngTotalFiles = 0
    ReDim ma_udtFiles(0)

    'fill the array
    Call FindFilesMatch(strRootPath)

    'exit!
    Exit Function

Err_Handler:
    GetFiles = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Sub FillSystems()

    Dim lngIdx As Long

    lstSystem.Enabled = False
    lstSystem.Visible = False
    lstSystem.Clear
    With g_udtWizard
        For lngIdx = 0 To (.lngTotalSections - 1)
            If .a_udtSections(lngIdx).lngTotalEmulators > 0 Then
                lstSystem.AddItem .a_udtSections(lngIdx).strName
                lstSystem.ItemData(lstSystem.NewIndex) = lngIdx
            End If
        Next lngIdx
    End With
    lstSystem.Enabled = True
    lstSystem.Visible = True

End Sub

Private Sub ShowScreen(ByVal lngScreenIdx_in As Long)

    With ma_udtScreens(lngScreenIdx_in)
        cmdAbort.Enabled = .blnAbortEnable
        cmdAbort.Visible = .blnAbortEnable
        cmdAbort.Caption = .strAbortButtonText
        cmdBack.Enabled = .blnBackEnable
        cmdBack.Visible = .blnBackEnable
        cmdBack.Caption = .strBackButtonText
        cmdNext.Enabled = .blnNextEnable
        cmdNext.Visible = .blnNextEnable
        cmdNext.Caption = .strNextButtonText
        picWiz(lngScreenIdx_in).ZOrder 0
    End With

End Sub

Private Sub cmdAbort_Click()
    Unload Me
End Sub


Private Sub cmdBack_Click()

    Screen.MousePointer = vbHourglass
    DoEvents

    With ma_udtScreens(m_lngCurrentScreen)
    
        Select Case m_lngCurrentScreen
            
            Case eWS.e_EMU_SEARCH
            
                'back to system select
                m_lngCurrentScreen = eWS.e_SYSTEM_SELECT
                Call ShowScreen(m_lngCurrentScreen)
                lstSystem.SetFocus
            
            Case eWS.e_EMU_SELECT

                'back to search screen
                pbSearch.Visible = False
                lblSearch.Visible = False
                lblSearchFolder.Visible = False
                lblPrompt(1).Visible = True
                txtSearch.Visible = True
                txtSearch.Enabled = True
                cmdBrowse.Visible = True
                cmdBrowse.Enabled = True
                m_lngCurrentScreen = eWS.e_EMU_SEARCH
                Call ShowScreen(m_lngCurrentScreen)
                If Len(txtSearch) > 0 Then
                    cmdNext.Visible = True
                    cmdNext.Enabled = True
                End If
                txtSearch.SetFocus

        End Select
    
    End With

    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdBrowse_Click()

    If m_xobjFunc.DirExists(txtSearch.Text) Then
        g_strBrowseForFolder = txtSearch.Text
    Else
        g_strBrowseForFolder = App.Path
    End If

    frmBrowseFolder.Show vbModal, Me

    If g_strBrowseForFolder <> "" Then
        txtSearch.Text = g_strBrowseForFolder
    End If
    txtSearch.SetFocus

End Sub

Private Sub cmdNext_Click()

    Dim blnInvalidSearchPath As Boolean
    Dim strSearchPath As String
    Dim a_strFileTypes() As String
    Dim lngTotalFileTypes As Long
    Dim strError As String
    Dim lngIdx As Long
    Dim LItem As ListItem
    Dim blnOneSelected As Boolean

    Screen.MousePointer = vbHourglass
    DoEvents

    With ma_udtScreens(m_lngCurrentScreen)
    
        Select Case m_lngCurrentScreen
            
            Case eWS.e_INTRO
            
                'show the system select screen
                m_lngCurrentScreen = eWS.e_SYSTEM_SELECT
                Call ShowScreen(m_lngCurrentScreen)
                lstSystem.SetFocus

            Case eWS.e_SYSTEM_SELECT
            
                'show the search screen
                m_lngCurrentScreen = eWS.e_EMU_SEARCH
                Call ShowScreen(m_lngCurrentScreen)
                lblSearch.Visible = False
                lblSearchFolder.Visible = False
                pbSearch.Visible = False
                txtSearch.Enabled = True
                txtSearch.Visible = True
                cmdBrowse.Enabled = True
                cmdBrowse.Visible = True
                If Len(txtSearch) > 0 Then
                    cmdNext.Enabled = True
                    cmdNext.Visible = True
                End If
                txtSearch.SetFocus

            Case eWS.e_EMU_SEARCH

                'validate the search path
                blnInvalidSearchPath = True
                If Len(txtSearch) > 2 Then
                    If Mid$(txtSearch, 2, 2) = ":\" Then
                        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(Left$(txtSearch, 1))) > 0 Then
                            blnInvalidSearchPath = False
                        End If
                    End If
                End If
                If blnInvalidSearchPath Then
                    'absolute path required
                    MsgBox GetINIMessageFromFile("Wizard", "Error_AbsolutePathRequired", "An absolute path is required! E.g. 'C:\Emulators'", g_udtWizard.strMessagesFile), vbExclamation, "GEMUS"
                    txtSearch.SetFocus
                Else
                    'do we have an existing path?
                    strSearchPath = txtSearch
                    If Right$(strSearchPath, 1) = "\" Then
                        If Len(strSearchPath) > 3 Then
                            strSearchPath = Left$(strSearchPath, Len(strSearchPath) - 1)
                        End If
                    End If
                    If m_xobjFunc.DirExists(strSearchPath) = False Then
                        'path not found
                        MsgBox GetINIMessageFromFile("Wizard", "Error_PathNotFound", "The specified path was not found!", g_udtWizard.strMessagesFile), vbExclamation, "GEMUS"
                        txtSearch.SetFocus
                    Else
                        'SEARCH
                        strError = EmulatorSearch(strSearchPath)
                        If strError = "" Then
                            'found! - init the selector screen
                            Call FillSelectorLV
                            If lvSelect.ListItems.Count > 0 Then
                                lvSelect.ListItems(1).Selected = False
                                Set lvSelect.SelectedItem = Nothing
                            End If
                            m_lngCurrentScreen = eWS.e_EMU_SELECT
                            Call ShowScreen(m_lngCurrentScreen)
                            For Each LItem In lvSelect.ListItems
                                If LItem.Checked Then
                                    cmdNext.Enabled = True
                                    cmdNext.Visible = True
                                    Exit For
                                End If
                            Next LItem
                            cmdSelectFile.Enabled = False
                            cmdWebSite.Enabled = False
                            lvSelect.SetFocus
                        ElseIf strError = "~~NONE~~" Then
                            'none to setup!
                            m_lngCurrentScreen = eWS.e_NONE_FOUND
                            Call ShowScreen(m_lngCurrentScreen)
                        Else
                            'ERROR!
                            MsgBox strError, vbCritical, "GEMUS"
                            Screen.MousePointer = vbDefault
                            Unload Me
                        End If
                    End If
                End If
                
            Case eWS.e_EMU_SELECT
                'end screen - save the new emulator(s) details to the emulators.ini file
                cmdNext.Enabled = False
                DoEvents
                blnOneSelected = False
                For Each LItem In lvSelect.ListItems
                    If LItem.Checked = True Then
                        blnOneSelected = True
                        Exit For
                    End If
                Next LItem
                If blnOneSelected = True Then
                    With g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex))
                        For Each LItem In lvSelect.ListItems
                            If LItem.Checked = False Then
                                .a_udtEmulators(Val(Right$(LItem.Key, Len(LItem.Key) - 1))).strFound = ""
                            End If
                        Next LItem
                    End With
                    strError = SaveEmulators()
                    If strError <> "" Then
                        MsgBox strError, vbCritical, "GEMUS"
                        Screen.MousePointer = vbDefault
                        Unload Me
                    Else
                        m_lngCurrentScreen = eWS.e_FINISHED
                        Call ShowScreen(m_lngCurrentScreen)
                    End If
                Else
                    cmdNext.Enabled = True
                    cmdNext.Visible = True
                    lvSelect.SetFocus
                End If
            Case eWS.e_NONE_FOUND
                'close the window
                m_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Wizard", "SearchPath", txtSearch.Text, True
                Screen.MousePointer = vbDefault
                Unload Me
            Case eWS.e_FINISHED
                'close the window
                g_udtWizard.blnAborted = False
                m_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Wizard", "SearchPath", txtSearch.Text, True
                Screen.MousePointer = vbDefault
                Unload Me
                
        End Select
    
    End With

    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSelectAll_Click()

    Dim LItem As ListItem

    Screen.MousePointer = vbHourglass
    DoEvents
    cmdSelectAll.Enabled = False
    With lvSelect
        .Visible = False
        .Enabled = False
        For Each LItem In .ListItems
            With LItem
                If .SubItems(1) <> m_strNotFound Then
                    .Checked = True
                    .Bold = True
                    cmdNext.Enabled = True
                    cmdNext.Visible = True
                End If
            End With
        Next LItem
        .Visible = True
        .Enabled = True
    End With
    cmdSelectAll.Enabled = True
    Screen.MousePointer = vbDefault
    lvSelect.SetFocus

End Sub

Private Sub cmdSelectFile_Click()

    Call GetFilenameForEmulator

End Sub

Private Sub cmdSelectNone_Click()

    Dim LItem As ListItem

    cmdNext.Enabled = False
    cmdNext.Visible = False
    Screen.MousePointer = vbHourglass
    DoEvents
    cmdSelectNone.Enabled = False
    With lvSelect
        .Visible = False
        .Enabled = False
        For Each LItem In .ListItems
            LItem.Checked = False
            LItem.Bold = False
        Next LItem
        .Visible = True
        .Enabled = True
    End With
    cmdSelectNone.Enabled = True
    Screen.MousePointer = vbDefault
    lvSelect.SetFocus

End Sub

Private Sub cmdWebSite_Click()

    Dim strURL As String
    Dim LItem As ListItem

    Set LItem = lvSelect.SelectedItem
    If LItem Is Nothing Then Exit Sub
    strURL = g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).a_udtEmulators(Val(Right$(LItem.Key, Len(LItem.Key) - 1))).strURL
    If Len(strURL) > 0 Then
        If InStr(strURL, "@") > 0 Then
            If Left$(UCase$(strURL), 7) <> "MAILTO:" Then
                strURL = "mailto:" & strURL
            End If
        End If
        Call m_xobjFunc.ShellExecuteFile(strURL, "", False)
    Else
        cmdWebSite.Enabled = False
    End If
    Set LItem = Nothing

End Sub

Private Sub Form_Load()

    Dim strSearchPath As String
    Const DesignScaleHeight As Single = 5055    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6735     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'create objects
    Set m_xobjFunc = New GBFuncs.clsFunctions
    Set m_xobjRegistry = New GBFuncs.clsRegistry

    'read in the language ini text (with shortcuts)
    m_strAbortButton = GetINIMessageFromFile("Wizard", "AbortButton", "&Abort", g_udtWizard.strMessagesFile)
    m_strBackButton = "< " & GetINIMessageFromFile("Wizard", "BackButton", "&Back", g_udtWizard.strMessagesFile)
    m_strNextButton = GetINIMessageFromFile("Wizard", "NextButton", "&Next", g_udtWizard.strMessagesFile) & " >"
    m_strCloseButton = GetINIMessageFromFile("Wizard", "CloseButton", "&Close", g_udtWizard.strMessagesFile)
    m_strFinishButton = GetINIMessageFromFile("Wizard", "FinishButton", "&Finish", g_udtWizard.strMessagesFile)
    m_strStartButton = GetINIMessageFromFile("Wizard", "StartButton", "&Start", g_udtWizard.strMessagesFile) & " >"
    lblPrompt(0) = GetINIMessageFromFile("Wizard", "SystemPrompt", "Select the &System would you like to setup", g_udtWizard.strMessagesFile) & ":"
    lblPrompt(1) = GetINIMessageFromFile("Wizard", "SearchPathPrompt", "Search &Path", g_udtWizard.strMessagesFile) & ":"
    lblPrompt(2) = GetINIMessageFromFile("Wizard", "SelectEmulatorsPrompt", "Select the &Emulators you wish to setup", g_udtWizard.strMessagesFile) & ":"
    cmdSelectAll.Caption = GetINIMessageFromFile("Wizard", "SelectAllButton", "Selec&t All", g_udtWizard.strMessagesFile)
    cmdSelectNone.Caption = GetINIMessageFromFile("Wizard", "SelectNoneButton", "Select N&one", g_udtWizard.strMessagesFile)
    m_strSelectFile = GetINIMessageFromFile("Wizard", "SelectFileButton", "Select F&ile", g_udtWizard.strMessagesFile)
    m_strChangeFile = GetINIMessageFromFile("Wizard", "ChangeFileButton", "Change F&ile", g_udtWizard.strMessagesFile)
    cmdWebSite.Caption = GetINIMessageFromFile("Wizard", "WebSiteButton", "&Web Site", g_udtWizard.strMessagesFile)

    'read in the language ini text (without shortcuts)
    m_strNotFound = GetINIMessageFromFile("Wizard", "NotFound", "Not Found", g_udtWizard.strMessagesFile)
    lblPrompt(4) = GetINIMessageFromFile("Wizard", "NoneToSetup", "Your GEMUS setup appears to be up-to-date for this system. All the emulators in the GEMUS.DAT file are already setup for your database!", g_udtWizard.strMessagesFile)
    lblPrompt(9) = GetINIMessageFromFile("Wizard", "WelcomeMessage", "This Wizard uses the GEMUS.DAT file located in your GameBase folder.  This file contains emulator definitions needed to automatically setup emulators for GameBase.", g_udtWizard.strMessagesFile)
    lblPrompt(10) = GetINIMessageFromFile("Wizard", "DownloadDATPrompt", "You can download the latest GEMUS.DAT file from", g_udtWizard.strMessagesFile)
    lblPrompt(11) = lblPrompt(10)
    lblPrompt(7) = GetINIMessageFromFile("Wizard", "FinishedText", "The Wizard has finished updating your GEMUS setup!", g_udtWizard.strMessagesFile)
    lblPrompt(8) = GetINIMessageFromFile("Wizard", "FinishedText2", "Did the Wizard not have the emulator you wanted?", g_udtWizard.strMessagesFile)
    lblTitle(0) = GetINIMessageFromFile("Wizard", "TitlePickASystem", "Pick a System", g_udtWizard.strMessagesFile)
    lblTitle(1) = GetINIMessageFromFile("Wizard", "TitleSearchForEmulators", "Search for Emulators", g_udtWizard.strMessagesFile)
    lblTitle(2) = GetINIMessageFromFile("Wizard", "TitleSelectEmulators", "Select Emulators", g_udtWizard.strMessagesFile)
    lblTitle(3) = GetINIMessageFromFile("Wizard", "TitleNoEmulatorsToSetup", "No Emulators to Setup!", g_udtWizard.strMessagesFile)
    lblTitle(5) = GetINIMessageFromFile("Wizard", "TitleFinished", "Finished!", g_udtWizard.strMessagesFile)
    lblTitle(6) = GetINIMessageFromFile("Wizard", "TitleWelcome", "Emulator Setup Wizard", g_udtWizard.strMessagesFile)
    m_strSelectEmulatorFile = GetINIMessageFromFile("Wizard", "SelectEmulatorFile", "Select Emulator File", g_udtWizard.strMessagesFile) & "..."
    lblWeb(0) = g_udtWizard.strHomePage
    lblWeb(1) = lblWeb(0)

    'read in the last used search path
    m_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Wizard", "SearchPath", strSearchPath, True, ""
    txtSearch = Left$(strSearchPath, 255)

    'read in the last file path
    m_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Wizard", "EmulatorFilePath", m_strLastFileDir, True, ""


    'fill the wizard's screen array
    With eWS
        .e_SYSTEM_SELECT = 0
        .e_EMU_SEARCH = 1
        .e_EMU_SELECT = 2
        .e_NONE_FOUND = 3
        .e_FINISHED = 5
        .e_INTRO = 6
    End With
    With ma_udtScreens(eWS.e_INTRO)
        .blnAbortEnable = True
        .blnBackEnable = False
        .blnNextEnable = True
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strStartButton
    End With
    With ma_udtScreens(eWS.e_SYSTEM_SELECT)
        .blnAbortEnable = True
        .blnBackEnable = False
        .blnNextEnable = True
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strNextButton
    End With
    With ma_udtScreens(eWS.e_EMU_SEARCH)
        .blnAbortEnable = True
        .blnBackEnable = True
        .blnNextEnable = False
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strNextButton
    End With
    With ma_udtScreens(eWS.e_EMU_SELECT)
        .blnAbortEnable = True
        .blnBackEnable = True
        .blnNextEnable = False
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strFinishButton
    End With
    With ma_udtScreens(eWS.e_NONE_FOUND)
        .blnAbortEnable = False
        .blnBackEnable = False
        .blnNextEnable = True
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strCloseButton
    End With
    With ma_udtScreens(eWS.e_FINISHED)
        .blnAbortEnable = False
        .blnBackEnable = False
        .blnNextEnable = True
        .strAbortButtonText = m_strAbortButton
        .strBackButtonText = m_strBackButton
        .strNextButtonText = m_strCloseButton
    End With

    'fill the systems box
    Call FillSystems
    lstSystem.ListIndex = 0

    'show the first wizard screen!
    m_lngCurrentScreen = eWS.e_INTRO
    Call ShowScreen(m_lngCurrentScreen)
    m_blnSearching = False
    m_lngNextSaveEmuIdx = 1
    m_blnByCode = False

    'assume aborted unless Finish clicked
    g_udtWizard.blnAborted = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If m_blnSearching = True Then
        Cancel = True
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set m_xobjFunc = Nothing
    Set m_xobjRegistry = Nothing

End Sub


Private Sub lblWeb_Click(Index As Integer)


    Dim strURL As String

    On Error GoTo Err_Handler

    strURL = lblWeb(Index).Caption
    If Len(strURL) > 0 Then
        If InStr(strURL, "@") > 0 Then
            If Left$(UCase$(strURL), 7) <> "MAILTO:" Then
                strURL = "mailto:" & strURL
            End If
        End If
        Call m_xobjFunc.ShellExecuteFile(strURL, "", False)
    End If

Err_Handler:

End Sub


Private Sub lvSelect_DblClick()

    Call GetFilenameForEmulator

End Sub

Private Sub lvSelect_ItemCheck(ByVal item As MSComctlLib.ListItem)

    Dim LItem As ListItem

    If m_blnByCode Then Exit Sub
    
    With item
        If .Checked = True Then
            If .SubItems(1) = m_strNotFound Then
                item.Checked = False
            Else
                item.Bold = True
                cmdNext.Enabled = True
                cmdNext.Visible = True
            End If
        Else
            item.Bold = False
            cmdNext.Enabled = False
            cmdNext.Visible = False
            For Each LItem In lvSelect.ListItems
                If LItem.Checked Then
                    cmdNext.Enabled = True
                    cmdNext.Visible = True
                    Exit For
                End If
            Next LItem
        End If
    End With

End Sub

Private Sub lvSelect_ItemClick(ByVal item As MSComctlLib.ListItem)

    If Not (lvSelect.SelectedItem Is Nothing) Then
        If Len(lvSelect.SelectedItem.Tag) > 0 Then
            txtInfo = lvSelect.SelectedItem.Tag
        Else
            txtInfo = ""
        End If
        cmdSelectFile.Enabled = True
        If lvSelect.SelectedItem.SubItems(1) <> m_strNotFound Then
            cmdSelectFile.Caption = m_strChangeFile
        Else
            cmdSelectFile.Caption = m_strSelectFile
        End If
        cmdWebSite.Enabled = (Len(g_udtWizard.a_udtSections(lstSystem.ItemData(lstSystem.ListIndex)).a_udtEmulators(Val(Right$(lvSelect.SelectedItem.Key, Len(lvSelect.SelectedItem.Key) - 1))).strURL) > 0)
        
    Else
        cmdWebSite.Enabled = False
        cmdSelectFile.Enabled = False
    End If

End Sub


Private Sub txtSearch_Change()

    cmdNext.Enabled = (Len(txtSearch) > 0)
    cmdNext.Visible = cmdNext.Enabled

End Sub

Private Sub txtSearch_GotFocus()

    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


