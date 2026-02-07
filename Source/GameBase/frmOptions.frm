VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frontend Options"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabOptions 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frmOptions.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGBGamePrompt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblGBGamePath"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraStartup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkBoss"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkTips"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraVersionCheck"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkCloseSidPlayers"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdBrowseGBGame"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&Recent Menu"
      TabPicture(1)   =   "frmOptions.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRecentUpdate"
      Tab(1).Control(1)=   "fraRecentStore"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdBrowseGBGame 
         Caption         =   "&Browse"
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Frame fraRecentStore 
         Caption         =   "Store Items"
         Height          =   1575
         Left            =   -74640
         TabIndex        =   24
         Top             =   600
         Width           =   4695
         Begin VB.OptionButton optRecentStore 
            Caption         =   "Games onl&y"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   4215
         End
         Begin VB.OptionButton optRecentStore 
            Caption         =   "Tunes &from the Music table only"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   4215
         End
         Begin VB.OptionButton optRecentStore 
            Caption         =   "Games a&nd tunes from both tables"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   4215
         End
      End
      Begin VB.Frame fraRecentUpdate 
         Caption         =   "Update"
         Height          =   1575
         Left            =   -74640
         TabIndex        =   20
         Top             =   2400
         Width           =   4695
         Begin VB.OptionButton optRecentUpdate 
            Caption         =   "&When an item is clicked/selected"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   4215
         End
         Begin VB.OptionButton optRecentUpdate 
            Caption         =   "When &an item is played/run"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   4215
         End
         Begin VB.CheckBox chkRecent 
            Caption         =   "Update when playing a tune that is lin&ked to a game"
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   1080
            Width           =   3975
         End
      End
      Begin VB.CheckBox chkCloseSidPlayers 
         Caption         =   "Sh&utdown music players on e&xit"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   3840
         Width           =   4695
      End
      Begin VB.Frame fraVersionCheck 
         Caption         =   "Version Check"
         Height          =   2175
         Left            =   2880
         TabIndex        =   6
         Top             =   540
         Width           =   2175
         Begin VB.TextBox txtProxy 
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   11
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdUsageStats 
            Caption         =   "?"
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox chkUsageStats 
            Caption         =   "Send anonymous usage statistics"
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chkNewVersion 
            Caption         =   "Check for new version on startup"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HTTP Pro&xy Host[:port]:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkTips 
         Caption         =   "Reset Tips (s&how all)"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   4560
         Width           =   4695
      End
      Begin VB.CheckBox chkBoss 
         Caption         =   "Pressing ESC in main window &Quits GameBase"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Frame fraStartup 
         Caption         =   "Startup"
         Height          =   2175
         Left            =   360
         TabIndex        =   1
         Top             =   540
         Width           =   2295
         Begin VB.CheckBox chkSplash 
            Caption         =   "&Show splash screen"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   1680
            Width           =   1935
         End
         Begin VB.ComboBox lstSelector 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkCalcTotals 
            Caption         =   "Calculate &view totals at startup"
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GB-Selector &view type:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.Label lblGBGamePath 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   360
         MouseIcon       =   "frmOptions.frx":0902
         TabIndex        =   13
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblGBGamePrompt 
         Caption         =   "'GBGame' temporary game-file &extraction folder:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3000
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
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


Private Sub SaveOptions()

    '-------------------------------------------------
    'Saves all the edited options from memory/controls
    '-------------------------------------------------

    Dim blnWritten As Boolean

    g_blnOptionsReClickGame = True

    '===========
    'general tab
    '===========

    'startup
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "CalcTotals", chkCalcTotals, False
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "SelectorViewType", lstSelector.ListIndex, False
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "ShowSplash", chkSplash, False

    blnWritten = True
    blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting("Startup", "VersionCheck", chkNewVersion.Value, g_strGameBaseDotINI)
    blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting("Startup", "SendUsageStats", chkUsageStats.Value, g_strGameBaseDotINI)
    blnWritten = blnWritten And g_xobjFunc.INI_WriteSetting("Startup", "ProxyHostPort", txtProxy.Text, g_strGameBaseDotINI)
    If Not blnWritten Then
        MsgBox "Could not save the Version Check details to GameBase.ini!", vbExclamation Or vbOKOnly, gc_strAPPNAME
    End If

    'gbgame folder
    If lblGBGamePath.Caption <> g_udtPathGBGame.strRealPath Then
        If Not g_xobjFunc.INI_WriteSetting("Paths", "GBGame", lblGBGamePath.Caption, g_strGameBaseDotINI) Then
            MsgBox "Could not save the new GBGame path to GameBase.ini!", vbExclamation Or vbOKOnly, gc_strAPPNAME
        Else
            Call GetPathInfo(lblGBGamePath.Caption, g_udtPathGBGame)
        End If
    End If

    'show tips
    If chkTips.Value = vbChecked Then
        Call ResetTips(True)
    End If

    'boss key
    g_blnBossKeyOn = (chkBoss = 1)
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "BossKeyOn", g_blnBossKeyOn, False

    'close sidplayers on exit
    g_blnShutDownSidPlayersOnExit = (chkCloseSidPlayers = 1)
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "CloseMusicPlayersOnExit", g_blnShutDownSidPlayersOnExit, False

    '===============
    'recent menu tab
    '===============

    If optRecentStore(gc_lngRECENT_GAMES).Value = True Then
        g_lngRecentStore = gc_lngRECENT_GAMES
    ElseIf optRecentStore(gc_lngRECENT_SIDS).Value = True Then
        g_lngRecentStore = gc_lngRECENT_SIDS
    Else
        g_lngRecentStore = gc_lngRECENT_BOTH
    End If
    If optRecentUpdate(gc_lngRECENT_SELECTED).Value = True Then
        g_lngRecentUpdate = gc_lngRECENT_SELECTED
    Else
        g_lngRecentUpdate = gc_lngRECENT_PLAYED
    End If
    g_blnRecentGameSid = (chkRecent.Value = vbChecked)
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "RecentStore", g_lngRecentStore, False
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "RecentUpdate", g_lngRecentUpdate, False
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Options", "RecentUpdateGameMusic", g_blnRecentGameSid, False

End Sub

Private Sub chkNewVersion_Click()

    chkUsageStats.Enabled = (chkNewVersion.Value = vbChecked)
    txtProxy.Enabled = (chkNewVersion.Value = vbChecked)
    txtProxy.BackColor = IIf((txtProxy.Enabled), lstSelector.BackColor, BackColor)

End Sub

Private Sub cmdBrowseGBGame_Click()

    'get an initial valid path to select
    If g_xobjFunc.DirExists(lblGBGamePath) Then
        g_strBrowseForFolder = lblGBGamePath
    ElseIf g_xobjFunc.DirExists(GetDefaultGBGameFolder()) Then
        g_strBrowseForFolder = GetDefaultGBGameFolder()
    ElseIf CreateDefaultGBGameFolder() = True Then
        g_strBrowseForFolder = GetDefaultGBGameFolder()
    Else
        g_strBrowseForFolder = App.Path
    End If
    'get the new path
    frmBrowseFolder.Show vbModal, Me

    'change the folder in the listbox if path is selected and save the last traversed path
    If g_strBrowseForFolder <> "" Then
        lblGBGamePath = g_strBrowseForFolder
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me

End Sub


Private Sub cmdOK_Click()
    
    'save the changed options and close window
    Call SaveOptions
    m_blnOKPressed = True
    Unload Me

End Sub

Private Sub cmdUsageStats_Click()

    MsgBox "This will send anonymous usage data to the GameBase developer" & vbCrLf & "to help improve the frontend in future versions." & vbCrLf & "This data ONLY consists of: -" & vbCrLf & vbCrLf & "- The databases you are using." & vbCrLf & vbCrLf & "- The emulators you are using.", vbInformation Or vbOKOnly, gc_strAPPNAME

End Sub

Private Sub Form_Load()

    Dim lngIdx As Long                          'loop index
    Const DesignScaleHeight As Single = 5895    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5655     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        'with shortcuts
        tabOptions.TabCaption(0) = .GetINIMessage("Options Window", "Tab_Title_General", "&General")
        tabOptions.TabCaption(1) = .GetINIMessage("Options Window", "Tab_Title_RecentMenu", "&Recent Menu")
        chkCalcTotals.Caption = .GetINIMessage("Options Window", "CalculateViewTotals", "Calculate &view totals at startup")
        chkTips.Caption = .GetINIMessage("Options Window", "ResetTips", "Reset Tips (s&how all)")
        chkBoss.Caption = .GetINIMessage("Options Window", "BossKey", "Pressing ESC in main window &Quits GameBase")
        chkCloseSidPlayers.Caption = .GetINIMessage("Options Window", "ShutdownMusicplayers", "Shutdown music players on e&xit")
        optRecentStore(0).Caption = .GetINIMessage("Options Window", "StoreGamesOnly", "Games onl&y")
        optRecentStore(1).Caption = .GetINIMessage("Options Window", "StoreMusicOnly", "Tunes &from the Music table only")
        optRecentStore(2).Caption = .GetINIMessage("Options Window", "StoreBoth", "Games a&nd tunes from both tables")
        optRecentUpdate(0).Caption = .GetINIMessage("Options Window", "UpdateSelected", "&When an item is clicked/selected")
        optRecentUpdate(1).Caption = .GetINIMessage("Options Window", "UpdatePlayed", "When &an item is played/run")
        chkRecent.Caption = .GetINIMessage("Options Window", "UpdateGameMusic", "Update when playing a tune that is lin&ked to a game")
        lblPrompt(3) = .GetINIMessage("Options Window", "GBSelectorViewType", "G&B-Selector view type") & ":"
        chkSplash.Caption = .GetINIMessage("Options Window", "ShowSplash", "&Show splash screen")
        chkNewVersion.Caption = .GetINIMessage("Options Window", "CheckForNewVersion", "Check for new version at startup")
        chkUsageStats.Caption = .GetINIMessage("Options Window", "SendUsageStats", "Send anonymous usage statistics")
        lblPrompt(0).Caption = .GetINIMessage("Options Window", "Proxy", "Pro&xy Host[:port]") & ":"
        lblGBGamePrompt.Caption = .GetINIMessage("Options Window", "GBGamePrompt", "'GBGame' temporary game-file &extraction folder") & ":"
        cmdBrowseGBGame.Caption = .GetINIMessage("Options Window", "GBGameBrowse", "&Browse")

        'without shortcuts
        Caption = .GetINIMessage("Options Window", "Caption", "Frontend Options")
        fraStartup.Caption = .GetINIMessage("Options Window", "Startup", "Startup")
        fraVersionCheck.Caption = .GetINIMessage("Options Window", "VersionCheck", "Version Check")
        lstSelector.AddItem "Icons"
        lstSelector.AddItem "Snug List"
        lstSelector.AddItem "Full Width List"
    End With

    'load the options
    Call LoadOptions
    
    'nothing changed yet
    g_blnOptionsReClickGame = False
    g_blnOptionsReLoadList = False
    m_blnOKPressed = False

End Sub

Private Sub LoadOptions()

    '------------------------------------------------------
    'Loads all the options into memory/controls for editing
    '------------------------------------------------------

    Dim lngCalcTotals As Long       'calculate view totals on startup
    Dim lngViewType As Long
    Dim lngSplash As Long
    Dim lngVerCheck As Long
    Dim lngSendUsage As Long

    '===========
    'general tab
    '===========

    'startup
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "CalcTotals", lngCalcTotals, False, 0
    chkCalcTotals = lngCalcTotals

    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "SelectorViewType", lngViewType, False, 0
    If lngViewType < 0 Or lngViewType > 2 Then
        lngViewType = 0
    End If
    lstSelector.ListIndex = lngViewType

    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "ShowSplash", lngSplash, False, 1
    chkSplash = lngSplash

    'version check
    lngVerCheck = Val(g_xobjFunc.INI_GetSetting("Startup", "VersionCheck", "0", g_strGameBaseDotINI))
    lngSendUsage = Val(g_xobjFunc.INI_GetSetting("Startup", "SendUsageStats", "0", g_strGameBaseDotINI))
    chkNewVersion.Value = IIf((lngVerCheck = 1), vbChecked, vbUnchecked)
    chkUsageStats.Value = IIf((lngSendUsage = 1), vbChecked, vbUnchecked)
    chkUsageStats.Enabled = (chkNewVersion.Value = vbChecked)
    txtProxy.Text = g_xobjFunc.INI_GetSetting("Startup", "ProxyHostPort", "", g_strGameBaseDotINI)
    txtProxy.Enabled = (chkNewVersion.Value = vbChecked)
    txtProxy.BackColor = IIf((txtProxy.Enabled), lstSelector.BackColor, BackColor)

    'gbgame path
    lblGBGamePath.Caption = g_udtPathGBGame.strRealPath

    'boss key
    chkBoss.Value = Abs(g_blnBossKeyOn)

    'close sidplayers on exit
    chkCloseSidPlayers = Abs(g_blnShutDownSidPlayersOnExit)

    '===============
    'recent menu tab
    '===============

    optRecentStore(g_lngRecentStore).Value = True
    optRecentUpdate(g_lngRecentUpdate).Value = True
    chkRecent.Value = Abs(g_blnRecentGameSid)

End Sub

Private Sub optRecentStore_Click(Index As Integer)

    If Index = gc_lngRECENT_SIDS Then
        chkRecent.Enabled = False
    Else
        chkRecent.Enabled = optRecentUpdate(gc_lngRECENT_PLAYED).Value = True
    End If

End Sub

Private Sub optRecentUpdate_Click(Index As Integer)

    If optRecentStore(gc_lngRECENT_SIDS).Value <> True Then
        chkRecent.Enabled = (Index = gc_lngRECENT_PLAYED)
    End If

End Sub


