VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPaths 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paths"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmPaths.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabPaths 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Games"
      TabPicture(0)   =   "frmPaths.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ctlGamePaths"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Music"
      TabPicture(1)   =   "frmPaths.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMusicPath"
      Tab(1).Control(1)=   "lblPrompt(0)"
      Tab(1).Control(2)=   "cmdMusicClear"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Pictures"
      TabPicture(2)   =   "frmPaths.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ctlScrnShotPaths"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Extras"
      TabPicture(3)   =   "frmPaths.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ctlExtrasPaths"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "P&hotos"
      TabPicture(4)   =   "frmPaths.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblPhotoPath"
      Tab(4).Control(1)=   "lblPrompt(1)"
      Tab(4).Control(2)=   "cmdPhotosClear"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton cmdPhotosClear 
         Caption         =   "Clea&r"
         Height          =   285
         Left            =   -67560
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdMusicClear 
         Caption         =   "Cle&ar"
         Height          =   285
         Left            =   -67560
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin GameBase.ctlPathSelector ctlGamePaths 
         Height          =   2340
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4128
      End
      Begin GameBase.ctlPathSelector ctlScrnShotPaths 
         Height          =   2340
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4128
      End
      Begin GameBase.ctlPathSelector ctlExtrasPaths 
         Height          =   2340
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4128
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Photos Root Folder:"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   7
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music Root Folder:"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   2
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPhotoPath 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -74760
         MouseIcon       =   "frmPaths.frx":0956
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   8175
      End
      Begin VB.Label lblMusicPath 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -74760
         MouseIcon       =   "frmPaths.frx":0C60
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1200
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPaths"
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

Private m_strLastMusicPath As String      'last sid path traversed
Private m_strLastPhotoPath As String    'last photo path traversed
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMusicClear_Click()

    lblMusicPath = ""
    tabPaths.SetFocus

End Sub

Private Sub cmdOK_Click()

    Dim lngIdx As Long      'loop index

    'save game paths
    With ctlGamePaths
        If g_xobjFunc.INI_DeleteNumberedValues("Games", .ListCount + 1, g_strPathFile) = True Then
            g_lngTotalGamePaths = .ListCount
            If g_lngTotalGamePaths > 0 Then
                ReDim ga_udtRegPathGame(g_lngTotalGamePaths - 1)
            End If
            For lngIdx = 0 To (g_lngTotalGamePaths - 1)
                Call GetPathInfo(.List(lngIdx), ga_udtRegPathGame(lngIdx))
                If g_xobjFunc.INI_WriteSetting("Games", Trim$(lngIdx + 1), ga_udtRegPathGame(lngIdx).strRealPath, g_strPathFile) = False Then
                    MsgBox "Could not save your Games paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
                    Exit For
                End If
            Next lngIdx
        Else
            MsgBox "Could not save your Games paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
        End If
    End With
    g_blnValidGamePathExists = False
    For lngIdx = 0 To (g_lngTotalGamePaths - 1)
        If ga_udtRegPathGame(lngIdx).ePathType <> e_PathLocation_Invalid Then
            g_blnValidGamePathExists = True
            Exit For
        End If
    Next lngIdx

    'save music path
    If lblMusicPath <> g_udtRegPathSid.strRealPath Then
        Call GetPathInfo(lblMusicPath, g_udtRegPathSid)
        If g_xobjFunc.INI_WriteSetting("Misc", "Music", g_udtRegPathSid.strRealPath, g_strPathFile) = False Then
            MsgBox "Could not save your Music path to " & g_strPathFile, vbExclamation, gc_strAPPNAME
        End If
    End If

    'save picture paths
    With ctlScrnShotPaths
        If g_xobjFunc.INI_DeleteNumberedValues("Pictures", .ListCount + 1, g_strPathFile) = True Then
            g_lngTotalScrnShotPaths = .ListCount
            If g_lngTotalScrnShotPaths > 0 Then
                ReDim ga_udtRegPathScrnShot(g_lngTotalScrnShotPaths - 1)
            End If
            For lngIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                Call GetPathInfo(.List(lngIdx), ga_udtRegPathScrnShot(lngIdx))
                If g_xobjFunc.INI_WriteSetting("Pictures", Trim$(lngIdx + 1), ga_udtRegPathScrnShot(lngIdx).strRealPath, g_strPathFile) = False Then
                    MsgBox "Could not save your Picture paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
                    Exit For
                End If
            Next lngIdx
        Else
            MsgBox "Could not save your Picture paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
        End If
    End With

    'save extras paths
    With ctlExtrasPaths
        If g_xobjFunc.INI_DeleteNumberedValues("Extras", .ListCount + 1, g_strPathFile) = True Then
            g_lngTotalExtrasPaths = .ListCount
            If g_lngTotalExtrasPaths > 0 Then
                ReDim ga_udtRegPathExtras(g_lngTotalExtrasPaths - 1)
            End If
            For lngIdx = 0 To (g_lngTotalExtrasPaths - 1)
                Call GetPathInfo(.List(lngIdx), ga_udtRegPathExtras(lngIdx))
                If g_xobjFunc.INI_WriteSetting("Extras", Trim$(lngIdx + 1), ga_udtRegPathExtras(lngIdx).strRealPath, g_strPathFile) = False Then
                    MsgBox "Could not save your Extras paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
                    Exit For
                End If
            Next lngIdx
        Else
            MsgBox "Could not save your Extras paths to " & g_strPathFile, vbExclamation, gc_strAPPNAME
        End If
    End With

    'save photos path
    If lblPhotoPath <> g_udtRegPathPhotos.strRealPath Then
        Call GetPathInfo(lblPhotoPath, g_udtRegPathPhotos)
        If g_xobjFunc.INI_WriteSetting("Misc", "Photos", g_udtRegPathPhotos.strRealPath, g_strPathFile) = False Then
            MsgBox "Could not save your Photos path to " & g_strPathFile, vbExclamation, gc_strAPPNAME
        End If
    End If

    'all done!
    g_blnOptionsReClickGame = True
    Unload Me

End Sub


Private Sub cmdPhotosClear_Click()

    lblPhotoPath = ""
    tabPaths.SetFocus

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3615    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 8895     'The required scalewidth for the form
    Dim strLastGamePath As String               'last game path traversed
    Dim strLastScrnShotPath As String           'last screenshot path traversed
    Dim strLastExtrasPath As String             'last extras path traversed
    Dim lngPathIdx As Long                      'loop index

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        'with shortcuts
        tabPaths.TabCaption(0) = .GetINIMessage("Paths Window", "GamesTab", "&Games")
        tabPaths.TabCaption(1) = .GetINIMessage("Paths Window", "MusicTab", "&Music")
        tabPaths.TabCaption(2) = .GetINIMessage("Paths Window", "PicturesTab", "&Pictures")
        tabPaths.TabCaption(3) = .GetINIMessage("Paths Window", "ExtrasTab", "&Extras")
        tabPaths.TabCaption(4) = .GetINIMessage("Paths Window", "PhotosTab", "P&hotos")
        cmdMusicClear.Caption = .GetINIMessage("Paths Window", "ClearMusicButton", "Cle&ar")
        cmdPhotosClear.Caption = .GetINIMessage("Paths Window", "ClearPhotosButton", "Clea&r")
        'without shortcuts
        Caption = .GetINIMessage("Paths Window", "Caption", "Paths")
        lblPrompt(0).Caption = .GetINIMessage("Paths Window", "MusicRootFolder", "Music Root Folder") & ":"
        lblPrompt(1).Caption = .GetINIMessage("Paths Window", "PhotosRootFolder", "Photos Root Folder") & ":"
    End With

    'load the last selected paths
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlGamePaths.BrowseRegKey, strLastGamePath, True, App.Path
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlScrnShotPaths.BrowseRegKey, strLastScrnShotPath, True, App.Path
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlExtrasPaths.BrowseRegKey, strLastExtrasPath, True, App.Path
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastMusicPath", m_strLastMusicPath, True, App.Path
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPhotoPath", m_strLastPhotoPath, True, App.Path
    'and set them in the path selector controls
    ctlGamePaths.SetLastTraversedPath strLastGamePath
    ctlScrnShotPaths.SetLastTraversedPath strLastScrnShotPath
    ctlExtrasPaths.SetLastTraversedPath strLastExtrasPath

    'load the game paths
    For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
        ctlGamePaths.AddItem ga_udtRegPathGame(lngPathIdx).strRealPath
    Next lngPathIdx

    'load the music path
    lblMusicPath = g_udtRegPathSid.strRealPath

    'load the picture paths
    For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
        ctlScrnShotPaths.AddItem ga_udtRegPathScrnShot(lngPathIdx).strRealPath
    Next lngPathIdx

    'load the extras paths
    For lngPathIdx = 0 To (g_lngTotalExtrasPaths - 1)
        ctlExtrasPaths.AddItem ga_udtRegPathExtras(lngPathIdx).strRealPath
    Next lngPathIdx

    'load the photos path
    lblPhotoPath = g_udtRegPathPhotos.strRealPath

    'nothing changed yet
    g_blnOptionsReClickGame = False
    g_blnOptionsReLoadList = False

End Sub

Private Sub lblSidPath_Click()

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'save the last selected paths for the path selectors
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlGamePaths.BrowseRegKey, ctlGamePaths.GetLastTraversedPath, True
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlScrnShotPaths.BrowseRegKey, ctlScrnShotPaths.GetLastTraversedPath, True
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", ctlExtrasPaths.BrowseRegKey, ctlExtrasPaths.GetLastTraversedPath, True
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastMusicPath", m_strLastMusicPath, True
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPhotoPath", m_strLastPhotoPath, True

End Sub


Private Sub lblMusicPath_Click()

    'get an initial valid path to select
    If g_xobjFunc.DirExists(lblMusicPath) Then
        m_strLastMusicPath = lblMusicPath
    ElseIf Not g_xobjFunc.DirExists(m_strLastMusicPath) Then
        m_strLastMusicPath = App.Path
    End If
    'get the new path
    g_strBrowseForFolder = m_strLastMusicPath
    frmBrowseFolder.Show vbModal, Me
    
    'change the folder in the listbox if path is selected and save the last traversed path
    If g_strBrowseForFolder <> "" Then
        lblMusicPath = g_strBrowseForFolder
        m_strLastMusicPath = g_strBrowseForFolder
    End If

End Sub

Private Sub lblPhotoPath_Click()

    'get an initial valid path to select
    If g_xobjFunc.DirExists(lblPhotoPath) Then
        m_strLastPhotoPath = lblPhotoPath
    ElseIf Not g_xobjFunc.DirExists(m_strLastPhotoPath) Then
        m_strLastPhotoPath = App.Path
    End If
    'get the new path
    g_strBrowseForFolder = m_strLastPhotoPath
    frmBrowseFolder.Show vbModal, Me
    
    'change the folder in the listbox if path is selected and save the last traversed path
    If g_strBrowseForFolder <> "" Then
        lblPhotoPath = g_strBrowseForFolder
        m_strLastPhotoPath = g_strBrowseForFolder
    End If

End Sub


