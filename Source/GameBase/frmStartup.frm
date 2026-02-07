VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartup 
   Caption         =   " GameBase"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ImageList iLstGameBase 
      Left            =   600
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvGameBase 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "iLstGameBase"
      SmallIcons      =   "iLstGameBase"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmStartup"
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

Public m_blnUnload As Boolean
Private Sub SelectGameBase()

    Dim LItem As ListItem

    Set LItem = lvGameBase.SelectedItem
    If LItem Is Nothing Then
        Exit Sub
    Else
        With LItem
            g_strGameBasePath = g_strAppPathWithBackslash & .Key
            g_strGameBasePathWBS = g_strAppPathWithBackslash & .Key & "\"
            g_strGameBaseName = .Key
            g_strMessagesFile = g_strAppPathWithBackslash & .Key & "\Text.ini"
            g_strGameBaseDBPath = g_strAppPathWithBackslash & .Key & "\" & .Key & ".mdb"
            g_strEmuFile = g_strAppPathWithBackslash & .Key & "\Emulators.ini"
            g_strMusicFile = g_strAppPathWithBackslash & .Key & "\MusicPlayers.ini"
            g_strPathFile = g_strAppPathWithBackslash & .Key & "\Paths.ini"
            g_strSettingsFile = g_strAppPathWithBackslash & .Key & "\Settings.ini"
            g_strGameBaseDotINI = g_strAppPathWithBackslash & "GameBase.ini"
        End With
        Set LItem = Nothing
        g_blnAbort = False
        Unload Me
    End If

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
Private Sub Form_Activate()

    Dim strGameBase As String
    Dim LItem As ListItem
    Dim blnFound As Boolean
    Dim lngTimeUp&

    'get settings
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Startup", "GameBase", strGameBase, True, ""

    'set the listitem
    If strGameBase = "" Then
        Set lvGameBase.SelectedItem = lvGameBase.ListItems(1)
        lvGameBase.ListItems(1).Selected = True
    Else
        blnFound = False
        For Each LItem In lvGameBase.ListItems
            If UCase$(LItem.Key) = UCase$(strGameBase) Then
                Set lvGameBase.SelectedItem = LItem
                LItem.Selected = True
                blnFound = True
            End If
        Next LItem
        If Not blnFound Then
            Set lvGameBase.SelectedItem = lvGameBase.ListItems(1)
            lvGameBase.ListItems(1).Selected = True
        End If
    End If

    g_strGameBasePath = ""
    g_strGameBasePathWBS = ""
    g_strGameBaseName = ""
    g_strMessagesFile = ""
    g_strGameBaseDBPath = ""
    g_strEmuFile = ""
    g_strMusicFile = ""
    g_strPathFile = ""
    g_strSettingsFile = ""
    g_strGameBaseDotINI = ""
    g_blnAbort = True

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3525    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4875     'The required scalewidth for the form
    Dim sngHeight As Single
    Dim sngWidth As Single
    Dim strDir As String
    Dim LItem As ListItem
    Dim lngIdx As Long
    Dim lngAttr As VbFileAttribute
    Dim lngViewType As Long
    Dim strDBPathFile As String

    m_blnUnload = False
    g_blnAbort = True
    
    Caption = " " & gc_strAPPNAME

    'set the form's proper width and height
    With g_xobjRegistry
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Startup", "Width", sngWidth, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Startup", "Height", sngHeight, False, -1
    End With
    If (sngWidth > 32) And (sngWidth < Screen.Width) And (sngHeight > 32) And (sngHeight < Screen.Height) Then
        Height = sngHeight
        Width = sngWidth
    Else
        Height = DesignScaleHeight + (Height - ScaleHeight)
        Width = DesignScaleWidth + (Width - ScaleWidth)
    End If

    'set the listview style
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "SelectorViewType", lngViewType, False, 0
    If lngViewType < 0 Or lngViewType > 2 Then
        lngViewType = 0
    End If
    If lngViewType = 0 Then
        lvGameBase.View = 0
    ElseIf lngViewType = 1 Then
        lvGameBase.ColumnHeaders(1).Width = lvGameBase.Width - 380
        lvGameBase.View = 2
    Else
        lvGameBase.ColumnHeaders(1).Width = lvGameBase.Width - 380
        lvGameBase.View = 3
        lvGameBase.FullRowSelect = True
    End If

    'load the icons and custom database name
    lngIdx = 2
    strDir = Dir(g_strAppPathWithBackslash, vbDirectory)
    Do Until strDir = ""
        lngAttr = GetAttr(g_strAppPathWithBackslash & strDir)
        If lngAttr And vbDirectory Then
            If (strDir <> ".") And (strDir <> "..") And (UCase$(strDir) <> "DOCS") Then
                strDBPathFile = g_strAppPathWithBackslash & strDir & "\" & strDir & ".mdb"
                If g_xobjFunc.FileExists(strDBPathFile, False, gc_strAPPNAME) = True Then
                    Set LItem = lvGameBase.ListItems.Add(, strDir, GetCustomDatabaseName(strDir))
                    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & strDir & "\Gfx\Startup_32x32.gif", False, gc_strAPPNAME) = False Then
                        LItem.Icon = 1
                        LItem.SmallIcon = 1
                    Else
                        iLstGameBase.ListImages.Add lngIdx, , LoadPicture(g_strAppPathWithBackslash & strDir & "\Gfx\Startup_32x32.gif")
                        LItem.Icon = lngIdx
                        LItem.SmallIcon = lngIdx
                        lngIdx = lngIdx + 1
                    End If
                    Set LItem = Nothing
                End If
            End If
        End If
        strDir = Dir()
    Loop

    With lvGameBase
        If .ListItems.Count = 0 Then
            m_blnUnload = True
        ElseIf .ListItems.Count = 1 Then
            g_strGameBasePath = g_strAppPathWithBackslash & .ListItems(1).Key
            g_strGameBasePathWBS = g_strAppPathWithBackslash & .ListItems(1).Key & "\"
            g_strGameBaseName = .ListItems(1).Key
            g_strMessagesFile = g_strAppPathWithBackslash & .ListItems(1).Key & "\Text.ini"
            g_strGameBaseDBPath = g_strAppPathWithBackslash & .ListItems(1).Key & "\" & .ListItems(1).Key & ".mdb"
            g_strEmuFile = g_strAppPathWithBackslash & .ListItems(1).Key & "\Emulators.ini"
            g_strMusicFile = g_strAppPathWithBackslash & .ListItems(1).Key & "\MusicPlayers.ini"
            g_strPathFile = g_strAppPathWithBackslash & .ListItems(1).Key & "\Paths.ini"
            g_strSettingsFile = g_strAppPathWithBackslash & .ListItems(1).Key & "\Settings.ini"
            g_strGameBaseDotINI = g_strAppPathWithBackslash & "GameBase.ini"
            g_blnAbort = False
        Else
            g_strGameBasePath = ""
            g_strGameBasePathWBS = ""
            g_strGameBaseName = ""
            g_strMessagesFile = ""
            g_strGameBaseDBPath = ""
            g_strEmuFile = ""
            g_strMusicFile = ""
            g_strPathFile = ""
            g_strSettingsFile = ""
            g_strGameBaseDotINI = ""
        End If
    End With

End Sub

Private Sub Form_Resize()

    If ScaleWidth > 32 Then
        lvGameBase.Width = ScaleWidth
        lvGameBase.ColumnHeaders(1).Width = lvGameBase.Width - 380
    End If
    If ScaleHeight > 32 Then
        lvGameBase.Height = ScaleHeight
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    With g_xobjRegistry
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Startup", "Width", Width, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Startup", "Height", Height, False
        If lvGameBase.ListItems.Count > 1 Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Startup", "GameBase", lvGameBase.SelectedItem.Key, True
        End If
    End With

End Sub


Private Sub lvGameBase_DblClick()

    'do stuff
    Call SelectGameBase
    
End Sub

Private Sub lvGameBase_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32, 13
            Call SelectGameBase
        Case 27
            g_blnAbort = True
            Unload Me
    End Select

End Sub

