VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemoveItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Item from Database"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmRemoveItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel (don't delete anything!)"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Remove (and delete selected files)"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4095
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   12090
      EndProperty
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Associated Files:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmRemoveItem"
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

Private m_strGameFile As String
Private m_strMusicFile As String
Private m_strPictureFile As String
Private m_strExtraFile As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim LItem As ListItem
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    lvFiles.Enabled = False
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents

    For Each LItem In lvFiles.ListItems
        If LItem.Selected = True Then
            Call g_xobjFunc.DeletePathSilent(LItem.SubItems(1))
        End If
    Next LItem

    g_blnRemoveItemConfirmed = True
    Screen.MousePointer = vbDefault
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3735    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 8655     'The required scalewidth for the form
    Dim strName As String
    Dim lngIdx As Long
    Dim LItem As ListItem
    Dim blnFound As Boolean
    Dim lngSSIdx As Long
    Dim strBase As String
    Dim strExt As String
    Dim lngDotPos As Long
    Dim strFileToFind As String
    Dim rsExtras As DAO.Recordset
    Dim FilAttr As VbFileAttribute

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'game or music?
    If g_eWhatIsSelected = e_SelType_IsGame Then
        strName = g_udtGame.strName
    Else
        strName = g_udtSid.strName
    End If

    'set the ini text
    With g_xobjFunc
        'no shortcuts
        Caption = .GetINIMessage("Remove Item Window", "Caption", "Confirm removal of") & " """ & strName & """"
        lvFiles.ColumnHeaders(1).Text = .GetINIMessage("Remove Item Window", "Type", "Type")
        lvFiles.ColumnHeaders(2).Text = .GetINIMessage("Remove Item Window", "Filename", "Filename")
        m_strGameFile = .GetINIMessage("Remove Item Window", "GameFile", "Game")
        m_strMusicFile = .GetINIMessage("Remove Item Window", "MusicFile", "Music")
        m_strPictureFile = .GetINIMessage("Remove Item Window", "PictureFile", "Picture")
        m_strExtraFile = .GetINIMessage("Remove Item Window", "ExtraFile", "Extra")
        'shortcuts
        lblPrompt(0).Caption = .GetINIMessage("Remove Item Window", "AssociatedFilesPrompt", "&Select any associated files you wish to delete") & ":"
        cmdOK.Caption = .GetINIMessage("Remove Item Window", "RemoveButton", "&Remove  (and delete selected files)")
        cmdCancel.Caption = .GetINIMessage("Remove Item Window", "CancelButton", "&Cancel  (don't delete anything!)")
    End With

    'add game files to the list
    If g_eWhatIsSelected = e_SelType_IsGame Then
        If g_udtGame.strGameFilename <> "" And g_udtGame.blnGameFileExists = True Then
            If g_lngTotalGamePaths > 0 Then
                For lngIdx = 0 To (g_lngTotalGamePaths - 1)
                    With ga_udtRegPathGame(lngIdx)
                        If .ePathType = e_PathLocation_FixedHD Or .ePathType = e_PathLocation_Removable Then
                            If g_xobjFunc.FileExists(.strWithBackSlash & g_udtGame.strGameFilename, False, "") = True Then
                                FilAttr = GetAttr(.strWithBackSlash & g_udtGame.strGameFilename)
                                If (Not (FilAttr And vbDirectory)) And (Not (FilAttr And vbSystem)) Then
                                    Set LItem = lvFiles.ListItems.Add(, , m_strGameFile)
                                    LItem.SubItems(1) = .strWithBackSlash & g_udtGame.strGameFilename
                                    Set LItem = Nothing
                                End If
                                Exit For
                            End If
                        End If
                    End With
                Next lngIdx
            End If
        End If
    End If

    'add music file to list
    If g_udtSid.blnFileExists = True Then
        If g_udtRegPathSid.ePathType = e_PathLocation_FixedHD Or g_udtRegPathSid.ePathType = e_PathLocation_Removable Then
            If g_udtRegPathSid.strWithBackSlash <> "" Then
                If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, False, "") = True Then
                    FilAttr = GetAttr(g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename)
                    If (Not (FilAttr And vbDirectory)) And (Not (FilAttr And vbSystem)) Then
                        Set LItem = lvFiles.ListItems.Add(, , m_strMusicFile)
                        LItem.SubItems(1) = g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename
                        Set LItem = Nothing
                    End If
                End If
            End If
        End If
    End If

    'add pictures to list
    If g_eWhatIsSelected = e_SelType_IsGame Then
        If g_udtGame.strScreenshotFilename <> "" Then
            If g_lngTotalScrnShotPaths > 0 Then
                If g_blnMamePicStyle = False Then
                    'GAMEBASE STYLE PICTURES
                    strBase = g_udtGame.strScreenshotFilename
                    lngDotPos = InStr(strBase, ".")
                    If lngDotPos > 1 And lngDotPos < Len(strBase) Then
                        strExt = Right$(strBase, Len(strBase) - (lngDotPos - 1))
                        strBase = Left$(strBase, Len(strBase) - Len(strExt))
                        lngSSIdx = 0
                        Do
                            If lngSSIdx = 0 Then
                                strFileToFind = strBase & strExt
                            Else
                                strFileToFind = strBase & "_" & Trim$(lngSSIdx) & strExt
                            End If
                            blnFound = False
                            For lngIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                                With ga_udtRegPathScrnShot(lngIdx)
                                    If .ePathType = e_PathLocation_FixedHD Or .ePathType = e_PathLocation_Removable Then
                                        If g_xobjFunc.FileExists(.strWithBackSlash & strFileToFind, False, "") = True Then
                                            FilAttr = GetAttr(.strWithBackSlash & strFileToFind)
                                            If (Not (FilAttr And vbDirectory)) And (Not (FilAttr And vbSystem)) Then
                                                Set LItem = lvFiles.ListItems.Add(, , m_strPictureFile)
                                                LItem.SubItems(1) = .strWithBackSlash & strFileToFind
                                                Set LItem = Nothing
                                            End If
                                            blnFound = True
                                            Exit For
                                        End If
                                    End If
                                End With
                            Next lngIdx
                            lngSSIdx = lngSSIdx + 1
                        Loop Until blnFound = False
                    End If
                Else
                    'MAME STYLE PICTURES
                    For lngIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                        With ga_udtRegPathScrnShot(lngIdx)
                            If .ePathType = e_PathLocation_FixedHD Or .ePathType = e_PathLocation_Removable Then
                                If g_xobjFunc.FileExists(.strWithBackSlash & g_udtGame.strScreenshotFilename, False, "") = True Then
                                    FilAttr = GetAttr(.strWithBackSlash & g_udtGame.strScreenshotFilename)
                                    If (Not (FilAttr And vbDirectory)) And (Not (FilAttr And vbSystem)) Then
                                        Set LItem = lvFiles.ListItems.Add(, , m_strPictureFile)
                                        LItem.SubItems(1) = .strWithBackSlash & g_udtGame.strScreenshotFilename
                                        Set LItem = Nothing
                                    End If
                                End If
                            End If
                        End With
                    Next lngIdx
                End If
            End If
        End If
    End If

    'add extras to list
    If g_eWhatIsSelected = e_SelType_IsGame Then
        If g_udtGame.blnHasExtras = True Then
            If g_lngTotalExtrasPaths > 0 Then
                Set rsExtras = g_dbGamebase.OpenRecordset("SELECT Path FROM Extras WHERE GA_Id = " & g_udtGame.lngID, dbOpenSnapshot)
                With rsExtras
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            For lngIdx = 0 To (g_lngTotalExtrasPaths - 1)
                                If ga_udtRegPathExtras(lngIdx).ePathType = e_PathLocation_FixedHD Or ga_udtRegPathExtras(lngIdx).ePathType = e_PathLocation_Removable Then
                                    If g_xobjFunc.FileExists(ga_udtRegPathExtras(lngIdx).strWithBackSlash & !Path, False, "") = True Then
                                        FilAttr = GetAttr(ga_udtRegPathExtras(lngIdx).strWithBackSlash & !Path)
                                        If (Not (FilAttr And vbDirectory)) And (Not (FilAttr And vbSystem)) Then
                                            Set LItem = lvFiles.ListItems.Add(, , m_strExtraFile)
                                            LItem.SubItems(1) = ga_udtRegPathExtras(lngIdx).strWithBackSlash & !Path
                                            Set LItem = Nothing
                                        End If
                                    End If
                                End If
                            Next lngIdx
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                Set rsExtras = Nothing
            End If
        End If
    End If

    'select nothing
    If lvFiles.ListItems.Count > 0 Then
        Set lvFiles.SelectedItem = Nothing
        lvFiles.ListItems(1).Selected = False
    End If

    'assume cancel is pressed
    g_blnRemoveItemConfirmed = False

End Sub

