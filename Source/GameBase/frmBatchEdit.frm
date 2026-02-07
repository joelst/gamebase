VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Edit"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1335
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating Records..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "frmBatchEdit"
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

Private m_blnCancelled As Boolean

Public m_blnUpdateGameInfo As Boolean
Public m_blnUpdateVersionInfo As Boolean
Public m_blnUpdatePersonalInfo As Boolean
Public m_blnUpdateMusicInfo As Boolean
Public m_blnUpdateGameFile As Boolean
Public m_blnUpdateSidFile As Boolean
Public m_blnUpdatePictureFile As Boolean
Public m_blnUpdateExtras As Boolean
Public m_blnExportToTextFile As Boolean
Public m_blnUpdateKV As Boolean



Private Sub ExportToTextFile()

    Dim lngListViewIdx As Long
    Dim strList As String
    Dim intFileNum As Integer
    Dim LItem As ListItem
    Dim lngCHIdx As Long

    On Error GoTo Err_Handler
    
    lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
    Caption = g_xobjFunc.GetINIMessage("Batch Update", "ExportCaption", "Exporting") & "..."
    lblPrompt = g_xobjFunc.GetINIMessage("Batch Update", "ExportPrompt1", "Exporting") & " "
    lblPrompt = lblPrompt & frmMain.lvMain(lngListViewIdx).ListItems.Count & " "
    lblPrompt = lblPrompt & g_xobjFunc.GetINIMessage("Batch Update", "ExportPrompt2", "items to Text File") & "..."
    
    With frmMain.lvMain(lngListViewIdx)
        If .ListItems.Count > 0 Then
            pbStatus.Min = 0
            pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
            pbStatus.Value = 0
            Screen.MousePointer = vbHourglass
            DoEvents
            'delete the existing file
            If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "viewlist.txt", False, gc_strAPPNAME) = True Then
                Call g_xobjFunc.DeletePathSilent(g_strAppPathWithBackslash & "viewlist.txt")
            End If
            'create the text list
            strList = Now & vbCrLf & g_strGameBaseDBPath & vbCrLf & frmMain.lstView.List(frmMain.lstView.ListIndex) & vbCrLf & vbCrLf
            If (.View = lvwReport) And (.ColumnHeaders.Count > 1) Then
                'name and subitems
                For Each LItem In .ListItems
                    strList = strList & LItem.Text & vbTab
                    For lngCHIdx = 1 To .ColumnHeaders.Count - 1
                        strList = strList & vbTab & LItem.SubItems(lngCHIdx)
                    Next lngCHIdx
                    strList = strList & vbCrLf
                    pbStatus.Value = pbStatus.Value + 1
                    DoEvents
                    If m_blnCancelled = True Then
                        Exit For
                    End If
                Next LItem
            Else
                'name only
                For Each LItem In .ListItems
                    strList = strList & LItem.Text & vbCrLf
                    pbStatus.Value = pbStatus.Value + 1
                    DoEvents
                    If m_blnCancelled = True Then
                        Exit For
                    End If
                Next LItem
            End If
            If m_blnCancelled = False Then
                pbStatus.Value = pbStatus.Max
                DoEvents
                'write the new file
                intFileNum = FreeFile()
                Open g_strAppPathWithBackslash & "viewlist.txt" For Output As #intFileNum
                Print #intFileNum, strList
                Close #intFileNum
                'show the file
                Call g_xobjFunc.ShellExecuteFile(g_strAppPathWithBackslash & "viewlist.txt", "", False)
            End If
            Screen.MousePointer = vbDefault
        End If
    End With

    Exit Sub

Err_Handler:
    MsgBox Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Sub
Private Sub UpdateGameFile()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldID As Field
    Dim fldFA As Field
    Dim fldSA As Field
    Dim fldCCode As Field
    Dim fldFilename As Field
    Dim fldFileToRun As Field
    Dim fldIndex As Field
    Dim fldVersion As Field

    'save game file, filetorun and index to all games in the current view (calculate FA, SA and CCode)
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Filename, FilenameIndex, FileToRun, FA, CCode, SA, Version FROM Games ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldID = !GA_Id
                Set fldFA = !FA
                Set fldSA = !SA
                Set fldCCode = !CCode
                Set fldFilename = !Filename
                Set fldFileToRun = !FileToRun
                Set fldIndex = !FilenameIndex
                Set fldVersion = !Version
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            fldFA = (g_strGameSelectorFilename <> "")
                            fldFilename = g_strGameSelectorFilename
                            fldFileToRun = g_strGameSelectorDefaultFile
                            fldIndex = g_strGameSelectorIndex
                            If g_strGameSelectorFilename <> "" Then
                                If fldSA = True Then
                                     fldCCode = 1 'both exist
                                Else
                                    fldCCode = 2 'only game exists
                                End If
                            Else
                                If fldSA = True Then
                                    fldCCode = 3 'only sid exists
                                Else
                                    fldCCode = 4 'none exist
                                End If
                            End If
                            If g_blnGameSelectorApplyAllVersion = True Then
                                fldVersion = g_lngGameSelectorVersion
                            End If
                            .Update
                        End If
                    End If
                    pbStatus.Value = pbStatus.Value + 1
                    DoEvents
                    If m_blnCancelled = True Then
                        Exit For
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldID = Nothing
                Set fldFA = Nothing
                Set fldSA = Nothing
                Set fldCCode = Nothing
                Set fldFilename = Nothing
                Set fldFileToRun = Nothing
                Set fldIndex = Nothing
                Set fldVersion = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Public Function UpdateGameInfo()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldYear As Field
    Dim fldPub As Field
    Dim fldPro As Field
    Dim fldMus As Field
    Dim fldGen As Field
    Dim fldLan As Field
    Dim fldPlayFr As Field
    Dim fldPlayTo As Field
    Dim fldPlaySim As Field
    Dim fldCon As Field
    Dim fldCom As Field
    Dim fldArt As Field
    Dim fldDev As Field
    Dim fldLic As Field
    Dim fldRar As Field
    Dim fldWLN As Field
    Dim fldWLU As Field
    Dim fldRR As Field

    'save game info to all games in the current view
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, YE_Id, PU_Id, PR_Id, MU_Id, GE_Id, AR_Id, DE_Id, LI_Id, RA_Id, PlayersFrom, PlayersTo, PlayersSim, LA_Id, Control, Comment, WebLink_Name, WebLink_URL, ReviewRating FROM Games ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldYear = !YE_ID
                Set fldPub = !PU_Id
                Set fldPro = !PR_Id
                Set fldMus = !MU_Id
                Set fldGen = !GE_Id
                Set fldLan = !LA_Id
                Set fldArt = !AR_Id
                Set fldDev = !DE_Id
                Set fldLic = !LI_Id
                Set fldRar = !RA_Id
                Set fldPlayFr = !PlayersFrom
                Set fldPlayTo = !PlayersTo
                Set fldPlaySim = !PlayersSim
                Set fldCon = !Control
                Set fldCom = !Comment
                Set fldWLN = !WebLink_Name
                Set fldWLU = !WebLink_URL
                Set fldRR = !ReviewRating
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            With g_udtEditGameInfo
                                If .blnUpdateComment Then
                                    fldCom = .strComment
                                End If
                                If .blnUpdateControl Then
                                    fldCon = .lngControlID
                                End If
                                If .blnUpdateGenre Then
                                    fldGen = .lngGenreID
                                End If
                                If .blnUpdateLanguage Then
                                    fldLan = .lngLanguageID
                                End If
                                If .blnUpdateMusician Then
                                    fldMus = .lngMusicianID
                                End If
                                If .blnUpdatePlayers Then
                                    fldPlayFr = .lngPlayersFrom
                                    fldPlayTo = .lngPlayersTo
                                    fldPlaySim = .blnSimPlay
                                End If
                                If .blnUpdateProgrammer Then
                                    fldPro = .lngProgrammerID
                                End If
                                If .blnUpdatePublisher Then
                                    fldPub = .lngPublisherID
                                End If
                                If .blnUpdateYear Then
                                    fldYear = .lngYearID
                                End If
                                If .blnUpdateArtist Then
                                    fldArt = .lngArtistID
                                End If
                                If .blnUpdateDeveloper Then
                                    fldDev = .lngDeveloperID
                                End If
                                If .blnUpdateLicense Then
                                    fldLic = .lngLicenseID
                                End If
                                If .blnUpdateRarity Then
                                    fldRar = .lngRarityID
                                End If
                                If .blnUpdateWebLink Then
                                    fldWLN = .strWebLinkName
                                    fldWLU = .strWebLinkURL
                                End If
                                If .blnUpdateReviewRating Then
                                    fldRR = .lngReviewRating
                                End If
                            End With
                            .Update
                        End If
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldYear = Nothing
                Set fldPub = Nothing
                Set fldPro = Nothing
                Set fldMus = Nothing
                Set fldGen = Nothing
                Set fldLan = Nothing
                Set fldArt = Nothing
                Set fldDev = Nothing
                Set fldLic = Nothing
                Set fldRar = Nothing
                Set fldPlayFr = Nothing
                Set fldPlayTo = Nothing
                Set fldPlaySim = Nothing
                Set fldCon = Nothing
                Set fldCom = Nothing
                Set fldWLN = Nothing
                Set fldWLU = Nothing
                Set fldRR = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Function

Public Sub UpdateKV()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldGemus As Field
    Dim a_udtSrcKV() As NotesKeysType
    Dim lngTotalSrcKV As Long
    Dim a_udtDstKVLines() As NotesKeysLinesType
    Dim lngTotalDstKVLines As Long
    Dim lngIdx As Long
    Dim strGemus As String
    Dim lngSrcIdx As Long
    Dim lngDstIdx As Long
    Dim strKey As String
    Dim strValue As String
    Dim blnFound As Boolean
    Dim a_udtToAdd() As NotesKeysType
    Dim lngTotalToAdd As Long

    If g_UpdateKVTypeEnum = e_Save_Current Then Exit Sub
    
    ReDim a_udtSrcKV(0)
    lngTotalSrcKV = 0

    ReDim a_udtDstKVLines(0)
    lngTotalDstKVLines = 0

    'save game info to all games in the current view
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Gemus FROM Games ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
            
                'convert kv text into array of key=values
                ReDim a_udtSrcKV(0)
                lngTotalSrcKV = g_xobjGEMUS.GetKVsFromString(g_strUpdateKVText, a_udtSrcKV)
            
                'loop through each game in the current view
                Set fldGemus = !Gemus
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            If g_UpdateKVTypeEnum = e_Save_All Then
                            
                                'JUST OVERWRITE

                                .Edit
                                fldGemus = g_strUpdateKVText
                                .Update
                            
                            Else

                                'need to break the dest kv pairs into lines
                                ReDim a_udtDstKVLines(0)
                                lngTotalDstKVLines = g_xobjGEMUS.GetKVLinesFromString(fldGemus, a_udtDstKVLines)
                                If g_UpdateKVTypeEnum = e_Add_All_Overwrite Then
                                
                                    'ADD KEYS OVERWRITING EXISTING ONES (EXCEPT KVSETS)
                                    
                                    'get list of new keys to add
                                    ReDim a_udtToAdd(0)
                                    lngTotalToAdd = 0
                                    For lngSrcIdx = 0 To lngTotalSrcKV - 1
                                        strKey = a_udtSrcKV(lngSrcIdx).strKey
                                        strValue = a_udtSrcKV(lngSrcIdx).strValue
                                        blnFound = False
                                        For lngDstIdx = 0 To lngTotalDstKVLines - 1
                                            With a_udtDstKVLines(lngDstIdx)
                                                If .blnIsKV And (LCase$(.strKey) = LCase$(strKey)) Then
                                                    If .strKey = "kvset" Then
                                                        If .strValue = strValue Then
                                                            blnFound = True
                                                        End If
                                                    Else
                                                        'update value
                                                        .strValue = strValue
                                                        blnFound = True
                                                    End If
                                                End If
                                            End With
                                            If blnFound Then Exit For
                                        Next lngDstIdx
                                        If Not blnFound Then
                                            ReDim Preserve a_udtToAdd(lngTotalToAdd)
                                            With a_udtToAdd(lngTotalToAdd)
                                                .strKey = strKey
                                                .strValue = strValue
                                            End With
                                            lngTotalToAdd = lngTotalToAdd + 1
                                        End If
                                    Next lngSrcIdx
                                    'add the new keys
                                    For lngIdx = 0 To lngTotalToAdd - 1
                                        ReDim Preserve a_udtDstKVLines(lngTotalDstKVLines)
                                        With a_udtDstKVLines(lngTotalDstKVLines)
                                            .blnFlagged = False
                                            .blnIsKV = True
                                            .strKey = a_udtToAdd(lngIdx).strKey
                                            .strValue = a_udtToAdd(lngIdx).strValue
                                            .strLine = .strKey & "=" & .strValue
                                        End With
                                        lngTotalDstKVLines = lngTotalDstKVLines + 1
                                    Next lngIdx

                                ElseIf g_UpdateKVTypeEnum = e_Add_All_Leave Then
                                
                                    'ADD KEYS LEAVING EXISTING ONES UNCHANGED
                                    
                                    'get list of new keys to add
                                    ReDim a_udtToAdd(0)
                                    lngTotalToAdd = 0
                                    For lngSrcIdx = 0 To lngTotalSrcKV - 1
                                        strKey = a_udtSrcKV(lngSrcIdx).strKey
                                        strValue = a_udtSrcKV(lngSrcIdx).strValue
                                        blnFound = False
                                        For lngDstIdx = 0 To lngTotalDstKVLines - 1
                                            With a_udtDstKVLines(lngDstIdx)
                                                If .blnIsKV And (LCase$(.strKey) = LCase$(strKey)) Then
                                                    If .strKey = "kvset" Then
                                                        If .strValue = strValue Then
                                                            blnFound = True
                                                        End If
                                                    Else
                                                        blnFound = True
                                                    End If
                                                End If
                                            End With
                                            If blnFound Then Exit For
                                        Next lngDstIdx
                                        If Not blnFound Then
                                            ReDim Preserve a_udtToAdd(lngTotalToAdd)
                                            With a_udtToAdd(lngTotalToAdd)
                                                .strKey = strKey
                                                .strValue = strValue
                                            End With
                                            lngTotalToAdd = lngTotalToAdd + 1
                                        End If
                                    Next lngSrcIdx
                                    'add the new keys
                                    For lngIdx = 0 To lngTotalToAdd - 1
                                        ReDim Preserve a_udtDstKVLines(lngTotalDstKVLines)
                                        With a_udtDstKVLines(lngTotalDstKVLines)
                                            .blnFlagged = False
                                            .blnIsKV = True
                                            .strKey = a_udtToAdd(lngIdx).strKey
                                            .strValue = a_udtToAdd(lngIdx).strValue
                                            .strLine = .strKey & "=" & .strValue
                                        End With
                                        lngTotalDstKVLines = lngTotalDstKVLines + 1
                                    Next lngIdx
                                
                                ElseIf g_UpdateKVTypeEnum = e_Remove_All_Matching_Keys Then
                                
                                    'REMOVE KEYS REGARDLESS OF VALUE
                                    
                                    For lngSrcIdx = 0 To lngTotalSrcKV - 1
                                        strKey = a_udtSrcKV(lngSrcIdx).strKey
                                        For lngDstIdx = 0 To lngTotalDstKVLines - 1
                                            With a_udtDstKVLines(lngDstIdx)
                                                If .blnIsKV And (LCase$(.strKey) = LCase$(strKey)) Then
                                                    .blnFlagged = True
                                                    'do not break here as could be multiple kvsets
                                                End If
                                            End With
                                        Next lngDstIdx
                                    Next lngSrcIdx
                                
                                ElseIf g_UpdateKVTypeEnum = e_Remove_All_Matching_Exactly Then
                                
                                    'REMOVE KEYS IF VALUES MATCH EXACTLY
                                
                                    For lngSrcIdx = 0 To lngTotalSrcKV - 1
                                        strKey = a_udtSrcKV(lngSrcIdx).strKey
                                        strValue = a_udtSrcKV(lngSrcIdx).strValue
                                        For lngDstIdx = 0 To lngTotalDstKVLines - 1
                                            With a_udtDstKVLines(lngDstIdx)
                                                If .blnIsKV And (LCase$(.strKey) = LCase$(strKey)) And (.strValue = strValue) Then
                                                    .blnFlagged = True
                                                    'do not break here as could be multiple kvsets
                                                End If
                                            End With
                                        Next lngDstIdx
                                    Next lngSrcIdx
    
                                End If

                                'write back to db the lines that weren't flagged
                                strGemus = ""
                                For lngIdx = 0 To lngTotalDstKVLines - 1
                                    With a_udtDstKVLines(lngIdx)
                                        If Not .blnFlagged Then
                                            If .blnIsKV Then
                                                strGemus = strGemus & .strKey & "=" & .strValue & vbCrLf
                                            Else
                                                strGemus = strGemus & .strLine & vbCrLf
                                            End If
                                        End If
                                    End With
                                Next lngIdx
                                strGemus = g_xobjFunc.RTrimWhitespace(strGemus)
                                .Edit
                                fldGemus = strGemus
                                .Update

                            End If
                        End If
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldGemus = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Private Sub UpdatePersonalInfo()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldHS As Field
    Dim fldDiff As Field
    Dim fldRating As Field
    Dim fldTP As Field
    Dim fldClassic As Field
    Dim fldLP As Field
    Dim fldAdult As Field
    Dim fldFav As Field
    Dim fldNotes As Field

    'save game info to all games in the current view
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, HighScore, DI_Id, Rating, Classic, TimesPlayed, DateLastPlayed, Adult, Fav, MemoText FROM Games ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldHS = !HighScore
                Set fldDiff = !DI_Id
                Set fldRating = !Rating
                Set fldTP = !TimesPlayed
                Set fldClassic = !Classic
                Set fldLP = !DateLastPlayed
                Set fldAdult = !Adult
                Set fldFav = !Fav
                Set fldNotes = !MemoText
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            With g_udtEditPersonalInfo
                                If .blnUpdateAdult = True Then
                                    fldAdult = .blnAdult
                                End If
                                If .blnUpdateDiff = True Then
                                    fldDiff = .lngDiffID
                                End If
                                If .blnUpdateFav = True Then
                                    fldFav = .blnFav
                                End If
                                If .blnUpdateHighScore = True Then
                                    fldHS = .strHighScore
                                End If
                                If .blnUpdateLP = True And .blnLP = True Then
                                    fldLP = ""
                                End If
                                If .blnUpdateNotes = True Then
                                    fldNotes = .strNotes
                                End If
                                If .blnUpdateRating = True Then
                                    fldClassic = .blnClassic
                                    fldRating = .lngRating
                                End If
                                If .blnUpdateTP = True And .blnTP = True Then
                                    fldTP = 0
                                End If
                            End With
                            .Update
                        End If
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldHS = Nothing
                Set fldDiff = Nothing
                Set fldRating = Nothing
                Set fldTP = Nothing
                Set fldClassic = Nothing
                Set fldLP = Nothing
                Set fldAdult = Nothing
                Set fldFav = Nothing
                Set fldNotes = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Public Function UpdateVersionInfo()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldCra As Field
    Dim fldTrainers As Field
    Dim fldHSS As Field
    Dim fldLen As Field
    Dim fldLenType As Field
    Dim fldNTSC As Field
    Dim fldTDE As Field
    Dim fldLoadScrn As Field
    Dim fldDocs As Field
    Dim fldCom As Field
    Dim fldWLN As Field
    Dim fldWLU As Field
    Dim fldPlayable As Field
    Dim fldOriginal As Field
    Dim fldTitleScreen As Field

    'save game info to all games in the current view
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, CR_Id, V_Trainers, V_HighScoreSaver, V_Length, V_LengthType, V_PalNTSC, V_TrueDriveEmu, V_LoadingScreen, V_IncludedDocs, V_Comment, V_WebLink_Name, V_WebLink_URL, V_Playable, V_Original, V_TitleScreen FROM Games ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldCra = !CR_Id
                Set fldTrainers = !V_Trainers
                Set fldHSS = !V_HighScoreSaver
                Set fldLen = !V_Length
                Set fldLenType = !V_LengthType
                Set fldNTSC = !V_PalNTSC
                Set fldTDE = !V_TrueDriveEmu
                Set fldLoadScrn = !V_LoadingScreen
                Set fldDocs = !V_IncludedDocs
                Set fldCom = !V_Comment
                Set fldWLN = !V_WebLink_Name
                Set fldWLU = !V_WebLink_URL
                Set fldPlayable = !V_Playable
                Set fldOriginal = !V_Original
                Set fldTitleScreen = !V_TitleScreen
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            With g_udtEditVersionInfo
                                If .blnUpdateCracker Then
                                    fldCra = .lngCrackerID
                                End If
                                If .blnUpdateDocs Then
                                    fldDocs = .lngDocs
                                End If
                                If .blnUpdateHSS Then
                                    fldHSS = .lngHSS
                                End If
                                If .blnUpdateLength Then
                                    fldLen = .lngGameLength
                                    fldLenType = .lngGameLengthType
                                End If
                                If .blnUpdateLoadingScreen Then
                                    fldLoadScrn = .lngLoadingScreen
                                End If
                                If .blnUpdateNTSC Then
                                    fldNTSC = .lngNTSC
                                End If
                                If .blnUpdateTDE Then
                                    fldTDE = .lngTDE
                                End If
                                If .blnUpdateTrainers Then
                                    fldTrainers = .lngTrainers
                                End If
                                If .blnUpdateVComment Then
                                    fldCom = .strVComment
                                End If
                                If .blnUpdateOriginal Then
                                    fldOriginal = .lngOriginal
                                End If
                                If .blnUpdatePlayable Then
                                    fldPlayable = .lngPlayable
                                End If
                                If .blnUpdateTitleScreen Then
                                    fldTitleScreen = .lngTitleScreen
                                End If
                                If .blnUpdateWebLink Then
                                    fldWLN = .strWebLinkName
                                    fldWLU = .strWebLinkURL
                                End If
                            End With
                            .Update
                        End If
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldCra = Nothing
                Set fldTrainers = Nothing
                Set fldHSS = Nothing
                Set fldLen = Nothing
                Set fldLenType = Nothing
                Set fldNTSC = Nothing
                Set fldTDE = Nothing
                Set fldLoadScrn = Nothing
                Set fldDocs = Nothing
                Set fldCom = Nothing
                Set fldWLN = Nothing
                Set fldWLU = Nothing
                Set fldPlayable = Nothing
                Set fldOriginal = Nothing
                Set fldTitleScreen = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing


End Function

Private Sub UpdateMusicInfo()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldFav As Field
    Dim fldAdult As Field
    Dim fldMusician As Field

    'save game info to all games in the current view
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, SFav, Adult, MU_Id FROM Music ORDER BY Name", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldFav = !SFav
                Set fldAdult = !Adult
                Set fldMusician = !MU_Id
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "M" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            With g_udtEditMusicInfo
                                If .blnUpdateAdult = True Then
                                    fldAdult = .blnAdult
                                End If
                                If .blnUpdateFav = True Then
                                    fldFav = .blnFav
                                End If
                                If .blnUpdateMusician = True Then
                                    fldMusician = .lngMusicianID
                                End If
                            End With
                            .Update
                        End If
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldFav = Nothing
                Set fldAdult = Nothing
                Set fldMusician = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Private Sub UpdateSidFile()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldID As Field
    Dim fldFA As Field
    Dim fldSA As Field
    Dim fldCCode As Field
    Dim fldFilename As Field
    Dim strFilename As String

    lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
    If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
        pbStatus.Min = 0
        pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
        pbStatus.Value = 0
        Screen.MousePointer = vbHourglass
        DoEvents
        strFilename = IIf((g_strSidSelectorFilename = g_strNone), "", g_strSidSelectorFilename)

        'save sid file to all GAMES in the current view (calculate FA, SA and CCode)
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, SidFilename, FA, CCode, SA FROM Games ORDER BY Name", dbOpenDynaset)
        With g_rsGames
            If Not .EOF Then
                Set fldID = !GA_Id
                Set fldFA = !FA
                Set fldSA = !SA
                Set fldCCode = !CCode
                Set fldFilename = !SidFilename
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            fldSA = (strFilename <> "")
                            fldFilename = strFilename
                            If strFilename <> "" Then
                                If fldFA = True Then
                                     fldCCode = 1 'both exist
                                Else
                                    fldCCode = 3 'only sid exists
                                End If
                            Else
                                If fldFA = True Then
                                    fldCCode = 2 'only game exists
                                Else
                                    fldCCode = 4 'none exist
                                End If
                            End If
                            .Update
                            pbStatus.Value = pbStatus.Value + 1
                        End If
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit For
                        End If
                    End If
                Next LItem
                DoEvents
                Set fldID = Nothing
                Set fldFA = Nothing
                Set fldSA = Nothing
                Set fldCCode = Nothing
                Set fldFilename = Nothing
            End If
            .Close
        End With
        Set g_rsGames = Nothing
    
        'save sid file to all SIDS in the current view
        If m_blnCancelled = False Then
            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Filename, SA FROM Music ORDER BY Name", dbOpenDynaset)
            With g_rsGames
                If Not .EOF Then
                    Set fldID = !GA_Id
                    Set fldSA = !SA
                    Set fldFilename = !Filename
                    For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                        If Left$(LItem.Key, 1) = "M" Then
                            .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                            If Not .NoMatch Then
                                .Edit
                                fldSA = (strFilename <> "")
                                fldFilename = strFilename
                                .Update
                                pbStatus.Value = pbStatus.Value + 1
                            End If
                            DoEvents
                            If m_blnCancelled = True Then
                                Exit For
                            End If
                        End If
                    Next LItem
                    pbStatus.Value = pbStatus.Max
                    DoEvents
                    Set fldID = Nothing
                    Set fldFA = Nothing
                    Set fldSA = Nothing
                    Set fldCCode = Nothing
                    Set fldFilename = Nothing
                End If
                .Close
            End With
            Set g_rsGames = Nothing
        End If
    End If

End Sub
Private Sub UpdatePictureFile()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim fldPicture As Field

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, ScrnShotFilename FROM Games", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
            If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
                Set fldPicture = !ScrnShotFilename
                pbStatus.Min = 0
                pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
                pbStatus.Value = 0
                Screen.MousePointer = vbHourglass
                DoEvents
                For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
                    If Left$(LItem.Key, 1) = "G" Then
                        .FindFirst "GA_Id = " & Right$(LItem.Key, Len(LItem.Key) - 1)
                        If Not .NoMatch Then
                            .Edit
                            fldPicture = g_strScrnShotSelectorFilename
                            .Update
                        End If
                    End If
                    pbStatus.Value = pbStatus.Value + 1
                    DoEvents
                    If m_blnCancelled = True Then
                        Exit For
                    End If
                Next LItem
                pbStatus.Value = pbStatus.Max
                DoEvents
                Set fldPicture = Nothing
            End If
        End If
        .Close
    End With
    Set g_rsGames = Nothing

End Sub

Private Sub UpdateExtras()

    Dim lngListViewIdx As Long
    Dim LItem As ListItem
    Dim lngID As Long

    'remove extras for all games in the list
    lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
    If frmMain.lvMain(lngListViewIdx).ListItems.Count > 0 Then
        pbStatus.Min = 0
        pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
        pbStatus.Value = 0
        Screen.MousePointer = vbHourglass
        DoEvents
        For Each LItem In frmMain.lvMain(lngListViewIdx).ListItems
            If Left$(LItem.Key, 1) = "G" Then
                lngID = Val(Right$(LItem.Key, Len(LItem.Key) - 1))
                g_dbGamebase.Execute "UPDATE Games SET Extras = False WHERE GA_Id = " & lngID
                g_dbGamebase.Execute "DELETE * FROM Extras WHERE GA_Id = " & lngID
            End If
            pbStatus.Value = pbStatus.Value + 1
            DoEvents
            If m_blnCancelled = True Then
                Exit For
            End If
        Next LItem
        pbStatus.Value = pbStatus.Max
        DoEvents
    End If


End Sub

Private Sub cmdCancel_Click()

    m_blnCancelled = True

End Sub

Private Sub Form_Activate()

    If m_blnUpdateGameInfo = True Then
        Call UpdateGameInfo
    ElseIf m_blnUpdateVersionInfo = True Then
        Call UpdateVersionInfo
    ElseIf m_blnUpdatePersonalInfo = True Then
        Call UpdatePersonalInfo
    ElseIf m_blnUpdateGameFile = True Then
        Call UpdateGameFile
    ElseIf m_blnUpdateSidFile = True Then
        Call UpdateSidFile
    ElseIf m_blnUpdatePictureFile = True Then
        Call UpdatePictureFile
    ElseIf m_blnUpdateMusicInfo = True Then
        Call UpdateMusicInfo
    ElseIf m_blnUpdateExtras = True Then
        Call UpdateExtras
    ElseIf m_blnExportToTextFile = True Then
        Call ExportToTextFile
    ElseIf m_blnUpdateKV = True Then
        Call UpdateKV
    End If
    Unload Me

End Sub
Private Sub Form_Load()

    Const DesignScaleHeight As Single = 1335    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5175     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    m_blnUpdateGameInfo = False
    m_blnUpdateVersionInfo = False
    m_blnUpdatePersonalInfo = False
    m_blnUpdateMusicInfo = False
    m_blnUpdateGameFile = False
    m_blnUpdateSidFile = False
    m_blnUpdatePictureFile = False
    m_blnUpdateExtras = False
    m_blnExportToTextFile = False
    m_blnCancelled = False
    m_blnUpdateKV = False

    cmdCancel.Caption = g_strButtonCancelText
    Caption = g_xobjFunc.GetINIMessage("Batch Update", "Caption", "Batch Update") & "..."
    lblPrompt = g_xobjFunc.GetINIMessage("Batch Update", "Prompt", "Updating records") & "..."

    Screen.MousePointer = vbHourglass
    DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


