VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportPersonal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Personal Info"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmImportPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   3120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkDoNotPreserve 
      Caption         =   "Reset all items that are n&ot updated to their defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   6615
   End
   Begin VB.Frame fraFields 
      Caption         =   "Fields to Import"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   6615
      Begin VB.CheckBox chkAdult 
         Caption         =   "&Adult Game"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chkRating 
         Caption         =   "&Rating"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkDLP 
         Caption         =   "Date &Last Played"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkTP 
         Caption         =   "&Times Played"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkDifficulty 
         Caption         =   "&Difficulty"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chkHighScore 
         Caption         =   "&High Score"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkFavMusic 
         Caption         =   "Favourite &Music"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkFavGames 
         Caption         =   "Favourite &Games"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Import"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label lblDestDB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblDestDBPrompt 
      Caption         =   "Destination GameBase database (currently loaded):"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label lblSrcDBPrompt 
      Caption         =   "&Source GameBase database to import from:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblSrcDB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "frmImportPersonal"
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

Private m_blnIsImporting As Boolean
Private m_blnCancelled As Boolean

Private Function ImportPersonalInfo() As Boolean

'    On Error GoTo Err_Handler

    Dim dbSRC As Database
    Dim strSQL As String
    Dim lngTotalGames As Long
    Dim lngTotalMusic As Long
    Dim lngTotalGamesUpdated As Long
    Dim lngTotalGamesReset As Long
    Dim lngTotalMusicUpdated As Long
    Dim lngTotalMusicReset As Long
    Dim lngTotalGamesUnprocessed As Long
    Dim lngTotalMusicUnprocessed As Long
    Dim rsGames As Recordset
    Dim rsSrc As Recordset
    Dim lngID As Long
    Dim strFieldList As String
    Dim blnDoFavGames As Boolean
    Dim blnDoFavMusic As Boolean
    Dim blnDoHighScore As Boolean
    Dim blnDoDifficulty As Boolean
    Dim blnDoTP As Boolean
    Dim blnDoDLP As Boolean
    Dim blnDoRating As Boolean
    Dim blnDoAdult As Boolean
    Dim blnDoClear As Boolean
    Dim blnDoGames As Boolean
    Dim strMsg As String
    Dim NL As String
    Dim BL As String
    Dim T As String
    Dim strDifficultySrc As String
    Dim strDifficultyDst As String
    Dim blnDifficultyDstHasUnknown As Boolean
    

    NL = vbCrLf
    BL = NL & NL
    T = vbTab

    blnDoFavGames = (chkFavGames.Value = vbChecked)
    blnDoFavMusic = (chkFavMusic.Value = vbChecked)
    blnDoHighScore = (chkHighScore.Value = vbChecked)
    blnDoDifficulty = (chkDifficulty.Value = vbChecked)
    blnDoTP = (chkTP.Value = vbChecked)
    blnDoDLP = (chkDLP.Value = vbChecked)
    blnDoRating = (chkRating.Value = vbChecked)
    blnDoAdult = (chkAdult.Value = vbChecked)
    blnDoClear = (chkDoNotPreserve.Value = vbChecked)

    lngTotalGamesUpdated = 0
    lngTotalGamesReset = 0
    lngTotalMusicUpdated = 0
    lngTotalMusicReset = 0

    Set dbSRC = OpenDatabase(lblSrcDB.Caption)

    If blnDoDifficulty Then
        'can only update Difficulty if the difficulty table is the same in both current and source dbs
        'and id 8 is (unknown)
        strDifficultySrc = ""
        strDifficultyDst = ""
        blnDifficultyDstHasUnknown = False
        strSQL = "SELECT DI_Id, Difficulty FROM Difficulty"
        Set rsGames = g_dbGamebase.OpenRecordset(strSQL, dbOpenSnapshot)
        With rsGames
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    strDifficultyDst = strDifficultyDst & !DI_Id & !Difficulty
                    If (!DI_Id = 8) And (UCase$(!Difficulty) = "(UNKNOWN)") Then blnDifficultyDstHasUnknown = True
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rsGames = Nothing
        Set rsSrc = dbSRC.OpenRecordset(strSQL, dbOpenSnapshot)
        With rsSrc
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    strDifficultySrc = strDifficultySrc & !DI_Id & !Difficulty
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rsSrc = Nothing

        If (strDifficultySrc <> strDifficultyDst) Or (blnDifficultyDstHasUnknown = False) Then
            MsgBox "Difficulty will not be imported as the Difficulty records differ between databases, or there is no ""(unknown)"" record", vbOKOnly Or vbExclamation Or vbApplicationModal, Caption
            blnDoDifficulty = False
        End If
    End If

    blnDoGames = blnDoFavGames Or blnDoHighScore Or blnDoDifficulty Or blnDoTP Or blnDoDLP Or blnDoRating Or blnDoAdult
    If blnDoGames Then

        strFieldList = "Games.GA_Id, Games.Adult, Games.Fav, Games.HighScore, Games.DI_Id, " & _
            "Games.TimesPlayed, Games.DateLastPlayed, Games.Rating, Games.Classic"
    
        strSQL = "SELECT " & strFieldList & " FROM Games ORDER BY GA_Id ASC"
    
        'process each game in the currently loaded db
        Set rsGames = g_dbGamebase.OpenRecordset(strSQL, dbOpenDynaset)
        With rsGames
            If Not .EOF Then
                .MoveLast
                .MoveFirst
                lngTotalGames = .RecordCount
                If lngTotalGames > 0 Then
                    pbStatus.Min = 0
                    pbStatus.Max = lngTotalGames
                    pbStatus.Value = 0
                    pbStatus.Visible = True
                    lblStatus.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "ImportStatusGames", "Importing to Games") & "..."
                    DoEvents
                    Do Until .EOF
                        'find this game in the src db
                        lngID = !GA_Id
                        strSQL = "SELECT " & strFieldList & " FROM Games WHERE GA_Id=" & lngID
                        Set rsSrc = dbSRC.OpenRecordset(strSQL, dbOpenSnapshot)
                        If Not rsSrc.EOF Then
                        
                            'IN SRC DB: COPY ACROSS FIELD VALUES
                            
                            rsSrc.MoveFirst
                            .Edit
                            If blnDoFavGames Then
                                !Fav = rsSrc!Fav
                            End If
                            If blnDoHighScore Then
                                !HighScore = rsSrc!HighScore
                            End If
                            If blnDoDifficulty Then
                                !DI_Id = rsSrc!DI_Id
                            End If
                            If blnDoTP Then
                                !TimesPlayed = rsSrc!TimesPlayed
                            End If
                            If blnDoDLP Then
                                !DateLastPlayed = rsSrc!DateLastPlayed
                            End If
                            If blnDoRating Then
                                !Rating = rsSrc!Rating
                                !Classic = rsSrc!Classic
                            End If
                            If blnDoAdult Then
                                !Adult = rsSrc!Adult
                            End If
                            .Update
                            lngTotalGamesUpdated = lngTotalGamesUpdated + 1
                            
                        ElseIf blnDoClear Then
                        
                            'NOT IN SRC DB: CLEAR FIELD VALUES

                            .Edit
                            If blnDoFavGames Then
                                !Fav = False
                            End If
                            If blnDoHighScore Then
                                !HighScore = ""
                            End If
                            If blnDoDifficulty Then
                                !DI_Id = 8 'unknown
                            End If
                            If blnDoTP Then
                                !TimesPlayed = 0
                            End If
                            If blnDoDLP Then
                                !DateLastPlayed = ""
                            End If
                            If blnDoRating Then
                                !Rating = 0 'unknown
                                !Classic = False
                            End If
                            If blnDoAdult Then
                                !Adult = False
                            End If
                            .Update
                            lngTotalGamesReset = lngTotalGamesReset + 1
            
                        End If
                        rsSrc.Close
                        Set rsSrc = Nothing
                        'next game
                        pbStatus.Value = .AbsolutePosition
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                    pbStatus.Value = pbStatus.Max
                    DoEvents
                End If
            End If
            .Close
        End With
        Set rsGames = Nothing
    End If

    If blnDoFavMusic = True And m_blnCancelled = False Then

        strFieldList = "Music.GA_Id, Music.SFav"
        strSQL = "SELECT " & strFieldList & " FROM Music ORDER BY GA_Id ASC"
    
        'process each game in the currently loaded db
        Set rsGames = g_dbGamebase.OpenRecordset(strSQL, dbOpenDynaset)
        With rsGames
            If Not .EOF Then
                .MoveLast
                .MoveFirst
                lngTotalMusic = .RecordCount
                If lngTotalMusic > 0 Then
                    pbStatus.Min = 0
                    pbStatus.Max = lngTotalMusic
                    pbStatus.Value = 0
                    pbStatus.Visible = True
                    lblStatus.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "ImportStatusMusic", "Importing to Music") & "..."
                    DoEvents
                    Do Until .EOF
                        'find this game in the src db
                        lngID = !GA_Id
                        strSQL = "SELECT " & strFieldList & " FROM Music WHERE GA_Id=" & lngID
                        Set rsSrc = dbSRC.OpenRecordset(strSQL, dbOpenSnapshot)
                        If Not rsSrc.EOF Then
                        
                            'IN SRC DB: COPY ACROSS FIELD VALUES
                            
                            rsSrc.MoveFirst
                            .Edit
                            !SFav = rsSrc!SFav
                            .Update
                            lngTotalMusicUpdated = lngTotalMusicUpdated + 1
                            
                        ElseIf blnDoClear Then
                        
                            'NOT IN SRC DB: CLEAR FIELD VALUES

                            .Edit
                            !SFav = False
                            .Update
                            lngTotalMusicReset = lngTotalMusicReset + 1

                        End If
                        rsSrc.Close
                        Set rsSrc = Nothing
                        'next music
                        pbStatus.Value = .AbsolutePosition
                        DoEvents
                        If m_blnCancelled = True Then
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                    pbStatus.Value = pbStatus.Max
                    DoEvents
                End If
            End If
            .Close
        End With
        Set rsGames = Nothing
    End If

    dbSRC.Close
    Set dbSRC = Nothing

    'show summary
    lngTotalGamesUnprocessed = (lngTotalGames - (lngTotalGamesUpdated + lngTotalGamesReset))
    lngTotalMusicUnprocessed = (lngTotalMusic - (lngTotalMusicUpdated + lngTotalMusicReset))
    strMsg = ""
    If m_blnCancelled Then
        strMsg = "OPERATION CANCELLED!" & BL
    End If
    If blnDoGames Then
        strMsg = strMsg & "GAMES:" & NL & _
            T & "Total:" & T & T & lngTotalGames & NL & _
            T & "Imported:" & T & lngTotalGamesUpdated & NL
        If blnDoClear Then
            strMsg = strMsg & T & "Reset:" & T & T & lngTotalGamesReset & NL
        End If
        If lngTotalGamesUnprocessed > 0 Then
            strMsg = strMsg & T & "Unprocessed:" & T & lngTotalGamesUnprocessed & NL
        End If
    End If
    If blnDoFavMusic Then
        strMsg = strMsg & NL & "MUSIC:" & NL & _
            T & "Total:" & T & T & lngTotalMusic & NL & _
            T & "Imported:" & T & lngTotalMusicUpdated & NL
        If blnDoClear Then
            strMsg = strMsg & T & "Reset:" & T & T & lngTotalMusicReset & NL
        End If
        If lngTotalMusicUnprocessed > 0 Then
            strMsg = strMsg & T & "Unprocessed:" & T & lngTotalMusicUnprocessed & NL
        End If
    End If
    
    MsgBox strMsg, vbInformation Or vbOKOnly Or vbApplicationModal, Caption

    'return true if no errors and not cancelled
    ImportPersonalInfo = (m_blnCancelled = False)
    Exit Function

Err_Handler:
    ImportPersonalInfo = False
    Set dbSRC = Nothing
    MsgBox Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME

End Function

Private Sub SetOKButton()

    cmdOK.Enabled = (lblSrcDB.Caption <> "") And (UCase$(lblSrcDB.Caption) <> UCase$(lblDestDB.Caption)) And _
    ( _
        (chkAdult.Value = vbChecked) Or _
        (chkDifficulty.Value = vbChecked) Or _
        (chkDLP.Value = vbChecked) Or _
        (chkFavGames.Value = vbChecked) Or _
        (chkFavMusic.Value = vbChecked) Or _
        (chkHighScore.Value = vbChecked) Or _
        (chkRating.Value = vbChecked) Or _
        (chkTP.Value = vbChecked) _
    )

End Sub


Private Sub chkAdult_Click()
    Call SetOKButton
End Sub

Private Sub chkDifficulty_Click()
    Call SetOKButton
End Sub


Private Sub chkDLP_Click()
    Call SetOKButton
End Sub


Private Sub chkFavGames_Click()
    Call SetOKButton
End Sub


Private Sub chkFavMusic_Click()
    Call SetOKButton
End Sub


Private Sub chkHighScore_Click()
    Call SetOKButton
End Sub


Private Sub chkRating_Click()
    Call SetOKButton
End Sub


Private Sub chkTP_Click()
    Call SetOKButton
End Sub


Private Sub cmdBrowse_Click()

    Dim strLastDir As String    'last dir traversed
    Dim strFile As String       'for splitting into path\file vars
    Dim strPath As String       'for splitting into path\file vars

    On Error GoTo Cancel_Pressed

    'set up the open dialog
    With cDialog
        .CancelError = True
        .DefaultExt = ".mdb"
        .DialogTitle = g_xobjFunc.GetINIMessage("Import Personal Info Window", "FilePrompt", "Select Source GameBase database")
        .Filter = "GameBase Databases (*.mdb)|*.mdb"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "ImportPersonalPath", strLastDir, True, App.Path
        If lblSrcDB = "" Then
            If g_xobjFunc.DirExists(strLastDir) = True Then
                .InitDir = strLastDir
            Else
                .InitDir = App.Path
            End If
            .Filename = ""
        ElseIf g_xobjFunc.FileExists(lblSrcDB, False, "") = True Then
            .Filename = lblSrcDB
        ElseIf g_xobjFunc.DirExists(strLastDir) = True Then
            .InitDir = strLastDir
            .Filename = ""
        Else
            .InitDir = App.Path
            .Filename = ""
        End If
    End With
    
    'get the file
    cDialog.ShowOpen
    lblSrcDB = cDialog.Filename

    'save the last path
    g_xobjFunc.GetFileAndPath cDialog.Filename, strPath, strFile
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "ImportPersonalPath", strPath, True

Cancel_Pressed:

    'set the ok button
    Call SetOKButton

End Sub

Private Sub cmdCancel_Click()

    If m_blnIsImporting = True Then
        m_blnCancelled = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()

    Dim strMsg As String
    Dim NL As String
    Dim BL As String
    Dim T As String
    Dim blnSuccess As Boolean
    
    'check not importing from the same .mdb file
    If UCase$(lblSrcDB.Caption) = UCase$(g_strGameBaseDBPath) Then
        MsgBox g_xobjFunc.GetINIMessage("Import Personal Info Window", "ErrorSameDB", "You can't import from the same database!"), vbExclamation Or vbOKOnly Or vbApplicationModal, Caption
        cmdBrowse.SetFocus
        Exit Sub
    End If
    
    'show the warning message
    NL = vbCrLf
    BL = NL & NL
    T = vbTab
    With g_xobjFunc
        strMsg = .GetINIMessage("Import Personal Info Window", "Warning1", "You are about to import the selected personal fields from database") & ":" & BL & _
                T & lblSrcDB.Caption & BL & _
                .GetINIMessage("Import Personal Info Window", "Warning2", "into the currently loaded database") & ":" & BL & _
                T & g_strGameBaseDBPath & BL
        If chkDoNotPreserve.Value = vbChecked Then
            strMsg = strMsg & .GetINIMessage("Import Personal Info Window", "Warning3", "All selected fields for games/music that are not updated will be reset to their default values.") & BL
        End If
        strMsg = strMsg & .GetINIMessage("Import Personal Info Window", "Warning4", "The import is ID based, meaning games/music records will be updated where the GameBase IDs match in both databases.") & BL
        strMsg = strMsg & .GetINIMessage("Import Personal Info Window", "Warning5", "ARE YOU SURE you want to continue?")
    End With

    If MsgBox(strMsg, vbExclamation Or vbYesNoCancel Or vbDefaultButton3 Or vbApplicationModal, Caption) = vbYes Then

        'do the import
        g_blnImportPersonalInfoSuccess = True 'started so force view refresh in main window
        m_blnIsImporting = True
        m_blnCancelled = False
        cmdOK.Enabled = False
        cmdBrowse.Enabled = False
        Screen.MousePointer = vbHourglass
        DoEvents

        blnSuccess = ImportPersonalInfo()

        m_blnIsImporting = False
        m_blnCancelled = False
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True
        cmdBrowse.Enabled = True
        
        If blnSuccess Then Unload Me

    End If

End Sub
Private Sub Form_Load()

    Const DesignScaleHeight As Single = 4575    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6855     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'read in INI text
    lblSrcDBPrompt.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "SourceDBPrompt", "&Source GameBase database to import from") & ":"
    lblDestDBPrompt.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "DestDBPrompt", "Destination GameBase (currently loaded)") & ":"
    chkDoNotPreserve.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "ClearCheckbox", "Reset selected &fields to their defaults for all items that are not updated")
    cmdOK.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "ImportButton", "&Import") & "..."
    chkFavGames.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "FavGamesField", "Favourite &Games")
    chkFavMusic.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "FavMusicField", "Favourite &Music")
    chkHighScore.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "HighScoreField", "&High Score")
    chkDifficulty.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "DifficultyField", "&Difficulty")
    chkTP.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "TimesPlayedField", "&Times Played")
    chkDLP.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "DateLastPlayedField", "Date &Last Played")
    chkRating.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "RatingField", "&Rating")
    chkAdult.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "AdultGameField", "&Adult Game")
    
    Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "Caption", "Import Personal Info")
    fraFields.Caption = g_xobjFunc.GetINIMessage("Import Personal Info Window", "FieldsToImport", "Fields to Import")

    lblDestDB.Caption = g_strGameBaseDBPath

    'disable the ok button until a source db is chosen and at least one field is ticked
    Call SetOKButton

    'not started yet
    m_blnIsImporting = False
    m_blnCancelled = False
    g_blnImportPersonalInfoSuccess = False

End Sub


