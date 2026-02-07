VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRomLister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to RomLister"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   ControlBox      =   0   'False
   Icon            =   "frmRomLister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkExt 
      Caption         =   "with filename &extension"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.OptionButton optRomName 
      Caption         =   "Game &Filename"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optRomName 
      Caption         =   "Game &ID"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDummy 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox chkScreenshots 
      Caption         =   "Convert &Screenshots"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6735
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line lineCommentTop 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line lineCommentBot 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   6840
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label lblPrompt2 
      Caption         =   "&Rom Name:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   6735
   End
   Begin VB.Label lblPrompt 
      Caption         =   "&Put converted screenshots here:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label lblDestScreenshotPath 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MouseIcon       =   "frmRomLister.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "frmRomLister"
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

Private m_strDestScreenshotPath As String 'last screenshot destination path traversed
Private m_blnIsExporting As Boolean
Private m_blnCancelled As Boolean
Private Sub chkScreenshots_Click()

    With lblDestScreenshotPath
        If chkScreenshots.Value = vbChecked Then
            .Enabled = True
            .BackColor = txtDummy.BackColor
        Else
            .Enabled = False
            .BackColor = cmdOK.BackColor
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    If m_blnIsExporting = True Then
        m_blnCancelled = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdHelp_Click()

    Call g_xobjFunc.ShellExecuteFile("http://www.bu22.com/wiki/exporting_gamebases_to_other_frontends", "", False)

End Sub

Private Sub cmdOK_Click()

    Dim blnWasLabelEnabled As Boolean
    
    blnWasLabelEnabled = lblDestScreenshotPath.Enabled

    m_blnIsExporting = True
    m_blnCancelled = False

    cmdOK.Enabled = False
    lblDestScreenshotPath.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Call ExportToRomLister

    m_blnIsExporting = False
    m_blnCancelled = False
    Screen.MousePointer = vbDefault
    cmdOK.Enabled = True
    lblDestScreenshotPath.Enabled = blnWasLabelEnabled

End Sub

Private Sub ExportToRomLister()

    Dim lngListViewIdx As Long
    Dim strXML As String
    Dim strEntry As String
    Dim intFileNum As Integer
    Dim LItem As ListItem
    Dim strOutputFile As String
    Dim strGameID As String
    Dim strRomName As String
    Dim strQuery As String
    Dim strDescription As String
    Dim strManufacturer As String
    Dim strYear As String
    Dim strPlayers As String
    Dim strControl As String
    Dim strGenre As String
    Dim lngNumExported As Long
    Dim blnDoScreenshots As Boolean
    Dim lngScrnShotIdx As Long
    Dim strScreenshotSource As String
    Dim strScreenshotDestPathWBS As String
    Dim strExtension As String
    Dim strViewTitle As String
    Dim strValidatedViewTitle As String
    Dim strChar As String
    Dim lngIdx As Long
    Dim blnRomNameIsGameID As Boolean
    Dim strGameFilenamePathPart As String
    Dim strRomNameNoExt As String
    Dim lngDotPos As Long
    Dim strTemp As String
    Dim blnWithFilenameExtension As Boolean

    On Error GoTo Err_Handler
    
    'generate the sql
    strQuery = "SELECT "
    strQuery = strQuery & "Genres.Genre, Genres.GE_Id, Genres.PG_Id, "
    strQuery = strQuery & "Games.PU_Id, Games.YE_Id, "
    strQuery = strQuery & "Publishers.Publisher, "
    strQuery = strQuery & "Games.GA_Id, Games.Name, Games.Comment, Games.Control, "
    strQuery = strQuery & "Games.Filename, Games.FileToRun, Games.FilenameIndex, "
    strQuery = strQuery & "Games.ScrnShotFilename, "
    strQuery = strQuery & "Games.FA, Games.V_Comment, "
    strQuery = strQuery & "Games.PlayersTo, Games.PlayersFrom, "
    strQuery = strQuery & "Games.PlayersSim, "
    strQuery = strQuery & "Games.Adult "
    strQuery = strQuery & "FROM Publishers INNER JOIN "
    strQuery = strQuery & "(Genres INNER JOIN "
    strQuery = strQuery & "Games "
    strQuery = strQuery & "ON Genres.GE_Id = Games.GE_Id) "
    strQuery = strQuery & "ON Publishers.PU_Id = Games.PU_Id "
    strQuery = strQuery & "WHERE Games.GA_Id = "

    'using game id or game filename as the rom name?
    blnRomNameIsGameID = optRomName(0).Value
    blnWithFilenameExtension = (chkExt.Value = vbChecked)
    
    'copying/renaming screenshots?
    blnDoScreenshots = IIf(((chkScreenshots.Value = vbChecked) And (g_lngTotalScrnShotPaths > 0)), True, False)
    If Not g_xobjFunc.DirExists(m_strDestScreenshotPath) Then blnDoScreenshots = False
    strScreenshotDestPathWBS = m_strDestScreenshotPath
    If Right$(strScreenshotDestPathWBS, 1) <> "\" Then strScreenshotDestPathWBS = strScreenshotDestPathWBS & "\"

    lngNumExported = 0

    strViewTitle = ga_udtViewData(frmMain.lstView.ListIndex).strTitle
    strValidatedViewTitle = ""
    For lngIdx = 1 To Len(strViewTitle)
        strChar = Mid$(strViewTitle, lngIdx, 1)
        If InStr("\/:*?""<>|", strChar) = 0 Then
            strValidatedViewTitle = strValidatedViewTitle & strChar
        End If
    Next lngIdx
    strTemp = IIf(blnRomNameIsGameID, "(Game ID)", "(Game Filename)")
    If strValidatedViewTitle <> "" Then
        strOutputFile = g_strAppPathWithBackslash & g_strGameBaseName & "-" & strValidatedViewTitle & "-" & strTemp & "-romlister-export.xml"
    Else
        strOutputFile = g_strAppPathWithBackslash & g_strGameBaseName & "-" & strTemp & "-romlister-export.xml"
    End If
    lngListViewIdx = frmMain.GetListViewIdxFromViewIdx(frmMain.lstView.ListIndex)
    
    With frmMain.lvMain(lngListViewIdx)
        If .ListItems.Count > 0 Then
            pbStatus.Min = 0
            pbStatus.Max = frmMain.lvMain(lngListViewIdx).ListItems.Count
            pbStatus.Value = 0
            pbStatus.Visible = True
            Screen.MousePointer = vbHourglass
            DoEvents
            'delete the existing file
            If g_xobjFunc.FileExists(strOutputFile, False, gc_strAPPNAME) = True Then
                Call g_xobjFunc.DeletePathSilent(strOutputFile)
            End If
            'create the xml file and copy/rename any screenshots
            strXML = strXML & "<?xml version=""1.0""?>" & vbCrLf _
                & "<!-- This file generated by ROMLister -->" & vbCrLf _
                & "<!DOCTYPE mame [" & vbCrLf _
                & "]>" & vbCrLf _
                & "<mame>" & vbCrLf
            For Each LItem In .ListItems
                pbStatus.Value = pbStatus.Value + 1
                lblStatus.Caption = LItem.Text
                DoEvents
                If Mid$(LItem.Key, 1, 1) = "G" Then
                
                    '============== its a game, we don't export music ==============
                    
                    'grab the id
                    strGameID = Right$(LItem.Key, Len(LItem.Key) - 1)

                    'Run the query, get the Game info from the database and store globally
                    Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery & strGameID, dbOpenSnapshot)
                    If g_rsGames.EOF Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Could not run database query [" & strQuery & "]", vbExclamation, gc_strAPPNAME
                        Exit Sub
                    End If
                    With g_rsGames
                        If !Filename <> "" Then

                            If blnRomNameIsGameID Then
                                strRomName = strGameID
                            Else
                                Call g_xobjFunc.GetFileAndPath(!Filename, strGameFilenamePathPart, strRomName)
                                If strRomName = "" Then
                                    .Close
                                    Set g_rsGames = Nothing
                                    MsgBox "Game with ID [" & strGameID & "] and Game Filename [" & !Filename & "] has empty filename part", vbExclamation, gc_strAPPNAME
                                    Exit Sub
                                End If
                                If Not blnWithFilenameExtension Then
                                    'strip the filename extension from the rom name
                                    lngDotPos = InStrRev(strRomName, ".")
                                    If lngDotPos > 0 Then
                                        strRomName = Left$(strRomName, lngDotPos - 1)
                                    End If
                                End If
                            End If

                            'format fields
                            strDescription = !Name
                            strYear = ga_strYearLookupTable(!YE_ID)
                            strManufacturer = !Publisher
                            strGenre = ga_strGenreParentLookupTable(!PG_Id) & !Genre
                            If !PlayersTo = !PlayersFrom Then
                                If !PlayersTo = gc_lngPLAYERSUNKNOWN Then
                                    strPlayers = "(" & g_strUnknown & ")"
                                Else
                                    strPlayers = !PlayersFrom & "P Only"
                                End If
                            Else
                                If !PlayersFrom = gc_lngPLAYERSUNKNOWN Then
                                    strPlayers = "?"
                                Else
                                    strPlayers = !PlayersFrom
                                End If
                                strPlayers = strPlayers & " - "
                                If !PlayersTo = gc_lngPLAYERSUNKNOWN Then
                                    strPlayers = strPlayers & "?"
                                Else
                                    strPlayers = strPlayers & !PlayersTo
                                End If
                            End If
                            If !PlayersSim = True Then
                                strPlayers = strPlayers & " (Simultaneous)"
                            End If
                            strControl = ga_strControlLookupTable(!Control)

                            'build the game node
                            strEntry = vbTab & "<game name=""" & strRomName & """>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<description>" & XMLEscape(strDescription) & "</description>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<year>" & strYear & "</year>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<manufacturer>" & XMLEscape(strManufacturer) & "</manufacturer>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<category>" & XMLEscape(strGenre) & "</category>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<gb_genre>" & XMLEscape(strGenre) & "</gb_genre>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<nPlayers>" & XMLEscape(strPlayers) & "</nPlayers>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<control type=""" & XMLEscape(strControl) & """ />" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<rl_buttons>buttons=2</rl_buttons>" & vbCrLf
                                strEntry = strEntry & vbTab & vbTab & "<rl_control>joy4way</rl_control>" & vbCrLf
                            strEntry = strEntry & vbTab & "</game>" & vbCrLf
                            strXML = strXML & strEntry
                            lngNumExported = lngNumExported + 1

                            'screenshots
                            If blnDoScreenshots = True And !ScrnShotFilename <> "" Then
                                For lngScrnShotIdx = 0 To (g_lngTotalScrnShotPaths - 1)
                                    strScreenshotSource = ga_udtRegPathScrnShot(lngScrnShotIdx).strWithBackSlash & !ScrnShotFilename
                                    If g_xobjFunc.FileExists(strScreenshotSource, False, "") Then
                                        'found screenshot, copy and rename to dest folder
                                        lngDotPos = InStrRev(strRomName, ".")
                                        If lngDotPos > 0 Then
                                            strRomNameNoExt = Left$(strRomName, lngDotPos - 1)
                                        Else
                                            strRomNameNoExt = strRomName
                                        End If
                                        strExtension = Right$(!ScrnShotFilename, Len(!ScrnShotFilename) - InStrRev(!ScrnShotFilename, "."))
                                        Call CopyFile(strScreenshotSource, strScreenshotDestPathWBS & strRomNameNoExt & "." & strExtension, False)
                                        Exit For
                                    End If
                                Next lngScrnShotIdx
                            End If

                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing

                End If
                If m_blnCancelled = True Then
                    Exit For
                End If
            Next LItem
            
            strXML = strXML & "</mame>" & vbCrLf
            
            If m_blnCancelled = False Then
                pbStatus.Value = pbStatus.Max
                DoEvents
                If lngNumExported > 0 Then
                    'write the new file
                    intFileNum = FreeFile()
                    Open strOutputFile For Output As #intFileNum
                    Print #intFileNum, strXML
                    Close #intFileNum
                    'show the file location
                    Screen.MousePointer = vbDefault
                    MsgBox "Exported successfully " & lngNumExported & " games to: -" & vbCrLf & vbCrLf & strOutputFile, vbInformation, gc_strAPPNAME
                    Unload Me
                Else
                    Screen.MousePointer = vbDefault
                    MsgBox "None Exported - only games with filenames/roms can be exported", vbInformation, gc_strAPPNAME
                End If
            End If
        End If
    End With

    Screen.MousePointer = vbDefault
    pbStatus.Visible = False
    Exit Sub

Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
    pbStatus.Visible = False

End Sub
Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2895    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6975     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    cmdHelp.Caption = g_strButtonHelpText
    With g_xobjFunc
        'with shortcuts
        optRomName(0).Caption = .GetINIMessage("RomLister Window", "GameIDOption", "Game &ID")
        optRomName(1).Caption = .GetINIMessage("RomLister Window", "GameFilenameOption", "Game &Filename")
        chkScreenshots.Caption = .GetINIMessage("RomLister Window", "ConvertCheckbox", "Convert &Screenshots")
        chkExt.Caption = .GetINIMessage("RomLister Window", "FilenameExtensionCheckbox", "with filename &extension")
        lblPrompt.Caption = .GetINIMessage("RomLister Window", "BrowseLabel", "&Put converted screenshots here")
        lblPrompt2.Caption = .GetINIMessage("RomLister Window", "RomNameLabel", "&Rom Name") & ":"
        'without shortcuts
        Caption = .GetINIMessage("RomLister Window", "Caption", "Export to RomLister")
    End With

    'load the last selected dest screenshot path
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "RomListerDestScreenshotPath", m_strDestScreenshotPath, True, g_strAppPathWithBackslash & "RomLister"
    lblDestScreenshotPath.Caption = m_strDestScreenshotPath
    
    'not started yet
    m_blnIsExporting = False

End Sub


Private Sub Form_Unload(Cancel As Integer)

    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "RomListerDestScreenshotPath", m_strDestScreenshotPath, True

End Sub


Private Sub lblDestScreenshotPath_Click()

    'get an initial valid path to select
    If g_xobjFunc.DirExists(lblDestScreenshotPath) Then
        m_strDestScreenshotPath = lblDestScreenshotPath
    ElseIf Not g_xobjFunc.DirExists(m_strDestScreenshotPath) Then
        m_strDestScreenshotPath = App.Path
    End If
    'get the new path
    g_strBrowseForFolder = m_strDestScreenshotPath
    frmBrowseFolder.Show vbModal, Me
    
    'change the folder in the listbox if path is selected and save the last traversed path
    If g_strBrowseForFolder <> "" Then
        lblDestScreenshotPath = g_strBrowseForFolder
        m_strDestScreenshotPath = g_strBrowseForFolder
    End If

End Sub


Private Sub optRomName_Click(Index As Integer)

    chkExt.Enabled = (Index = 1)

End Sub


