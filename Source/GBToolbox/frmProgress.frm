VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importing RomCenter 2.00 Datafile..."
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1215
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Initialising..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2010 James Burrows
' Email: gamebase.frontend@leesti.com
' Homepage: http://www.bu22.com/
'----------------------------------------------
'                The GBToolbox
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

Private m_blnActivated As Boolean
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_strDefaultTypes As String

Private Const cDELIM = "¬"

Private Type DataFileEntriesType
    strName As String
    strFile As String
    strDBFile As String
    strCRC As String
    lngSize As Long
    blnFA As Boolean
    strComment As String
    strVersionComment As String
    lngYearID As Long
    strPublisher As String
    strLanguage As String
    lngPAL_NTSC As Long
    strCracker As String
    lngTrainers As Long
    blnHSS As Boolean
End Type
Private ma_udtEntry() As DataFileEntriesType
Private m_lngTotalEntries As Long
Private m_lngMaxEntries As Long

Private Type FilesFoundType
    strFilename As String
    lngFilesize As Long
    strCRC As String
End Type
Private ma_udtFiles() As FilesFoundType
Private m_lngTotalFiles As Long
Private m_lngMaxFiles As Long

Private ma_strFileTypes() As String
Private m_lngTotalFileTypes As Long
Private m_lngMaxFileTypes As Long

Private m_CRC As clsCRC

Private m_blnCancelPressed As Boolean

Private Type YearTestType
    strYear As String
    lngID As Long
End Type
Private Function CRC32_and_Size_Compare() As String

    'match files as we know there is at least one to check

    Dim lngFileIdx As Long
    Dim lngEntryIdx As Long
    Dim lngMatched As Long
    Dim lngLenRomPath As Long
    Dim lngPos As Long
    Dim strExt As String
    
    On Error GoTo Err_Handler

    'init vars
    CRC32_and_Size_Compare = ""

    'get CRC-32 for each file
    lblPrompt = "Grabbing CRC-32 values for " & m_lngTotalFiles & " files..."
    DoEvents
    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxFiles
    Set m_CRC = New clsCRC
    For lngFileIdx = 0 To m_lngMaxFiles
        ma_udtFiles(lngFileIdx).strCRC = UCase$(Hex(m_CRC.CalculateFile(ma_udtFiles(lngFileIdx).strFilename)))
        pbStatus.Value = lngFileIdx
        DoEvents
        If m_blnCancelPressed = True Then Exit For
    Next lngFileIdx
    Set m_CRC = Nothing

    If m_blnCancelPressed = True Then Exit Function

    If g_udtRC2ImportOptions.blnMatch = False Then
        m_strDefaultTypes = ";"
    End If

    'compare
    lblPrompt = "CRC-32/Size Matches found:"
    lblStatus = "0"
    DoEvents
    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxEntries
    lngMatched = 0
    lngLenRomPath = Len(g_udtRC2ImportOptions.strRomPathWBS)
    For lngEntryIdx = 0 To m_lngMaxEntries
        With ma_udtEntry(lngEntryIdx)
            'add to filetypes if necessary
            If g_udtRC2ImportOptions.blnMatch = False Then
                lngPos = InStr(.strFile, ".")
                If lngPos > 0 Then
                    strExt = LCase$(Right$(.strFile, Len(.strFile) - lngPos))
                    If InStr(m_strDefaultTypes, ";" & strExt & ";") = 0 Then
                        m_strDefaultTypes = m_strDefaultTypes & strExt & ";"
                    End If
                End If
            End If
            'compare
            For lngFileIdx = 0 To m_lngMaxFiles
                If ma_udtFiles(lngFileIdx).strCRC = .strCRC Then
                    If ma_udtFiles(lngFileIdx).lngFilesize = .lngSize Then
                        'match found! - save relative path/filename
                        .blnFA = True
                        .strDBFile = Right$(ma_udtFiles(lngFileIdx).strFilename, Len(ma_udtFiles(lngFileIdx).strFilename) - lngLenRomPath)
                        lngMatched = lngMatched + 1
                        lblStatus = lngMatched
                        Exit For
                    End If
                End If
            Next lngFileIdx
        End With
        pbStatus.Value = lngEntryIdx
        DoEvents
        If m_blnCancelPressed = True Then Exit For
    Next lngEntryIdx

    If g_udtRC2ImportOptions.blnMatch = False Then
        m_strDefaultTypes = Right$(m_strDefaultTypes, Len(m_strDefaultTypes) - 1)
    End If

    'exit!
    Exit Function

Err_Handler:
    CRC32_and_Size_Compare = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
    Set m_CRC = Nothing

End Function

Private Function CreateNewDB(strDBFile_out As String) As String

    Dim strFolder As String
    Dim strGBName As String
    Dim strTextFile As String
    
    On Error GoTo Err_Handler

    Set m_xobjFunc = New GBFuncs.clsFunctions
    CreateNewDB = ""
    strFolder = g_udtRC2ImportOptions.strGameBaseFolder
    strGBName = g_udtRC2ImportOptions.strGameBaseName

    If m_xobjFunc.FileExists(g_strAppPathWithBackSlash & "Empty.mdb", False, "GB-Toolbox") = True Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Call MkDir(strFolder)
        Call MkDir(strFolder & "\Gfx")
        Call MkDir(strFolder & "\Scripts")
        Call MkDir(strFolder & "\Sounds")
        Call MkDir(strFolder & "\Splash")
        strDBFile_out = strFolder & "\" & strGBName & ".mdb"
        Call FileCopy(g_strAppPathWithBackSlash & "Empty.mdb", strDBFile_out)
        Call CreateFile(strFolder & "\Credits.txt", "Created by GBToolbox v" & Trim$(App.Major) & "." & Trim$(App.Minor) & " on " & Format$(Now, "dddd dd mmmm yyyy") & " by Unknown!" & vbCrLf & vbCrLf & "Imported RomCenter v2.00 Datafile: -" & vbCrLf & frmRomCenter2.lblDataFile & vbCrLf & vbCrLf & "[CREDITS]" & vbCrLf & "Author=" & g_strRC2Author & vbCrLf & "Version=" & g_strRC2Version & vbCrLf & "Comment=" & g_strRC2Comment & vbCrLf)
        strTextFile = GetResourceFileAsString(101)
        Call CreateFile(strFolder & "\Text.ini", strTextFile)
        strTextFile = GetResourceFileAsString(102)
        Call CreateFile(strFolder & "\Gfx\gfx.txt", strTextFile)
        Call CreateFile(strFolder & "\Sounds\sounds.txt", "Place your own Classic Game WAV files here.  See the GameBase documentation for more details.")
        Call CreateFile(strFolder & "\Splash\Splash.txt", "Place your own splash screen(s) and startup WAV files here.  See the GameBase documentation for more details.")
        If (g_udtRC2ImportOptions.blnScan = True) And (g_udtRC2ImportOptions.strRomPath <> "") Then
            Call m_xobjFunc.INI_WriteSetting("Games", "1", g_udtRC2ImportOptions.strRomPath, strFolder & "\Paths.ini")
            Call m_xobjFunc.INI_WriteSetting("File Types", "SupportTypes", m_strDefaultTypes, strFolder & "\Emulators.ini")
        End If
    Else
        CreateNewDB = "Could not find Empty.mdb in the GameBase installation folder." & vbCrLf & vbCrLf & "You can download another Empty.mdb file from the GameBase web site."
    End If

    Set m_xobjFunc = Nothing
    Exit Function

Err_Handler:
    
    Set m_xobjFunc = Nothing
    CreateNewDB = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Sub DoImport()

    Dim strError As String
    Dim blnDoScan As Boolean
    Dim strDBFile As String

    m_strDefaultTypes = ""

    If m_blnCancelPressed = True Then
        Unload Me
    End If

    'load the datafile into memory
    lblStatus = "0"
    If g_udtRC2ImportOptions.blnIgnoreDups = False Then
        lblPrompt = "Loading Datafile into memory..."
        DoEvents
        strError = ReadInFile()
    Else
        lblPrompt = "Loading Datafile into memory (skipping duplicates)..."
        DoEvents
        strError = ReadInFileDups()
    End If

    If m_blnCancelPressed = True Then Exit Sub

    If strError <> "" Then
        MsgBox strError, vbInformation, "GBToolbox"
        Exit Sub
    End If

    'scan if required
    lblStatus = ""
    If g_udtRC2ImportOptions.blnScan = True Then
        'get the unique file extensions
        If g_udtRC2ImportOptions.blnMatch = True Then
            lblPrompt = "Getting unique FileTypes list..."
            DoEvents
            If m_blnCancelPressed = True Then Exit Sub
            blnDoScan = GetUniqueFileTypes()
            If m_blnCancelPressed = True Then Exit Sub
        Else
            blnDoScan = True
        End If
        If blnDoScan = True Then
            'grab the list of files to CRC/Size check
            lblPrompt = "Preparing File List..."
            DoEvents
            If m_blnCancelPressed = True Then Exit Sub
            strError = GetFiles()
            If m_blnCancelPressed = True Then Exit Sub
            If strError <> "" Then
                MsgBox strError, vbInformation, "GBToolbox"
                Exit Sub
            End If
            'CRC/Size check them
            If m_lngTotalFiles > 0 Then
                If m_blnCancelPressed = True Then Exit Sub
                strError = CRC32_and_Size_Compare()
                If m_blnCancelPressed = True Then Exit Sub
                If strError <> "" Then
                    MsgBox strError, vbInformation, "GBToolbox"
                    Exit Sub
                End If
            End If
        End If
    End If

    If m_blnCancelPressed = True Then Exit Sub

    'free the mem - file array no longer needed
    ReDim ma_udtFiles(0)
    m_lngTotalFiles = 0
    m_lngMaxFiles = 0

    'perform TOSEC 1.0 parsing to get extra DB info from the filename
    If g_udtRC2ImportOptions.blnTosec = True Then
        If m_blnCancelPressed = True Then Exit Sub
        lblPrompt = "Obtaining DB info from TOSEC filenames..."
        lblStatus = ""
        DoEvents
        strError = GetTOSECv1Data()
        If m_blnCancelPressed = True Then Exit Sub
        If strError <> "" Then
            MsgBox strError, vbInformation, "GBToolbox"
            Exit Sub
        End If
    End If

    'now create the new database!
    cmdCancel.Enabled = False
    
    If g_blnRC2ExistingDB = False Then
        lblPrompt = "Creating New Database..."
        DoEvents
        strError = CreateNewDB(strDBFile)
        If strError <> "" Then
            MsgBox strError, vbInformation, "GBToolbox"
            Exit Sub
        End If
    Else
        strDBFile = g_strRC2DBFile
    End If
    
    'update with dat file info
    g_blnRC2Imported = True
    g_strRC2GBFolderWBS = g_udtRC2ImportOptions.strGameBaseFolderWBS
    lblPrompt = "Adding games to database..."
    DoEvents
    If g_udtRC2ImportOptions.blnTosec = False Then
        strError = ImportNoTOSEC(strDBFile)
    Else
        strError = ImportTOSEC(strDBFile)
    End If
    If strError <> "" Then
        MsgBox strError, vbInformation, "GBToolbox"

    End If

End Sub
Private Function GetFiles() As String

    'gets a list of files to search in the scan

    Dim strRomPath As String

    On Error GoTo Err_Handler

    'init vars
    GetFiles = ""
    m_lngTotalFiles = 0
    strRomPath = g_udtRC2ImportOptions.strRomPath

    'fill the array
    If g_udtRC2ImportOptions.blnMatch = False Then
        Call FindFiles(strRomPath)
    Else
        Call FindFilesMatch(strRomPath)
    End If

    If m_blnCancelPressed = True Then Exit Function

    'get proper values
    If m_lngTotalFiles > 0 Then
        m_lngMaxFiles = UBound(ma_udtFiles)
        m_lngTotalFiles = m_lngMaxFiles + 1
    End If

    'exit!
    Exit Function

Err_Handler:
    GetFiles = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Function GetTOSECv1Data() As String

    'gets any data it can from the TOSEC filename of each game

    Dim lngEntryIdx As Long
    Dim strFile As String
    Dim strName As String
    Dim lngPos As Long
    Dim lngPos2 As Long
    Dim strLower As String
    Dim lngLen As Long
    Dim strTemp As String
    Const cMAXLANG = 15
    Const cMAXYEARS = 48
    Dim a_strShortLang(cMAXLANG) As String
    Dim a_strLongLang(cMAXLANG) As String
    Dim lngIdx As Long
    Dim lngIdx2 As Long
    Dim blnFound As Boolean
    Dim a_udtYears(cMAXYEARS) As YearTestType

    On Error GoTo Err_Handler

    'init vars
    GetTOSECv1Data = ""
    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxEntries
    pbStatus.Value = 0
    DoEvents
    a_strShortLang(0) = "us"
    a_strShortLang(1) = "jp"
    a_strShortLang(2) = "uk"
    a_strShortLang(3) = "en"
    a_strShortLang(4) = "fr"
    a_strShortLang(5) = "de"
    a_strShortLang(6) = "sp"
    a_strShortLang(7) = "nl"
    a_strShortLang(8) = "it"
    a_strShortLang(9) = "cz"
    a_strShortLang(10) = "sw"
    a_strShortLang(11) = "hk"
    a_strShortLang(12) = "au"
    a_strShortLang(13) = "pl"
    a_strShortLang(14) = "br"
    a_strShortLang(15) = "eu"
    a_strLongLang(0) = "United States (English)"
    a_strLongLang(1) = "Japanese"
    a_strLongLang(2) = "United Kingdom (English)"
    a_strLongLang(3) = "English"
    a_strLongLang(4) = "French"
    a_strLongLang(5) = "German"
    a_strLongLang(6) = "Spanish"
    a_strLongLang(7) = "Dutch"
    a_strLongLang(8) = "Italian"
    a_strLongLang(9) = "Czech"
    a_strLongLang(10) = "Swedish"
    a_strLongLang(11) = "Hong Kong"
    a_strLongLang(12) = "Australian"
    a_strLongLang(13) = "Polish"
    a_strLongLang(14) = "Brazilian"
    a_strLongLang(15) = "European"

    a_udtYears(0).strYear = "(19xx)"
    a_udtYears(1).strYear = "(197x)"
    a_udtYears(2).strYear = "(198x)"
    a_udtYears(3).strYear = "(199x)"
    a_udtYears(4).strYear = "(20xx)"
    a_udtYears(5).strYear = "(200x)"
    a_udtYears(6).strYear = "(201x)"
    a_udtYears(7).strYear = "(202x)"
    a_udtYears(0).lngID = 2
    a_udtYears(1).lngID = 3
    a_udtYears(2).lngID = 4
    a_udtYears(3).lngID = 5
    a_udtYears(4).lngID = 6
    a_udtYears(5).lngID = 7
    a_udtYears(6).lngID = 8
    a_udtYears(7).lngID = 9
    For lngIdx = 8 To 48
        a_udtYears(lngIdx).strYear = "(" & (lngIdx + 1962) & ")"
        a_udtYears(lngIdx).lngID = lngIdx + 2
    Next lngIdx
    

    For lngEntryIdx = 0 To m_lngMaxEntries
        'parse any info we can
        With ma_udtEntry(lngEntryIdx)
            'cache for speed
            strFile = .strFile
            strName = .strName
            If Trim$(strName <> "") Then
            
                '===============
                'skip "ZZZ-UNK-"
                '===============

                If UCase$(Left$(strFile, 8)) = "ZZZ-UNK-" Then
                    strFile = Right$(strFile, Len(strFile) - 8)
                End If
                If UCase$(Left$(strName, 8)) = "ZZZ-UNK-" Then
                    strName = Right$(strName, Len(strName) - 8)
                End If
                Call Trim$(strFile)
                Call Trim$(strName)

                '=================================================
                'set the name - strip (year)(publisher) TOSEC info
                '=================================================

                If strName <> "" Then
                    lngPos = InStr(strName, "(")
                    If lngPos > 0 Then
                        'skip (demo)|(preview)|(playable preview)
                        If (LCase$(Mid$(strName, lngPos, 6)) = "(demo)") Or (LCase$(Mid$(strName, lngPos, 9)) = "(preview)") Or (LCase$(Mid$(strName, lngPos, 10)) = "(playable ") Then
                            lngPos2 = InStr(lngPos + 1, strName, ")")
                            If lngPos2 > 0 Then
                                .strName = Trim$(Left$(strName, lngPos2))
                                lngPos = lngPos2
                            Else
                                .strName = strName
                            End If
                        Else
                            .strName = Trim$(Left$(strName, lngPos - 1))
                        End If
                        'name now contains just the name
                        'skip year and publisher
                        lngPos = InStr(lngPos, strName, "(")
                        If lngPos > 0 Then
                            lngPos2 = InStr(lngPos + 1, strName, "(")
                            If lngPos2 > 0 Then
                                lngPos2 = InStr(lngPos2 + 1, strName, ")")
                                If lngPos2 > 0 Then
                                    .strName = .strName & " " & Right$(strName, Len(strName) - lngPos2)
                                End If
                            End If
                        End If
                    Else
                        .strName = strName
                    End If
                Else
                    .strName = "ZZZ-UNK-"
                End If

                '=======================================================
                'ok - we have a name, now check the filename for db info
                '=======================================================

                strLower = LCase$(strFile)
                'high score saver
                .blnHSS = (InStr(strLower, "[f save]") > 0)
                .strName = Replace(.strName, "[f save]", "", , , vbTextCompare)
                'pal/ntsc
                If InStr(strLower, "(pal)") > 0 Then
                    If InStr(strLower, "[f ntsc]") > 0 Then
                        .lngPAL_NTSC = 1
                    Else
                        .lngPAL_NTSC = 0
                    End If
                ElseIf InStr(strLower, "(ntsc)") > 0 Then
                    .lngPAL_NTSC = 2
                Else
                    If InStr(strLower, "[f ntsc]") > 0 Then
                        .lngPAL_NTSC = 1
                    Else
                        .lngPAL_NTSC = 3
                    End If
                End If
                .strName = Replace(.strName, "(pal)", "", , , vbTextCompare)
                .strName = Replace(.strName, "(ntsc)", "", , , vbTextCompare)
                .strName = Replace(.strName, "[f ntsc]", "", , , vbTextCompare)
                'comment
                If InStr(strLower, "(pd)") > 0 Then
                    .strComment = "Public Domain"
                Else
                    .strComment = ""
                End If
                .strName = Replace(.strName, "(pd)", "", , , vbTextCompare)
                'cracker
                .strVersionComment = ""
                lngPos = InStr(strLower, "[cr")
                If lngPos > 0 Then
                    lngPos2 = InStr(lngPos, strLower, "]")
                    If lngPos2 > 0 Then
                        If lngPos2 = (lngPos + 3) Then
                            .strCracker = ""
                            .strVersionComment = "[cracked by ?]"
                            .strName = Replace(.strName, "[cr]", "", , , vbTextCompare)
                        Else
                            .strCracker = Trim$(Mid$(strFile, lngPos + 3, (lngPos2 - lngPos) - 3))
                            .strName = Replace(.strName, "[cr " & .strCracker & "]", "", , , vbTextCompare)
                        End If
                    Else
                        .strCracker = "~NONE~"
                    End If
                Else
                    .strCracker = "~NONE~"
                End If
                'num trainers
                lngPos = InStr(strLower, "[t")
                If lngPos > 0 Then
                    lngPos2 = InStr(lngPos, strLower, "]")
                    If lngPos2 > 0 Then
                        If lngPos2 = (lngPos + 2) Then
                            .lngTrainers = -1
                            .strVersionComment = .strVersionComment & " [trained]"
                            .strName = Replace(.strName, "[t]", "", , , vbTextCompare)
                        Else
                            strTemp = Trim$(Mid$(strFile, lngPos + 3, (lngPos2 - lngPos) - 3))
                            .strName = Replace(.strName, "[t " & strTemp & "]", "", , , vbTextCompare)
                            If Left$(strTemp, 1) = "+" Then
                                .lngTrainers = Val(Right$(strTemp, Len(strTemp) - 1))
                                If .lngTrainers < -1 Then
                                    .lngTrainers = -1
                                End If
                            Else
                                .lngTrainers = -1
                                .strVersionComment = .strVersionComment & " [trained]"
                            End If
                        End If
                    Else
                        .lngTrainers = 0
                    End If
                Else
                    .lngTrainers = 0
                End If
                'language
                .strLanguage = ""
                For lngIdx = 0 To cMAXLANG
                    If InStr(strLower, "(" & a_strShortLang(lngIdx) & ")") > 0 Then
                        .strLanguage = a_strLongLang(lngIdx)
                        .strName = Replace(.strName, "(" & a_strShortLang(lngIdx) & ")", "", , , vbTextCompare)
                        Exit For
                    End If
                Next lngIdx
                If .strLanguage = "" Then
                    For lngIdx = 3 To 6
                        If InStr(strLower, "(m" & lngIdx & ")") > 0 Then
                            .strLanguage = "Multi-Language (" & lngIdx & ")"
                            .strName = Replace(.strName, "(m" & lngIdx & ")", "", , , vbTextCompare)
                            Exit For
                        End If
                    Next lngIdx
                End If
                If .strLanguage = "" Then
                    For lngIdx = 0 To cMAXLANG
                        lngPos = InStr(strLower, "(" & a_strShortLang(lngIdx) & "-")
                        If lngPos > 0 Then
                            lngPos2 = InStr(lngPos, strLower, ")")
                            If lngPos2 = (lngPos + 6) Then
                                strTemp = Trim$(Mid$(strFile, lngPos + 4, (lngPos2 - lngPos) - 4))
                                blnFound = False
                                For lngIdx2 = 0 To cMAXLANG
                                    If a_strShortLang(lngIdx2) = LCase$(strTemp) Then
                                        .strLanguage = a_strLongLang(lngIdx) & " / " & a_strLongLang(lngIdx2)
                                        .strName = Replace(.strName, "(" & a_strShortLang(lngIdx) & "-" & a_strShortLang(lngIdx2) & ")", "", , , vbTextCompare)
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngIdx2
                                If blnFound = False Then
                                    .strLanguage = a_strLongLang(lngIdx) & " / " & Trim$(Left$(strTemp, 128))
                                    .strName = Replace(.strName, Trim$(Mid$(strFile, lngPos, (lngPos2 - lngPos) + 1)), "", , , vbTextCompare)
                                End If
                            ElseIf lngPos2 = (lngPos + 4) Then
                                .strLanguage = a_strLongLang(lngIdx)
                                .strName = Replace(.strName, Trim$(Mid$(strFile, lngPos, (lngPos2 - lngPos) + 1)), "", , , vbTextCompare)
                                Exit For
                            ElseIf lngPos2 > 0 Then
                                .strLanguage = "Multi-Language (?)"
                                .strName = Replace(.strName, Trim$(Mid$(strFile, lngPos, (lngPos2 - lngPos) + 1)), "", , , vbTextCompare)
                                Exit For
                            End If
                        End If
                    Next lngIdx
                End If
                If .strLanguage = "" Then
                    .strLanguage = "English"
                End If
                'year
                .lngYearID = 1
                For lngIdx = 0 To cMAXYEARS
                    lngPos = InStr(strLower, a_udtYears(lngIdx).strYear)
                    If lngPos > 0 Then
                        .lngYearID = a_udtYears(lngIdx).lngID
                        Exit For
                    End If
                Next lngIdx
                .strName = Replace(.strName, "(y)", "", , , vbTextCompare)
                'publisher
                .strPublisher = ""
                If lngPos > 0 Then
                    lngPos = InStr(lngPos + 1, strLower, "(")
                    If lngPos > 0 Then
                        lngPos2 = InStr(lngPos, strLower, ")")
                        If lngPos2 > (lngPos + 1) Then
                            .strPublisher = Trim$(Mid$(strFile, lngPos + 1, (lngPos2 - lngPos) - 1))
                            If .strPublisher = "-" Then
                                .strPublisher = ""
                            End If
                        End If
                    End If
                End If
                .strName = Replace(.strName, "(-)", "", , , vbTextCompare)
            Else
                'set to nothing so we ignore it in the database write
                .strName = ""
            End If
        End With
        'update ui and get events
        pbStatus.Value = lngEntryIdx
        DoEvents
        If m_blnCancelPressed = True Then Exit For
    Next lngEntryIdx

    'exit!
    Exit Function

Err_Handler:
    GetTOSECv1Data = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Function GetUniqueFileTypes() As Boolean

    'gets the unique filetypes into an array
    'returns true if at least one filetype was found
    'returns false otherwise

    Dim lngEntryIdx As Long
    Dim lngFileTypeIdx As Long
    Dim blnFound As Boolean
    Dim strFileType As String
    Dim lngCheckMax As Long
    Dim strFile As String
    Dim lngDotPos As Long

    If m_lngTotalEntries = 0 Then
        GetUniqueFileTypes = False
    End If
    m_lngTotalFileTypes = 0

    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxEntries

    For lngEntryIdx = 0 To m_lngMaxEntries
        blnFound = False
        'get the filetype
        strFile = ma_udtEntry(lngEntryIdx).strFile
        lngDotPos = InStrRev(strFile, ".")
        If lngDotPos > 0 Then
            strFileType = UCase$(Right$(strFile, Len(strFile) - lngDotPos + 1))
            'check it's unique
            lngCheckMax = m_lngTotalFileTypes - 1
            If m_lngTotalFileTypes > 0 Then
                For lngFileTypeIdx = 0 To lngCheckMax
                    If strFileType = ma_strFileTypes(lngFileTypeIdx) Then
                        blnFound = True
                        Exit For
                    End If
                Next lngFileTypeIdx
            End If
        Else
            'no filename extension - just ignore it
            blnFound = True
        End If
        If blnFound = False Then
            'add to filetypes
            ReDim Preserve ma_strFileTypes(m_lngTotalFileTypes)
            ma_strFileTypes(m_lngTotalFileTypes) = strFileType
            m_lngTotalFileTypes = m_lngTotalFileTypes + 1
            m_strDefaultTypes = m_strDefaultTypes & strFileType & ";"
        End If
        'update statusbar
        pbStatus.Value = lngEntryIdx
        DoEvents
        If m_blnCancelPressed = True Then Exit For
    Next lngEntryIdx

    If m_blnCancelPressed = True Then Exit Function

    If m_lngTotalFileTypes > 0 Then
        GetUniqueFileTypes = True
        m_lngMaxFileTypes = UBound(ma_strFileTypes)
        m_lngTotalFileTypes = m_lngMaxFileTypes + 1
    Else
        GetUniqueFileTypes = False
    End If

End Function

Private Function ImportNoTOSEC(strDBFile_in As String) As String

    Dim lngEntryIdx As Long
    Dim dbGameBase As DAO.Database
    Dim rsGameBase As DAO.Recordset
    Dim lngTestVar As Long
    Dim sngDatabaseVersion As Single

    On Error GoTo Err_Handler
    
    'init vars
    ImportNoTOSEC = ""
    lngTestVar = 0
    Set dbGameBase = OpenDatabase(strDBFile_in)
    lngTestVar = 1
    Set rsGameBase = dbGameBase.OpenRecordset("SELECT * FROM Games", dbOpenDynaset)
    lngTestVar = 2
    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxEntries
    pbStatus.Value = 0
    
    sngDatabaseVersion = GetDatabaseVersion(strDBFile_in)
    
    'loop through and add all the games to the Games table
    For lngEntryIdx = 0 To m_lngMaxEntries
        If ma_udtEntry(lngEntryIdx).strName <> "" Then
            With rsGameBase
                .AddNew
                !Name = Left$(ma_udtEntry(lngEntryIdx).strName, 255)
                !YE_ID = 1
                !Comment = ""
                !FileToRun = ""
                !FilenameIndex = 0
                !ScrnShotFilename = ""
                !MU_Id = 38
                !GE_Id = 1
                !PU_Id = 37
                !DI_Id = 8
                !CR_Id = 8
                !SidFileName = ""
                !DateLastPlayed = ""
                !TimesPlayed = 0
                !HighScore = ""
                !SA = False
                If g_udtRC2ImportOptions.blnScan = True Then
                    !Filename = ma_udtEntry(lngEntryIdx).strDBFile
                    !CCode = IIf((ma_udtEntry(lngEntryIdx).blnFA = False), 4, 2)
                    !FA = ma_udtEntry(lngEntryIdx).blnFA
                Else
                    !Filename = ma_udtEntry(lngEntryIdx).strFile
                    !CCode = 4
                    !FA = False
                End If
                !CRC = ma_udtEntry(lngEntryIdx).strCRC
                !Filesize = ma_udtEntry(lngEntryIdx).lngSize
                !Fav = False
                !PR_Id = 1
                !LA_Id = 1
                !Extras = False
                !Classic = False
                !Rating = 0
                !V_LoadingScreen = 0
                !V_HighScoreSaver = 0
                !V_IncludedDocs = 0
                !V_PalNTSC = 1
                !V_TrueDriveEmu = 0
                !V_Length = 0
                If sngDatabaseVersion >= 2.5 Then
                    !V_LengthType = 19
                End If
                !V_Trainers = -1
                !PlayersFrom = -1
                !PlayersTo = -1
                !PlayersSim = False
                !V_Comment = ""
                !Adult = False
                !MemoText = ""
                !Prequel = 0
                !Sequel = 0
                !Related = 0
                !Control = 0
                !AR_Id = 1
                !DE_Id = 1
                !LI_Id = 1
                !RA_Id = 1
                !WebLink_Name = ""
                !WebLink_URL = ""
                !V_WebLink_Name = ""
                !V_WebLink_URL = ""
                !V_TitleScreen = 0
                !V_Playable = 0
                !V_Original = 0
                !CloneOf = 0
                .Update
            End With
        End If
        pbStatus.Value = lngEntryIdx
        DoEvents
    Next lngEntryIdx
    
    'exit!
    rsGameBase.Close
    dbGameBase.Close
    Set rsGameBase = Nothing
    Set dbGameBase = Nothing
    Exit Function

Err_Handler:
    
    If lngTestVar > 1 Then
        rsGameBase.Close
    End If
    If lngTestVar > 0 Then
        dbGameBase.Close
    End If
    Set rsGameBase = Nothing
    Set dbGameBase = Nothing
    ImportNoTOSEC = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Function ImportTOSEC(strDBFile_in As String) As String

    Dim lngEntryIdx As Long
    Dim dbGameBase As DAO.Database
    Dim rsGameBase As DAO.Recordset
    Dim rsLink As DAO.Recordset
    Dim lngTestVar As Long
    Dim lngCR_Id As Long
    Dim lngLA_Id As Long
    Dim lngPU_Id As Long

    On Error GoTo Err_Handler
    
    'init vars
    ImportTOSEC = ""
    lngTestVar = 0
    Set dbGameBase = OpenDatabase(strDBFile_in)
    lngTestVar = 1
    Set rsGameBase = dbGameBase.OpenRecordset("SELECT * FROM Games", dbOpenDynaset)
    lngTestVar = 2
    pbStatus.Min = 0
    pbStatus.Max = m_lngMaxEntries
    pbStatus.Value = 0
    If g_blnRC2ExistingDB = False Then
        dbGameBase.Execute "DELETE * FROM Languages WHERE LA_Id > 1"
    End If

    'loop through and add all the games to the Games table
    For lngEntryIdx = 0 To m_lngMaxEntries
        If ma_udtEntry(lngEntryIdx).strName <> "" Then
            'set/get the cracker
            ma_udtEntry(lngEntryIdx).strCracker = Left$(ma_udtEntry(lngEntryIdx).strCracker, 128)
            ma_udtEntry(lngEntryIdx).strPublisher = Left$(ma_udtEntry(lngEntryIdx).strPublisher, 128)
            ma_udtEntry(lngEntryIdx).strLanguage = Left$(ma_udtEntry(lngEntryIdx).strLanguage, 128)
            If ma_udtEntry(lngEntryIdx).strCracker = "" Then
                lngCR_Id = 8
            ElseIf ma_udtEntry(lngEntryIdx).strCracker = "~NONE~" Then
                lngCR_Id = 7
            Else
                Set rsLink = dbGameBase.OpenRecordset("SELECT * FROM Crackers WHERE Cracker = """ & ma_udtEntry(lngEntryIdx).strCracker & """", dbOpenDynaset)
                lngTestVar = 3
                With rsLink
                    If .EOF Then
                        .AddNew
                        lngCR_Id = !CR_Id
                        !Cracker = ma_udtEntry(lngEntryIdx).strCracker
                        .Update
                    Else
                        .MoveFirst
                        lngCR_Id = !CR_Id
                    End If
                    .Close
                    lngTestVar = 2
                End With
                Set rsLink = Nothing
            End If
            'set/get the publisher
            If ma_udtEntry(lngEntryIdx).strPublisher = "" Then
                lngPU_Id = 37
            Else
                Set rsLink = dbGameBase.OpenRecordset("SELECT * FROM Publishers WHERE Publisher = """ & ma_udtEntry(lngEntryIdx).strPublisher & """", dbOpenDynaset)
                lngTestVar = 3
                With rsLink
                    If .EOF Then
                        .AddNew
                        lngPU_Id = !PU_Id
                        !Publisher = ma_udtEntry(lngEntryIdx).strPublisher
                        .Update
                    Else
                        .MoveFirst
                        lngPU_Id = !PU_Id
                    End If
                    .Close
                    lngTestVar = 2
                End With
                Set rsLink = Nothing
            End If
            'set/get the language
            If ma_udtEntry(lngEntryIdx).strLanguage = "" Then
                lngLA_Id = 1
            Else
                Set rsLink = dbGameBase.OpenRecordset("SELECT * FROM Languages WHERE Language = """ & ma_udtEntry(lngEntryIdx).strLanguage & """", dbOpenDynaset)
                lngTestVar = 3
                With rsLink
                    If .EOF Then
                        .AddNew
                        lngLA_Id = !LA_Id
                        !Language = ma_udtEntry(lngEntryIdx).strLanguage
                        .Update
                    Else
                        .MoveFirst
                        lngLA_Id = !LA_Id
                    End If
                    .Close
                    lngTestVar = 2
                End With
                Set rsLink = Nothing
            End If
            'add the game
            With rsGameBase
                .AddNew
                !PU_Id = lngPU_Id
                !CR_Id = lngCR_Id
                !LA_Id = lngLA_Id
                !Name = Left$(ma_udtEntry(lngEntryIdx).strName, 255)
                !YE_ID = ma_udtEntry(lngEntryIdx).lngYearID
                !Comment = Left$(ma_udtEntry(lngEntryIdx).strComment, 125)
                !FileToRun = ""
                !FilenameIndex = 0
                !ScrnShotFilename = ""
                !MU_Id = 38
                !GE_Id = 1
                !DI_Id = 8
                !PR_Id = 1
                !SidFileName = ""
                !DateLastPlayed = ""
                !TimesPlayed = 0
                !HighScore = ""
                !SA = False
                If g_udtRC2ImportOptions.blnScan = True Then
                    !Filename = ma_udtEntry(lngEntryIdx).strDBFile
                    !CCode = IIf((ma_udtEntry(lngEntryIdx).blnFA = False), 4, 2)
                    !FA = ma_udtEntry(lngEntryIdx).blnFA
                Else
                    !Filename = ma_udtEntry(lngEntryIdx).strFile
                    !CCode = 4
                    !FA = False
                End If
                !CRC = ma_udtEntry(lngEntryIdx).strCRC
                !Filesize = ma_udtEntry(lngEntryIdx).lngSize
                !Fav = False
                !Extras = False
                !Classic = False
                !Rating = 0
                !V_LoadingScreen = False
                !V_HighScoreSaver = ma_udtEntry(lngEntryIdx).blnHSS
                !V_IncludedDocs = False
                !V_PalNTSC = ma_udtEntry(lngEntryIdx).lngPAL_NTSC
                !V_TrueDriveEmu = False
                !V_Length = 0
                !V_Trainers = ma_udtEntry(lngEntryIdx).lngTrainers
                !PlayersFrom = -1
                !PlayersTo = -1
                !PlayersSim = False
                !V_Comment = Left$(ma_udtEntry(lngEntryIdx).strVersionComment, 100)
                !Adult = False
                !MemoText = ""
                !Prequel = 0
                !Sequel = 0
                !Related = 0
                !Control = 0
                !AR_Id = 1
                !DE_Id = 1
                !LI_Id = 1
                !RA_Id = 1
                !WebLink_Name = ""
                !WebLink_URL = ""
                !V_WebLink_Name = ""
                !V_WebLink_URL = ""
                !V_TitleScreen = 0
                !V_Playable = 0
                !V_Original = 0
                !CloneOf = 0
                .Update
            End With
        End If
        pbStatus.Value = lngEntryIdx
        DoEvents
    Next lngEntryIdx
    
    'exit!
    rsGameBase.Close
    dbGameBase.Close
    Set rsGameBase = Nothing
    Set dbGameBase = Nothing
    Exit Function

Err_Handler:
    
    If lngTestVar > 2 Then
        rsLink.Close
    End If
    If lngTestVar > 1 Then
        rsGameBase.Close
    End If
    If lngTestVar > 0 Then
        dbGameBase.Close
    End If
    Set rsLink = Nothing
    Set rsGameBase = Nothing
    Set dbGameBase = Nothing
    ImportTOSEC = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function StripNulls(strSource_inout As String) As String
 
    '------------------------------------------------------------------------------
    'Used by FindFiles() to strip trailing NULL characters off the passed in string
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

Public Sub FindFiles(strPath_in As String)
    
    Dim strFilename As String           'Working filename variable
    Dim strDirName As String            'Working SubDirectory variable
    Dim a_strDirNames() As String       'Buffer for directory name entries
    Dim intSubDirCount As Integer       'Number of directories in current path
    Dim intSubDirIdx As Integer         'loop index
    Dim hSearch As Long                 'Search Handle
    Dim udtWFD As WIN32_FIND_DATA       'structure for API
    Dim intContinue As Integer          'continue searching flag
    Dim lngPos As Long
    Dim strExt As String
    
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
            If m_blnCancelPressed = True Then Exit Do
        Loop
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Walk through this directory
    hSearch = FindFirstFile(strPath_in & "*.*", udtWFD)
    intContinue = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While intContinue
            strFilename = StripNulls(udtWFD.cFileName)
            If (strFilename <> ".") And (strFilename <> "..") Then
                If Not (GetFileAttributes(strPath_in & strFilename) And FILE_ATTRIBUTE_DIRECTORY) Then
                    'add to files
                    ReDim Preserve ma_udtFiles(m_lngTotalFiles)
                    With ma_udtFiles(m_lngTotalFiles)
                        .strFilename = strPath_in & strFilename
                        .lngFilesize = (udtWFD.nFileSizeHigh * MAXDWORD) + udtWFD.nFileSizeLow
                    End With
                    m_lngTotalFiles = m_lngTotalFiles + 1
                End If
            End If
            'Get next file
            intContinue = FindNextFile(hSearch, udtWFD)
            DoEvents
            If m_blnCancelPressed = True Then intContinue = False
        Wend
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Any sub-directories?
    If intSubDirCount > 0 Then
        'Recursively process them
        For intSubDirIdx = 0 To intSubDirCount - 1
            'add to accumulated filesize
            Call FindFiles(strPath_in & a_strDirNames(intSubDirIdx))
            DoEvents
            If m_blnCancelPressed = True Then Exit For
        Next intSubDirIdx
    End If
    
End Sub
Public Sub FindFilesMatch(strPath_in As String)
    
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
            If m_blnCancelPressed = True Then Exit Do
        Loop
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Walk through this directory
    For lngFTIdx = 0 To m_lngMaxFileTypes
        hSearch = FindFirstFile(strPath_in & "*" & ma_strFileTypes(lngFTIdx), udtWFD)
        intContinue = True
        If hSearch <> INVALID_HANDLE_VALUE Then
            While intContinue
                strFilename = StripNulls(udtWFD.cFileName)
                If (strFilename <> ".") And (strFilename <> "..") Then
                    If Not (GetFileAttributes(strPath_in & strFilename) And FILE_ATTRIBUTE_DIRECTORY) Then
                        'add to files
                        ReDim Preserve ma_udtFiles(m_lngTotalFiles)
                        With ma_udtFiles(m_lngTotalFiles)
                            .strFilename = strPath_in & strFilename
                            .lngFilesize = (udtWFD.nFileSizeHigh * MAXDWORD) + udtWFD.nFileSizeLow
                        End With
                        m_lngTotalFiles = m_lngTotalFiles + 1
                    End If
                End If
                'Get next file
                intContinue = FindNextFile(hSearch, udtWFD)
                DoEvents
                If m_blnCancelPressed = True Then intContinue = False
            Wend
            'tidy up
            intContinue = FindClose(hSearch)
        End If
        DoEvents
        If m_blnCancelPressed = True Then Exit For
    Next lngFTIdx
    
    
    'Any sub-directories?
    If intSubDirCount > 0 Then
        'Recursively process them
        For intSubDirIdx = 0 To intSubDirCount - 1
            'add to accumulated filesize
            Call FindFilesMatch(strPath_in & a_strDirNames(intSubDirIdx))
            DoEvents
            If m_blnCancelPressed = True Then Exit For
        Next intSubDirIdx
    End If

End Sub
Private Function ReadInFile() As String

    'read in the file without removing duplicates

    Dim intFileNum As Integer
    Dim strLine As String
    Dim a_strFields() As String
    Dim lngNameIdx As Long
    Dim lngFileIdx As Long
    Dim blnAddZip As Boolean

'    On Error GoTo Err_Handler

    'init vars
    ReadInFile = ""
    lngNameIdx = g_udtRC2ImportOptions.lngNameColIdx
    lngFileIdx = g_udtRC2ImportOptions.lngFileColIdx
    blnAddZip = g_udtRC2ImportOptions.blnAddZip
    
    'open the file
    intFileNum = FreeFile()
    Open g_strRC2DataFile For Input As #intFileNum

    'read in each line, test it and store the required data
    m_lngTotalEntries = 0
    Do Until EOF(intFileNum)
        Line Input #intFileNum, strLine
        If Left$(strLine, 1) = cDELIM Then
            a_strFields() = Split(strLine, cDELIM)
            If UBound(a_strFields) = 10 Then
                ReDim Preserve ma_udtEntry(m_lngTotalEntries)
                With ma_udtEntry(m_lngTotalEntries)
                    If a_strFields(7) <> "N" Then
                        .lngSize = a_strFields(7)
                    Else
                        .lngSize = a_strFields(9)
                    End If
                    .strCRC = UCase$(a_strFields(6))
                    .strFile = a_strFields(lngFileIdx)
                    If blnAddZip Then
                        .strFile = .strFile & ".zip"
                    End If
                    .strName = a_strFields(lngNameIdx)
                End With
                m_lngTotalEntries = m_lngTotalEntries + 1
                lblStatus = m_lngTotalEntries
                DoEvents
                If m_blnCancelPressed = True Then Exit Do
            End If
        End If
    Loop

    Close #intFileNum

    If m_blnCancelPressed = True Then
        Exit Function
    End If
    
    'exit
    If m_lngTotalEntries = 0 Then
        ReadInFile = "There were no games to import!"
    Else
        m_lngMaxEntries = UBound(ma_udtEntry)
        m_lngTotalEntries = m_lngMaxEntries + 1
    End If
    Exit Function
    
Err_Handler:
    ReadInFile = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
    If Err.Number = 13 Then
        ReadInFile = ReadInFile & vbCrLf & vbCrLf & "This may not be a valid v2.00 datafile, it may be a misreported v1.92 datafile!"
    End If

End Function

Private Function ReadInFileDups() As String

    'read in the file removing duplicates

    Dim intFileNum As Integer
    Dim strLine As String
    Dim a_strFields() As String
    Dim lngNameIdx As Long
    Dim lngFileIdx As Long
    Dim blnAddZip As Boolean
    Dim strName As String
    Dim blnFound As Boolean
    Dim lngIdx As Long
    Dim lngCheckMax As Long

    On Error GoTo Err_Handler

    'init vars
    ReadInFileDups = ""
    lngNameIdx = g_udtRC2ImportOptions.lngNameColIdx
    lngFileIdx = g_udtRC2ImportOptions.lngFileColIdx
    blnAddZip = g_udtRC2ImportOptions.blnAddZip
    
    'open the file
    intFileNum = FreeFile()
    Open g_strRC2DataFile For Input As #intFileNum

    'read in each line, test it and store the required data
    m_lngTotalEntries = 0
    Do Until EOF(intFileNum)
        Line Input #intFileNum, strLine
        If Left$(strLine, 1) = cDELIM Then
            a_strFields() = Split(strLine, cDELIM)
            If UBound(a_strFields) = 10 Then
                'check if name not already added
                strName = a_strFields(lngNameIdx)
                blnFound = False
                If m_lngTotalEntries > 0 Then
                    lngCheckMax = m_lngTotalEntries - 1
                    For lngIdx = 0 To lngCheckMax
                        If strName = ma_udtEntry(lngIdx).strName Then
                            blnFound = True
                            Exit For
                        End If
                    Next lngIdx
                End If
                'add if not duplicate
                If blnFound = False Then
                    ReDim Preserve ma_udtEntry(m_lngTotalEntries)
                    With ma_udtEntry(m_lngTotalEntries)
                        If a_strFields(7) <> "N" Then
                            .lngSize = a_strFields(7)
                        Else
                            .lngSize = a_strFields(9)
                        End If
                        .strCRC = UCase$(a_strFields(6))
                        .strFile = a_strFields(lngFileIdx)
                        If blnAddZip Then
                            .strFile = .strFile & ".zip"
                        End If
                        .strName = strName
                    End With
                    m_lngTotalEntries = m_lngTotalEntries + 1
                    lblStatus = m_lngTotalEntries
                    DoEvents
                    If m_blnCancelPressed = True Then Exit Do
                End If
            End If
        End If
    Loop

    Close #intFileNum

    If m_blnCancelPressed = True Then
        Exit Function
    End If

    'exit
    If m_lngTotalEntries = 0 Then
        ReadInFileDups = "There were no games to import!"
    Else
        m_lngMaxEntries = UBound(ma_udtEntry)
        m_lngTotalEntries = m_lngMaxEntries + 1
    End If
    Exit Function
    
Err_Handler:
    ReadInFileDups = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
    If Err.Number = 13 Then
        ReadInFileDups = ReadInFileDups & vbCrLf & vbCrLf & "This may not be a valid v2.00 datafile, it may be a misreported v1.92 datafile!"
    End If
    
End Function

Private Sub cmdCancel_Click()

    cmdCancel.Enabled = False
    m_blnCancelPressed = True

End Sub

Private Sub Form_Activate()

    If m_blnActivated = False Then
        m_blnActivated = True

        '=============
        'DO THE IMPORT
        '=============

        Call DoImport
        g_blnCancelPressed = m_blnCancelPressed
        Unload Me

    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 1215    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5175     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    g_blnCancelPressed = False
    m_blnCancelPressed = False
    Caption = "Importing " & frmRomCenter2.lblDataFile & "..."
    m_blnActivated = False
    Screen.MousePointer = vbHourglass
    DoEvents

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Screen.MousePointer = vbDefault

End Sub


