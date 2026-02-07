VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GameBase Launcher"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
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

Private Const m_blnInVBIDE As Boolean = False
Private Const mc_strAPP_PATH_WBS_WHEN_IN_VBIDE = "C:\GameBase\Build\"

Private Const mc_strLOGFILENAME = "GBLauncher.log"
Private m_intLogFileNum As Integer
Private m_blnIsDBOpen As Boolean
Private Enum LogStatusEnum
    e_do_force_quit
    e_do_allow_continue
End Enum
Private Type CLPType
    strDBName As String
    strGameID As String
    strGameFilename As String
    strEmuIdx As String
    blnWait As Boolean
End Type
Private m_udtCLP As CLPType

Private Const mc_strAPPNAME As String = "GameBase Launcher"
Private Const mc_strAUTHOR As String = "James Burrows"
Private Const mc_sngREQ_DB_VERSION As Single = 2.5

Private m_strAPPVERSION As String

Private m_sngDBVersion As Single
Private m_lngDBMajorVersion As Long
Private m_lngDBMinorVersion As Long
Private m_strErrorUpdatingDatabase As String
Private m_lngOfficialDBUpdate As Long

Private m_dbGamebase As Database
Private m_rsGames As Recordset

Private m_strAppPathWithBackslash As String
Private m_strGameBaseName As String
Private m_strGameBaseDBPath As String
Private m_strPathFile As String
Private m_strSettingsFile As String
Private m_strGameBaseDotINI As String

Private m_xobjFunc As GBFuncs.clsFunctions

Private m_blnReZip As Boolean
Private m_blnRezipAsk As Boolean
Private me_RezipAdditional As Gemus.EnumRezipAdditionalFiles
Private m_blnZipNotify As Boolean
Private m_blnRezipLegacy As Boolean
Private m_blnPromptForGamePathDevice As Boolean

Private Enum PathLocationEnum
    e_PathLocation_Removable
    e_PathLocation_FixedHD
    e_PathLocation_CD
    e_PathLocation_Invalid
End Enum
Private Type PathsType
    strWithBackSlash As String
    strRealPath As String
    ePathType As PathLocationEnum
    blnIsSet As Boolean
End Type
Private ma_udtRegPathGame() As PathsType
Private m_udtPathGBGame As PathsType
Private m_lngTotalGamePaths As Long
Private m_blnValidGamePathExists As Boolean

'Game data
Private Type GameDataType
    lngID As Long
    lngPlayersFrom As Long
    lngPlayersTo As Long
    blnGameFileExists As Boolean
    blnTrueDriveEmulation As Boolean
    lngNTSC As Long
    strComment As String
    strVersionComment As String
    strGameFilename As String
    strFileToRun As String
    lngGameFilenameIndex As Long
    lngControl As Long
    strPairs As String
End Type
Private m_udtGame As GameDataType

'for getting drive types
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_FIXED = 3
Private Const DRIVE_CDROM = 5

Private Function CreateDefaultGBGameFolder() As Boolean

    On Error GoTo Err_Handler

    Call MkDir(GetDefaultGBGameFolder())
    CreateDefaultGBGameFolder = m_xobjFunc.DirExists(GetDefaultGBGameFolder())
    Exit Function

Err_Handler:
    CreateDefaultGBGameFolder = False

End Function
Private Function GetDefaultGBGameFolder() As String

    GetDefaultGBGameFolder = Left$(m_strAppPathWithBackslash, 3) & "GBGame"

End Function
Private Sub GetGameIDFromGameFilename()

    Dim lngIdx As Long
    Dim a_strQueries(3) As String
    Dim strSuccessfulQuery As String

    'exact match first
    a_strQueries(0) = "SELECT GA_Id FROM Games WHERE Filename = """ & WCR(m_udtCLP.strGameFilename) & """"

    'filename without path next
    a_strQueries(1) = "SELECT GA_Id FROM Games WHERE Filename LIKE ""*\" & WCR(m_udtCLP.strGameFilename) & """"

    'filename with path, but no filename extension next
    a_strQueries(2) = "SELECT GA_Id FROM Games WHERE Filename LIKE """ & WCR(m_udtCLP.strGameFilename) & ".*"""

    'filename without path and without filename extension last
    a_strQueries(3) = "SELECT GA_Id FROM Games WHERE Filename LIKE ""*\" & WCR(m_udtCLP.strGameFilename) & ".*"""

    'try and find the game ID from the above queries
    strSuccessfulQuery = ""
    For lngIdx = 0 To UBound(a_strQueries)
        Set m_rsGames = m_dbGamebase.OpenRecordset(a_strQueries(lngIdx), dbOpenSnapshot)
        If m_rsGames.EOF Then
            m_rsGames.Close
            Set m_rsGames = Nothing
            Call LogStatus("SQL [" & a_strQueries(lngIdx) & "] returned no rows", e_do_allow_continue)
        Else
            strSuccessfulQuery = a_strQueries(lngIdx)
            Exit For
        End If
    Next lngIdx

    'failed?
    If strSuccessfulQuery = "" Then
        Call LogStatus( _
                "Game ID unable to be determined from Game Filename [" & m_udtCLP.strGameFilename & "] or sql SELECT error", _
                e_do_force_quit _
            )
    End If

    'success!
    With m_rsGames
        .MoveFirst
        m_udtCLP.strGameID = !GA_Id
        .Close
    End With
    Set m_rsGames = Nothing

    Call LogStatus("Game ID determined [" & m_udtCLP.strGameID & "] from Game Filename [" & m_udtCLP.strGameFilename & "]", e_do_allow_continue)
    Call LogStatus("SQL used to determine Game ID [" & strSuccessfulQuery & "]", e_do_allow_continue)

End Sub

Private Sub ShowUsage()

    Dim NL As String
    Dim BL As String
    Dim T As String
    Dim s As String

    T = vbTab
    NL = vbCrLf
    BL = NL & NL

    s = "GBLauncher.exe db=<db name> [id=<game id>|fn=<game filename>] [emu=<emu index>] [wait=1]" & BL _
    & "e.g.  GBLauncher.exe db=""GameBase Amiga"" fn=Cannon_Fodder.adf emu=2 wait=1" & BL _
    & "db=<db name>" & T & T & "Required. Name of GameBase database containing game to launch." & BL _
    & "id=<game id>" & T & T & "Required if 'fn' argument is not specified." & NL _
    & T & T & T & "This is the GameBase Game ID of the game to launch." & BL _
    & "fn=<game filename>" & T & "Required if 'id' argument is not specified." & NL _
    & T & T & T & "This is the GameBase Game Filename of the game to launch." & NL _
    & T & T & T & "Can be a path-less filename, or can include the relative path" & NL _
    & T & T & T & "stored for the game in the GameBase database (with/without filename extension)." & BL _
    & "emu=<emu index>" & T & T & "Optional. If specified, the corresponding emulator from" & NL _
    & T & T & T & "the database's Emulators.ini file will be used. Defaults to the default" & NL _
    & T & T & T & "emulator used when someone clicks the 'Play Game' button in GameBase." & BL _
    & "wait=1" & T & T & T & "Optional. Causes GBLauncher to wait until the emulator is closed" & NL _
    & T & T & T & "before quitting."

    MsgBox s, vbInformation, mc_strAPPNAME & " v" & m_strAPPVERSION & "  (c)2010 " & mc_strAUTHOR

End Sub

Private Function WCR(strSearchString_in As String) As String

    'WildCard Replacer

    WCR = Replace(strSearchString_in, "[", "[[]")
    WCR = Replace(WCR, "?", "[?]")
    WCR = Replace(WCR, "*", "[*]")
    WCR = Replace(WCR, "!", "[!]")
    WCR = Replace(WCR, "#", "[#]")

End Function
Private Sub GetPathInfo(strPath_in As String, udtPath_inout As PathsType)

    '-------------------------------------------------
    'Fills the PathsType structure passed in.
    '
    'Inputs:    strPath_in:     The path to test
    '
    'Modify:    udtPath_inout:  The structure to fill.
    '-------------------------------------------------

    Dim strValidateLocal As String
    Dim iLen As Integer
    
    'has to have at least 3 letters ("c:\")
    With udtPath_inout
        .strRealPath = strPath_in
        iLen = Len(.strRealPath)
       
        If iLen > 2 Then
            'convert root "\" to current drive root
            If Left$(.strRealPath, 1) = "\" Then
                .strRealPath = Left$(m_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 1))
                iLen = Len(.strRealPath)
            End If
            
            'convert relative path to absolute
            If Mid$(.strRealPath, 2, 1) <> ":" Then
                .strRealPath = m_strAppPathWithBackslash & .strRealPath
                iLen = Len(.strRealPath)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If m_xobjFunc.DirExists(.strRealPath) = True Then
                    .strRealPath = CurDir()
                End If
            End If
                    
            'if specified path does not exist, try to validate on the current drive
            If m_xobjFunc.DirExists(.strRealPath) = False Then
                strValidateLocal = Left$(m_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 3))
                If m_xobjFunc.DirExists(strValidateLocal) = True Then
                    .strRealPath = strValidateLocal
                End If
            End If
        
            'set path and path with backslash
            .strWithBackSlash = .strRealPath
            If Right$(.strWithBackSlash, 1) <> "\" Then .strWithBackSlash = .strWithBackSlash & "\"
            
            'test drive type
            .ePathType = GetTypeOfDrive(Left$(.strRealPath, 3))
            If .ePathType = e_PathLocation_FixedHD Then
                'fixed drive - make sure path exists
                If m_xobjFunc.DirExists(.strRealPath) = False Then
                    .ePathType = e_PathLocation_Invalid
                End If
            End If
        Else
            .ePathType = e_PathLocation_Invalid
        End If
    End With
    
End Sub
Private Function GetTypeOfDrive(strRootPath_in As String) As PathLocationEnum

    '---------------------------------------------------
    'gets the type of drive given the root path supplied
    '
    'path supplied in "c:\" format
    '---------------------------------------------------

    Dim lngDriveType As Long        'result from API

    lngDriveType = GetDriveType(strRootPath_in)

    Select Case lngDriveType
    Case DRIVE_FIXED, DRIVE_REMOTE, DRIVE_RAMDISK
        GetTypeOfDrive = e_PathLocation_FixedHD
    Case DRIVE_CDROM
        GetTypeOfDrive = e_PathLocation_CD
    Case DRIVE_REMOVABLE
        GetTypeOfDrive = e_PathLocation_Removable
    Case Else
        GetTypeOfDrive = e_PathLocation_Invalid
    End Select

End Function
Private Sub CreateLogFile()

    On Error GoTo Err_Handler

    'delete existing log file (always start with empty log)
    If m_xobjFunc.FileExists(m_strAppPathWithBackslash & "GBLauncher.log", False, mc_strAPPNAME) = True Then
        Call m_xobjFunc.DeletePathSilent(m_strAppPathWithBackslash & mc_strLOGFILENAME)
    End If

    m_intLogFileNum = FreeFile()
    Open m_strAppPathWithBackslash & mc_strLOGFILENAME For Output As #m_intLogFileNum
    Print #m_intLogFileNum, "***** " & mc_strAPPNAME & " v" & m_strAPPVERSION & " *****"
    Print #m_intLogFileNum, "Started: " & Format(Now(), "yyyy-MM-dd HH:mm:ss")

    Exit Sub

Err_Handler:
    MsgBox "Could not create Log File. " & Err.Description & "  [" & Err.Number & ", " & Err.LastDllError & "]", vbCritical, mc_strAPPNAME
    Set m_xobjFunc = Nothing
    End

End Sub
Private Sub Form_Load()

    On Error GoTo Err_Handler

    Dim strDBPathFile As String
    Dim lngTemp As Long
    Dim lngPathIdx As Long
    Dim strPaths() As String
    Dim strPath As String
    Dim strTemp As String
    Dim strGBGamePath As String

    m_blnIsDBOpen = False
    Set m_xobjFunc = New GBFuncs.clsFunctions

    'app path and version
    m_strAppPathWithBackslash = App.Path
    If Right$(m_strAppPathWithBackslash, 1) <> "\" Then m_strAppPathWithBackslash = m_strAppPathWithBackslash & "\"

    If m_blnInVBIDE = True Then
        m_strAppPathWithBackslash = mc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If

    m_strAPPVERSION = Format$(Val(App.Major), "0") & "." & Format$(Val(App.Minor), "0")

    'create log file
    Call CreateLogFile
    Call LogStatus("App Path: [" & m_strAppPathWithBackslash & "]", e_do_allow_continue)

    'parse the command line
    Call ParseCommandLine

    'set all path globals
    strDBPathFile = m_strAppPathWithBackslash & m_udtCLP.strDBName & "\" & m_udtCLP.strDBName & ".mdb"
    If m_xobjFunc.FileExists(strDBPathFile, False, mc_strAPPNAME) = True Then
        m_strGameBaseName = m_udtCLP.strDBName
        m_strGameBaseDBPath = strDBPathFile
        m_strPathFile = m_strAppPathWithBackslash & m_udtCLP.strDBName & "\Paths.ini"
        m_strSettingsFile = m_strAppPathWithBackslash & m_udtCLP.strDBName & "\Settings.ini"
        m_strGameBaseDotINI = m_strAppPathWithBackslash & "GameBase.ini"
        Call LogStatus("Database name: [" & m_strGameBaseName & "]", e_do_allow_continue)
        Call LogStatus("Database file: [" & m_strGameBaseDBPath & "]", e_do_allow_continue)
        Call LogStatus("Paths file: [" & m_strPathFile & "]", e_do_allow_continue)
        Call LogStatus("Settings file: [" & m_strSettingsFile & "]", e_do_allow_continue)
    Else
        Call LogStatus("Database: [" & strDBPathFile & "] not found", e_do_force_quit)
    End If

    'get rezipping settings
    m_blnReZip = (m_xobjFunc.INI_GetSetting("General", "ReZipGames", "0", m_strSettingsFile) = "1")
    m_blnRezipAsk = (m_xobjFunc.INI_GetSetting("General", "ReZipAsk", "1", m_strSettingsFile) = "1")
    m_blnZipNotify = (m_xobjFunc.INI_GetSetting("General", "ReZipNotify", "1", m_strSettingsFile) = "1")
    m_blnRezipLegacy = (m_xobjFunc.INI_GetSetting("General", "ReZipLegacy", "0", m_strSettingsFile) = "1")
    lngTemp = Val(m_xobjFunc.INI_GetSetting("General", "ReZipAdditional", "2", m_strSettingsFile))
    m_blnPromptForGamePathDevice = (m_xobjFunc.INI_GetSetting("General", "RequestGameDeviceInsert", "1", m_strSettingsFile) = "1")
    If lngTemp = 0 Then
        me_RezipAdditional = eRezipAdditionalIgnore
        strTemp = "Ignore"
    ElseIf lngTemp = 1 Then
        me_RezipAdditional = eRezipAdditionalZip
        strTemp = "Zip"
    Else
        me_RezipAdditional = eRezipAdditionalAsk
        strTemp = "Ask"
    End If
    Call LogStatus("ReZipGames: [" & m_blnReZip & "]", e_do_allow_continue)
    Call LogStatus("ReZipAsk: [" & m_blnRezipAsk & "]", e_do_allow_continue)
    Call LogStatus("ReZipNotify: [" & m_blnZipNotify & "]", e_do_allow_continue)
    Call LogStatus("ReZipLegacy: [" & m_blnRezipLegacy & "]", e_do_allow_continue)
    Call LogStatus("ReZipAdditional: [" & strTemp & "]", e_do_allow_continue)
    Call LogStatus("RequestGameDeviceInsert: [" & m_blnPromptForGamePathDevice & "]", e_do_allow_continue)

    'open the database
    Set m_dbGamebase = OpenDatabase(m_strGameBaseDBPath)
    m_blnIsDBOpen = True
    Call LogStatus("Opened Database", e_do_allow_continue)
    
    'grab the database version info
    Set m_rsGames = m_dbGamebase.OpenRecordset("Config", dbOpenSnapshot)
    With m_rsGames
        If Not .EOF Then
            .MoveFirst
            m_lngDBMajorVersion = !MajorVersion
            m_lngDBMinorVersion = !MinorVersion
            m_lngOfficialDBUpdate = !OfficialUpdate
            .Close
        Else
            .Close
            m_dbGamebase.Close
            Call LogStatus("Missing Version/Update info in Config Table!", e_do_force_quit)
        End If
    End With
    Set m_rsGames = Nothing
    m_sngDBVersion = m_lngDBMajorVersion + (CSng(m_lngDBMinorVersion / 10))
    If m_sngDBVersion < mc_sngREQ_DB_VERSION Then
        Call LogStatus("Minimum required GameBase database version: [" & mc_sngREQ_DB_VERSION & "], you have: [" & m_sngDBVersion & "]", e_do_force_quit)
    End If
    Call LogStatus("Database version validated [" & m_sngDBVersion & "]", e_do_allow_continue)

    'GBGame path
    strGBGamePath = m_xobjFunc.INI_GetSetting("Paths", "GBGame", "", m_strGameBaseDotINI)
    If strGBGamePath = "" Then
        strGBGamePath = GetDefaultGBGameFolder()
    End If
    On Error Resume Next
    Call MkDir(strGBGamePath)
    On Error GoTo Err_Handler
    Call GetPathInfo(strGBGamePath, m_udtPathGBGame)
    If m_udtPathGBGame.ePathType = e_PathLocation_Invalid Then
        'not valid, try to use default path
        If Not m_xobjFunc.DirExists(GetDefaultGBGameFolder()) Then
            If Not CreateDefaultGBGameFolder() Then
                'could not create default folder
                Call LogStatus("Invalid temp GBGame path in GameBase.ini [" & strGBGamePath & "]", e_do_force_quit)
            End If
        End If
        'using default as specified does not exist
        Call GetPathInfo(GetDefaultGBGameFolder(), m_udtPathGBGame)
        Call LogStatus("Invalid temp GBGame path in GameBase.ini. Using [" & GetDefaultGBGameFolder() & "]", e_do_allow_continue)
    Else
        Call LogStatus("GBGame path found: [" & m_udtPathGBGame.strRealPath & "]", e_do_allow_continue)
    End If

    'game paths
    lngPathIdx = 1
    m_lngTotalGamePaths = 0
    ReDim strPaths(0)
    Do
        strPath = m_xobjFunc.INI_GetSetting("Games", Trim$(lngPathIdx), "", m_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        m_lngTotalGamePaths = lngPathIdx - 1
        ReDim ma_udtRegPathGame(m_lngTotalGamePaths - 1)
        For lngPathIdx = 0 To (m_lngTotalGamePaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ma_udtRegPathGame(lngPathIdx))
        Next lngPathIdx
    End If
    m_blnValidGamePathExists = False
    For lngPathIdx = 0 To (m_lngTotalGamePaths - 1)
        If ma_udtRegPathGame(lngPathIdx).ePathType <> e_PathLocation_Invalid Then
            m_blnValidGamePathExists = True
            Call LogStatus("Game path found: [" & ma_udtRegPathGame(lngPathIdx).strRealPath & "]", e_do_allow_continue)
        End If
    Next lngPathIdx
    If m_blnValidGamePathExists = False Then
        Call LogStatus("No valid game paths exist as specified in [" & m_strPathFile & "]", e_do_force_quit)
    End If

    If m_udtCLP.strGameID = "" Then
        Call GetGameIDFromGameFilename
    End If

    'load the game info
    Call LoadGameInfo(m_udtCLP.strGameID)

    'run the game
    Call RunGame(m_udtCLP.strEmuIdx)

    'all done
    Call ShutDown

Err_Handler:
    Call LogStatus("Error in Form_Load(): " & Err.Description & "  [" & Err.Number & ", " & Err.LastDllError & "]", e_do_force_quit)

End Sub
Private Sub ParseCommandLine()

    Dim strCommand As String
    Dim objMatches As MatchCollection
    Dim objMatch As Match
    Dim intSubMatchIdx As Integer
    Dim objRegex As RegExp
    Dim strWait As String
    Dim lngIDMatches As Long
    Dim lngFilenameMatches As Long

    m_udtCLP.strGameID = ""
    m_udtCLP.strGameFilename = ""

    'show usage and exit if no command line params
    If Command$ = "" Then
        Call ShowUsage
        Call LogStatus("Empty command line - usage shown in MsgBox", e_do_force_quit)
    End If
    strCommand = Command$
    Call LogStatus("Command line: [" & strCommand & "]", e_do_allow_continue)

    'init the regex object
    Set objRegex = New RegExp
    objRegex.IgnoreCase = True
    objRegex.Global = False

    'get the database
    objRegex.Pattern = "(db|database)=(""([^""]+)""|([^"" ]+))"
    Set objMatches = objRegex.Execute(strCommand)
    If objMatches.Count = 0 Then
        Call LogStatus("Missing or invalid database in command line.", e_do_force_quit)
    End If
    m_udtCLP.strDBName = objMatches.Item(0).SubMatches(1)
    If Left$(m_udtCLP.strDBName, 1) = """" And Right$(m_udtCLP.strDBName, 1) = """" Then
        m_udtCLP.strDBName = Mid$(m_udtCLP.strDBName, 2, Len(m_udtCLP.strDBName) - 2)
    End If
    Call LogStatus("db param: [" & m_udtCLP.strDBName & "]", e_do_allow_continue)

    'get the game id
    objRegex.Pattern = "(id|game|gameid)=(""([0-9]+)""|([0-9]+))"
    Set objMatches = objRegex.Execute(strCommand)
    lngIDMatches = objMatches.Count
    If lngIDMatches > 0 Then
        m_udtCLP.strGameID = objMatches.Item(0).SubMatches(1)
        If Left$(m_udtCLP.strGameID, 1) = """" And Right$(m_udtCLP.strGameID, 1) = """" Then
            m_udtCLP.strGameID = Mid$(m_udtCLP.strGameID, 2, Len(m_udtCLP.strGameID) - 2)
        End If
        Call LogStatus("id param: [" & m_udtCLP.strGameID & "]", e_do_allow_continue)
    End If

    'get the game filename
    objRegex.Pattern = "(fn|filename)=(""([^""]+)""|([^"" ]+))"
    Set objMatches = objRegex.Execute(strCommand)
    lngFilenameMatches = objMatches.Count
    If lngFilenameMatches > 0 Then
        m_udtCLP.strGameFilename = objMatches.Item(0).SubMatches(1)
        If Left$(m_udtCLP.strGameFilename, 1) = """" And Right$(m_udtCLP.strGameFilename, 1) = """" Then
            m_udtCLP.strGameFilename = Mid$(m_udtCLP.strGameFilename, 2, Len(m_udtCLP.strGameFilename) - 2)
        End If
        Call LogStatus("filename param: [" & m_udtCLP.strGameFilename & "]", e_do_allow_continue)
    End If
    
    'check id/filename
    If lngIDMatches = 0 And lngFilenameMatches = 0 Then
        Call LogStatus("Missing or invalid id/filename param in command line.", e_do_force_quit)
    ElseIf lngIDMatches > 0 And lngFilenameMatches > 0 Then
        Call LogStatus("Please specify only one of id/filename params in command line.", e_do_force_quit)
    End If

    'get the emulator index
    objRegex.Pattern = "(emu|emulator)=(""([0-9]+)""|([0-9]+))"
    Set objMatches = objRegex.Execute(strCommand)
    If objMatches.Count = 0 Then
        m_udtCLP.strEmuIdx = -1 'default emulator
        Call LogStatus("emu param missing: using -1 (default)", e_do_allow_continue)
    Else
        m_udtCLP.strEmuIdx = objMatches.Item(0).SubMatches(1)
        If Left$(m_udtCLP.strEmuIdx, 1) = """" And Right$(m_udtCLP.strEmuIdx, 1) = """" Then
            m_udtCLP.strEmuIdx = Mid$(m_udtCLP.strEmuIdx, 2, Len(m_udtCLP.strEmuIdx) - 2)
        End If
        Call LogStatus("emu param: [" & m_udtCLP.strEmuIdx & "]", e_do_allow_continue)
    End If

    'get the wait flag
    objRegex.Pattern = "(wait)=(""([01])""|([01]))"
    Set objMatches = objRegex.Execute(strCommand)
    If objMatches.Count = 0 Then
        m_udtCLP.blnWait = False
        Call LogStatus("wait param missing: using 0 (default)", e_do_allow_continue)
    Else
        strWait = objMatches.Item(0).SubMatches(1)
        If strWait = """1""" Or strWait = """0""" Then
        
            strWait = Mid$(strWait, 2, Len(strWait) - 2)
        End If
        m_udtCLP.blnWait = IIf((strWait = "1"), True, False)
        Call LogStatus("wait param: [" & strWait & "]", e_do_allow_continue)
    End If

    Set objRegex = Nothing

End Sub
Private Sub LogStatus(strError As String, eDoQuit As LogStatusEnum)

    'MsgBox strError, vbCritical
    Print #m_intLogFileNum, "--> " & strError
    If eDoQuit = e_do_force_quit Then Call ShutDown

End Sub
Private Sub LoadGameInfo(strGameID As String)

    Dim strQuery As String
    Dim lngID As Long
    
    lngID = Val(strGameID)

    'build the query
    strQuery = "SELECT GA_Id, Comment, Control, Filename, FileToRun, FilenameIndex, FA, V_Comment, " _
        & "PlayersTo, PlayersFrom, V_PalNTSC, Gemus, V_TrueDriveEmu " _
        & "FROM Games WHERE GA_Id = " & lngID

    'Run the query, get the Game info from the database and store globally
    Set m_rsGames = m_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
    If m_rsGames.EOF Then
        Set m_rsGames = Nothing
        Call LogStatus("Game ID [" & lngID & "] not found or sql SELECT error [" & strQuery & "]", e_do_force_quit)
    End If
    With m_rsGames
        .MoveFirst
        m_udtGame.lngID = lngID
        m_udtGame.strPairs = !Gemus
        m_udtGame.strGameFilename = !FileName
        m_udtGame.blnGameFileExists = ((!FA = True) And (m_blnValidGamePathExists = True))
        m_udtGame.strFileToRun = !FileToRun
        m_udtGame.lngGameFilenameIndex = !FilenameIndex
        m_udtGame.strComment = !Comment
        m_udtGame.lngPlayersFrom = !PlayersFrom
        m_udtGame.lngPlayersTo = !PlayersTo
        m_udtGame.lngNTSC = !V_PalNTSC
        m_udtGame.lngControl = !Control
        m_udtGame.blnTrueDriveEmulation = !V_TrueDriveEmu
        m_udtGame.strVersionComment = !V_Comment
        .Close
    End With
    Set m_rsGames = Nothing
    Call LogStatus("SQL used to read game information [" & strQuery & "]", e_do_allow_continue)
    Call LogStatus("Game information loaded for game ID [" & lngID & "]", e_do_allow_continue)

End Sub
Private Sub RunGame(ByVal lngEmuIdx_in As Long)

    Dim strError As String
    Dim lngGPIdx As Long
    Dim strFileToRun As String
    Dim strDefaultFile As String
    Dim objGEMUS As Gemus.clsGemus

    'find actual game path\file
    If m_udtGame.blnGameFileExists = True Then
        strError = FindGameFile(lngGPIdx, m_udtGame.strGameFilename)
        If strError <> "" Then
            'not found or aborted
            If strError <> "~~ABORTED~~" Then
                Call LogStatus("Gamefile [" & m_udtGame.strGameFilename & "] not located: " & strError, e_do_force_quit)
            Else
                Call LogStatus("Launching game  [" & m_udtGame.strGameFilename & "] was aborted by user", e_do_force_quit)
            End If
        Else
            'game found!
            strDefaultFile = m_udtGame.strFileToRun
            strFileToRun = ma_udtRegPathGame(lngGPIdx).strWithBackSlash & m_udtGame.strGameFilename
            Call LogStatus("Attempting to launch found game file [" & strFileToRun & "] in emulator [" & lngEmuIdx_in & "]", e_do_allow_continue)
            With m_udtGame
                'extract and run game
                Set objGEMUS = New Gemus.clsGemus
                Call objGEMUS.SetSilentMode(True)
                Call objGEMUS.Init(m_strGameBaseName)
                Call objGEMUS.SetRezipNotify(m_blnZipNotify)
                Call objGEMUS.SetRezip(m_blnReZip, m_blnRezipAsk, me_RezipAdditional)
                Call objGEMUS.SetReZipLegacy(m_blnRezipLegacy)
                If m_udtCLP.blnWait = True Then
                    Call objGEMUS.SetWaitIfNotRezipping(True)
                End If
                strError = objGEMUS.RunEmulator(lngEmuIdx_in, strFileToRun _
                , .lngGameFilenameIndex, strDefaultFile, m_udtPathGBGame.strRealPath, True, .lngPlayersFrom _
                , .lngPlayersTo, .lngControl, .strComment, .strVersionComment, .lngNTSC, .blnTrueDriveEmulation, .strPairs)
                Set objGEMUS = Nothing
            End With

            If strError <> "" Then
                Call LogStatus("GEMUS failed running game [" & strFileToRun & "]: " & strError, e_do_force_quit)
            Else
                Call LogStatus("Game Launched by GEMUS - no errors reported", e_do_allow_continue)
            End If
        End If
    Else
        Call LogStatus("Gamefile [" & m_udtGame.strGameFilename & "] not found in any of your game paths: " & strError, e_do_force_quit)
    End If

End Sub
Private Sub ShutDown()

    If m_blnIsDBOpen = True Then m_dbGamebase.Close
    Set m_xobjFunc = Nothing
    Print #m_intLogFileNum, "Bye!"
    Close m_intLogFileNum
    End

End Sub

Private Function FindGameFile(lngGPIdx_out As Long, strFilename_in As String) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult

    If m_lngTotalGamePaths = 0 Then
        FindGameFile = m_xobjFunc.GetINIMessage("Error Messages", "Error_NoGamePaths", "You have not set up any Game Paths!")
    ElseIf m_blnValidGamePathExists = False Then
        FindGameFile = m_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidGamePath", "There are no valid Game Paths!")
    Else
        'test each path
        strNotFound = m_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!")
        For lngPathIdx = 0 To (m_lngTotalGamePaths - 1)
            With ma_udtRegPathGame(lngPathIdx)
                If m_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                    'found file
                    lngGPIdx_out = lngPathIdx
                    FindGameFile = ""
                    Exit Function
                Else
                    'file not found in current game path
                    If (.ePathType = e_PathLocation_CD) Or (.ePathType = e_PathLocation_Removable) Then
                        If m_blnPromptForGamePathDevice = True Then
                            'path is CD or removable drive... let user insert media
                            Do
                                mbRes = MsgBox(strNotFound & vbCrLf & vbCrLf & "'" & .strWithBackSlash & strFilename_in & "'", vbExclamation Or vbAbortRetryIgnore, mc_strAPPNAME)
                                If mbRes = vbAbort Then
                                    FindGameFile = "~~ABORTED~~"
                                    Exit Function
                                ElseIf mbRes = vbIgnore Then
                                    Exit Do
                                Else
                                    If m_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                                        'found file
                                        lngGPIdx_out = lngPathIdx
                                        FindGameFile = ""
                                        Exit Function
                                    End If
                                End If
                            Loop
                        End If
                    End If
                End If
            End With
        Next lngPathIdx
        'if here then the game wasn't found
        FindGameFile = m_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInGamePaths", "The file was not found in any of your Game Paths!")
    End If

End Function
