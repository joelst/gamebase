VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGemus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'external objects
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_xobjRegistry As GBFuncs.clsRegistry

Private m_blnSilentMode As Boolean

Private m_blnSimulateRun As Boolean
Private m_strSimulateKV As String

Private m_strPathFile As String 'Paths.ini
Private m_udtRegPathSid As PathsType
Private ma_udtRegPathGame() As PathsType
Private ma_udtRegPathScrnShot() As PathsType
Private ma_udtRegPathExtras() As PathsType
Private m_lngTotalGamePaths As Long
Private m_lngTotalScrnShotPaths As Long
Private m_lngTotalExtrasPaths As Long
Private m_udtRegPathPhotos As PathsType
Private m_blnValidGamePathExists As Boolean


Private Const mc_strNumbersAllowed As String = "0123456789"
Private m_blnUseShort As Boolean
Private m_lngEmuIdx As Long
Private Const gc_lngPLAYERSUNKNOWN As Long = -1
Private m_blnScriptFileErrorDone As Boolean

Private Type SendKeysBlockType
    lngSecsToWait As Long
    strKeysToSend As String
End Type

'general member variables
Private m_strAppPathWithBackSlash As String 'app path
Private m_strLangINIPathFile As String      'language ini path\file
Private m_strAppName As String              'client application's name (e.g. "GameBase")
Private m_strCommandCom As String           'for shelling dos progs
Private m_strEmuFile As String
Private m_strGBName As String
Private m_strGBGamePath As String
Private m_blnWaitIfNotRezipping As Boolean

'repack stuff
Private m_blnUserWantsToRepack As Boolean
Private m_blnAskUserBeforeRepack As Boolean
Private m_blnNotifyUserIfRepackWasOK As Boolean
Private m_blnAttemptRepack As Boolean
Public Enum EnumRezipAdditionalFiles
    eRezipAdditionalIgnore = 0
    eRezipAdditionalZip = 1
    eRezipAdditionalAsk = 2
End Enum
Private m_blnRezipLegacy As Boolean
Private me_RepackAddtional As EnumRezipAdditionalFiles
Private Type UnzippedFileType
    strName As String
    strUCaseName As String
    strCRC As String
    blnChecked As Boolean
End Type
Private WithEvents m_objZipper As cZip
Attribute m_objZipper.VB_VarHelpID = -1

'emulators
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

Private Type EmulatorType
    strName As String
    strFullPathAndFilename As String
    strFullPathOnly As String
    strFullPathOnlyWithBackSlash As String
    strFullFilenameOnly As String
    strShortPathAndFilename As String
    strShortPathOnly As String
    strShortPathOnlyWithBackSlash As String
    strShortFilenameOnly As String
    strScript As String
    blnUseShortFilenames As Boolean
    ePathType As PathLocationEnum
    strFiles() As String
    lngNumFiles As Long
    strTypes As String
    blnUseDefaultTypes As Boolean
    lngWait As Long
End Type
Private ma_udtEmulators() As EmulatorType
Private m_lngTotalEmulators As Long

'file types
Private ma_strEmulatorFileTypes() As String
Private m_lngTotalEmulatorFileTypes As Long
Private Const mc_strDEFAULT_EMU_FILETYPES As String = ""

'manual script fields
Public Type ManualGEMUSFields
    lngControl As Long
    lngPlayers As Long
    lngPalNTSC As Long
    blnTDE As Boolean
    strComment As String
    strVComment As String
    strNotes As String
End Type

'script
Public Enum ScriptValueTypeEnum
    e_ScriptValType_Number
    e_ScriptValType_List
    e_ScriptValType_Text
End Enum
Public Type ScriptFieldType
    strName As String
    strOperators() As String
    lngTotalOperators As Long
    eValueType As ScriptValueTypeEnum
    strListValues() As String
    lngTotalListValues As Long
End Type
Public Enum ScriptFunctionParamTypeEnum
    e_ScriptFuncParamType_FileNumber
    e_ScriptFuncParamType_INISection
    e_ScriptFuncParamType_Text
    e_ScriptFuncParamType_List
End Enum
Public Type ScriptFunctionType
    strName As String
    lngTotalParams As Long
    eParamType() As ScriptFunctionParamTypeEnum
    lngTotalListValues As Long
    strListValues() As String
End Type
Private ma_udtScriptFields() As ScriptFieldType
Private m_lngMaxScriptFields As Long
Private ma_udtScriptFunctions() As ScriptFunctionType
Private m_lngMaxScriptFunctions As Long
Private ma_strScriptEnvVars() As String
Private m_lngMaxScriptEnvVars As Long
Private m_lngTotalScriptLines As String
Private ma_strScriptLines() As String
Private ma_strScriptLinesPre() As String
Private m_lngTotalScriptLinesToProcess As Long
Private ma_strScriptLinesToProcess() As String

'language-text from ini file
Private m_strLang_EndOfLineExpected As String
Private m_strLang_ScriptError As String
Private m_strLang_Line As String
Private m_strLang_UnknownCommand As String
Private m_strLang_MissingParams As String
Private m_strLang_TooManyParams As String
Private m_strLang_Expected As String
Private m_strLang_GameNotSupported As String
Private m_strLang_InvalidFunctionParam As String
Private m_strLang_Number As String
Private m_strLang_ExpressionReqd As String
Private m_strLang_UnknownField As String
Private m_strLang_InvalidOp As String
Private m_strLang_InvalidValue As String
Private m_strLang_Without As String
Private m_strLang_TooMany As String
Private m_strLang_Statements As String

'database fields
Private m_lngFldNumPlayersFrom As Long
Private m_lngFldNumPlayersTo As Long
Private m_lngFldControl As Long
Private m_strFldGameComment As String
Private m_strFldVersionComment As String
Private m_lngFldPALNTSC As Long
Private m_blnFldTrueDriveEmu As Boolean

'key=value fields
Public Type NotesKeysType
    strKey As String
    strValue As String
End Type
Private ma_strNotesKeys() As NotesKeysType
Private m_lngTotalNotesKeys As Long
Public Type NotesKeysLinesType
    strKey As String
    strValue As String
    blnIsKV As Boolean
    strLine As String
    blnFlagged As Boolean
End Type

'environment fields
Private m_strFldGameType As String
Private m_lngFldNumGameFiles As Long

'environment variables
Private m_lngImageIndex As Long
Private m_strImageName As String
Private m_strC64ImageName As String
Private m_strEnvGFFull As String
Private m_strEnvGFFile As String
Private m_strEnvGFPath As String
Private m_strEnvGFNoExt As String
Private ma_strExtraGFFull(1 To 20) As String
Private ma_strExtraGFPath(1 To 20) As String
Private ma_strExtraGFFile(1 To 20) As String
Private m_strExtraGFNoExt(1 To 20) As String

'emulator command line
Private m_strCommandLine As String

'API
Private Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const SW_MINIMIZE = 6
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_SILENT = &H4
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const INFINITE = &HFFFF
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_FIXED = 3
Private Const DRIVE_CDROM = 5

'for shelling files with their default program
Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hWnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional parameter
    lpClass       As String   'Optional parameter
    hkeyClass     As Long     'Optional parameter
    dwHotKey      As Long     'Optional parameter
    hIcon         As Long     'Optional parameter
    hProcess      As Long     'Optional parameter
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const SW_SHOW = 5
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Private Const PROCESS_TERMINATE = &H1
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const TH32CS_SNAPPROCESS = &H2

Private Sub CreateLogFile()

    On Error GoTo Err_Handler
    
    g_strLogFileTrace = ""
    ReDim ga_strLogFileTrace(0)
    g_intLogFileNum = 0

    'delete existing log file (always start with empty log)
    If m_xobjFunc.FileExists(m_strAppPathWithBackSlash & "Gemus.log", False, "GEMUS") = True Then
        Call m_xobjFunc.DeletePathSilent(m_strAppPathWithBackSlash & "Gemus.log")
    End If

    g_intLogFileNum = FreeFile()
    Open m_strAppPathWithBackSlash & "Gemus.log" For Output As #g_intLogFileNum
    Print #g_intLogFileNum, "***** " & "GEMUS LOG FILE" & " *****"
    Print #g_intLogFileNum, "Started: " & Format(Now(), "yyyy-MM-dd HH:mm:ss")

    Exit Sub

Err_Handler:
    g_intLogFileNum = 0

End Sub
Private Sub CloseLogFile()

    On Error GoTo Err_Handler

    If g_intLogFileNum > 0 Then
        Print #g_intLogFileNum, "Bye!"
        Close #g_intLogFileNum
    End If

Err_Handler:

End Sub
Public Function GetAllValidFileTypesArray(a_strFileTypes_out() As String) As Long

    'read in all the emulator filetypes and return an array that has them all in

    Dim lngIdx As Long
    Dim lngIdx2 As Long
    Dim lngFTIdx As Long
    Dim lngTotal As Long
    Dim blnDoneDefaults As Boolean
    Dim a_strTypes() As String
    Dim blnFound As Boolean
    Dim lngFindIdx As Long
    Dim strUC As String
    Dim strFileTypes As String

    blnDoneDefaults = False
    lngTotal = 0
    ReDim a_strFileTypes_out(0)
    strFileTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", "", m_strEmuFile)
    Call SetEmuFileTypes(strFileTypes)

    For lngIdx = 0 To m_lngTotalEmulators - 1
        With ma_udtEmulators(lngIdx)
            If .blnUseDefaultTypes = True Then
                'emu uses default file types
                If blnDoneDefaults = False Then
                    For lngIdx2 = 0 To m_lngTotalEmulatorFileTypes - 1
                        blnFound = False
                        strUC = UCase(ma_strEmulatorFileTypes(lngIdx2))
                        For lngFindIdx = 0 To lngTotal - 1
                            If UCase(a_strFileTypes_out(lngFindIdx)) = strUC Then
                                blnFound = True
                                Exit For
                            End If
                        Next lngFindIdx
                        If blnFound = False Then
                            ReDim Preserve a_strFileTypes_out(lngTotal)
                            a_strFileTypes_out(lngTotal) = strUC
                            lngTotal = lngTotal + 1
                        End If
                    Next lngIdx2
                    blnDoneDefaults = True
                End If
            Else
                'emu uses its own file types
                If .strTypes <> "" Then
                    a_strTypes = Split(.strTypes, ";")
                    For lngFTIdx = 0 To UBound(a_strTypes)
                        a_strTypes(lngFTIdx) = Trim$(a_strTypes(lngFTIdx))
                        If ValidateEmuFileType(a_strTypes(lngFTIdx)) = True Then
                            blnFound = False
                            strUC = UCase$(a_strTypes(lngFTIdx))
                            For lngFindIdx = 0 To lngTotal - 1
                                If UCase(a_strFileTypes_out(lngFindIdx)) = strUC Then
                                    blnFound = True
                                    Exit For
                                End If
                            Next lngFindIdx
                            If blnFound = False Then
                                ReDim Preserve a_strFileTypes_out(lngTotal)
                                a_strFileTypes_out(lngTotal) = strUC
                                lngTotal = lngTotal + 1
                            End If
                        End If
                    Next lngFTIdx
                End If
            End If
        End With
    Next lngIdx

    GetAllValidFileTypesArray = lngTotal

End Function

Public Function GetKVSet(a_udtKV() As NotesKeysType, ByVal strSetName As String)

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim lngPos As Long
    Dim strKey As String
    Dim strVal As String
    Dim strLine As String
    Dim blnInDesiredSet As Boolean
    
    'default to no key=value pairs found
    ReDim a_udtKV(0)
    GetKVSet = 0
    
    'if no Gemus.ini file, abort
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then Exit Function
    
    'read in the lines
    f = FreeFile()
    Open strFile For Input As #f
    blnInDesiredSet = False
    Do Until EOF(f)
        Line Input #f, strLine
        strLine = Trim$(strLine)
        If Left$(strLine, 1) = "[" And Right$(strLine, 1) = "]" Then
            'new section
            If Len(strLine) > 2 Then
                strLine = Mid$(strLine, 2, Len(strLine) - 2)
                If LCase$(strLine) = LCase$(strSetName) Then
                    blnInDesiredSet = True
                ElseIf blnInDesiredSet = True Then
                    Exit Do
                End If
            End If
        ElseIf blnInDesiredSet = True Then
            'in correct section: grab key=value pair and add to array
            lngPos = InStr(strLine, "=")
            If lngPos > 1 And Left$(strLine, 1) <> ";" Then
                strKey = Trim$(Left$(strLine, lngPos - 1))
                If strKey <> "" Then
                    strVal = Right$(strLine, Len(strLine) - lngPos)
                    strVal = Trim$(strVal)
                    ReDim Preserve a_udtKV(GetKVSet)
                    a_udtKV(GetKVSet).strKey = strKey
                    a_udtKV(GetKVSet).strValue = strVal
                    GetKVSet = GetKVSet + 1
                End If
            End If
        End If
    Loop
    Close #f

    Exit Function

Err_Handler:
    ReDim a_udtKV(0)
    GetKVSet = 0

End Function

Public Function GetKVSetAsText(ByVal strSetName_in As String, strText_out As String, strError_out As String) As Boolean

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim strLine As String
    Dim strTrimmedLine As String
    Dim blnInTHESet As Boolean
    Dim blnDoneTHESet As Boolean

    strText_out = ""
    strError_out = ""

    'if no Gemus.ini file then use the incoming kvset as the whole contents
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = True Then
        'read in the lines
        blnInTHESet = False 'we're not in THE set until we encounter it
        blnDoneTHESet = False
        f = FreeFile()
        Open strFile For Input As #f
        Do Until EOF(f)
            Line Input #f, strLine
            strTrimmedLine = Trim$(strLine)
            If LCase$(strTrimmedLine) = "[" & LCase$(strSetName_in) & "]" Then
                'We're in THE set!
                blnInTHESet = True
                blnDoneTHESet = True
            Else
                'Are we still in the set?
                If Left$(strTrimmedLine, 1) = "[" And Right$(strTrimmedLine, 1) = "]" Then
                    If blnDoneTHESet Then
                        Exit Do
                    End If
                End If
                If blnInTHESet Then
                    strText_out = strText_out & strLine & vbCrLf
                End If
            End If
        Loop
        Close #f
    End If

    strText_out = m_xobjFunc.RTrimWhitespace(strText_out)

    GetKVSetAsText = True
    Exit Function

Err_Handler:
    strError_out = Err.Description & " (" & Err.Number & ")"
    GetKVSetAsText = False

End Function
Public Function GetKVSetNames(a_strKVSets() As String)

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim lngPos As Long
    Dim strKey As String
    Dim strVal As String
    Dim strLine As String
    
    'default to no kvsets found
    ReDim a_strKVSets(0)
    GetKVSetNames = 0
    
    'if no Gemus.ini file, abort
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then Exit Function
    
    'read in the lines
    f = FreeFile()
    Open strFile For Input As #f
    Do Until EOF(f)
        Line Input #f, strLine
        strLine = Trim$(strLine)
        If Left$(strLine, 1) = "[" And Right$(strLine, 1) = "]" Then
            'new section
            If Len(strLine) > 2 Then
                strLine = Mid$(strLine, 2, Len(strLine) - 2)
                If LCase$(strLine) <> "global" Then
                    ReDim Preserve a_strKVSets(GetKVSetNames)
                    a_strKVSets(GetKVSetNames) = strLine
                    GetKVSetNames = GetKVSetNames + 1
                End If
            End If
        End If
    Loop
    Close #f

    Exit Function

Err_Handler:
    ReDim a_strKVSets(0)
    GetKVSetNames = 0

End Function
Public Function GetKVsFromString(ByVal strKVs_in As String, a_udtKVs_out() As NotesKeysType)

    'converts a string containing lines of key=value pairs, into an array of key=value pairs

    Dim lngIdx As Long
    Dim lngPos As Long
    Dim a_strLines() As String
    Dim strKey As String
    Dim strVal As String

    GetKVsFromString = 0
    ReDim a_udtKVs_out(0)

    If Trim$(strKVs_in) = "" Then Exit Function

    a_strLines = Split(strKVs_in, vbCrLf)
    For lngIdx = 0 To UBound(a_strLines)
        a_strLines(lngIdx) = Trim$(a_strLines(lngIdx))
        If (Len(a_strLines(lngIdx)) > 0) And (a_strLines(lngIdx) <> "=") Then
            lngPos = InStr(a_strLines(lngIdx), "=")
            If lngPos > 1 Then
                strKey = Trim$(Left$(a_strLines(lngIdx), lngPos - 1))
                strVal = Trim$(Right$(a_strLines(lngIdx), Len(a_strLines(lngIdx)) - lngPos))
            Else
                strKey = Trim$(a_strLines(lngIdx))
                strVal = ""
            End If
            If strKey <> "" Then
                ReDim Preserve a_udtKVs_out(GetKVsFromString)
                a_udtKVs_out(GetKVsFromString).strKey = strKey
                a_udtKVs_out(GetKVsFromString).strValue = strVal
                GetKVsFromString = GetKVsFromString + 1
            End If
        End If
    Next lngIdx

End Function


Public Function GetKVLinesFromString(ByVal strKVs_in As String, a_udtKVLines_out() As NotesKeysLinesType)

    'converts a string containing lines of key=value pairs, into an array of lines,
    'some of which may be key=value pairs, some maybe comments, blank, etc

    Dim lngIdx As Long
    Dim lngPos As Long
    Dim a_strLines() As String

    GetKVLinesFromString = 0
    ReDim a_udtKVLines_out(0)

    If Trim$(strKVs_in) = "" Then Exit Function

    a_strLines = Split(strKVs_in, vbCrLf)
    For lngIdx = 0 To UBound(a_strLines)
        
        ReDim Preserve a_udtKVLines_out(GetKVLinesFromString)

        With a_udtKVLines_out(GetKVLinesFromString)

            .strLine = a_strLines(lngIdx)
            .blnIsKV = False
            .strKey = ""
            .strValue = ""
            .blnFlagged = False
            
            a_strLines(lngIdx) = Trim$(a_strLines(lngIdx))
            lngPos = InStr(a_strLines(lngIdx), "=")
            If lngPos > 1 Then
                .strKey = Trim$(Left$(a_strLines(lngIdx), lngPos - 1))
                .strValue = Trim$(Right$(a_strLines(lngIdx), Len(a_strLines(lngIdx)) - lngPos))
                .blnIsKV = True
            End If
    
        End With

        GetKVLinesFromString = GetKVLinesFromString + 1

    Next lngIdx

End Function
Public Function RenameKVSet(strCurrentSetName_in As String, strNewSetName_in As String, strError_out As String) As Boolean

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim strLine As String
    Dim strTrimmedLine As String
    Dim strToWrite As String

    strToWrite = ""
    
    If LCase$(strCurrentSetName_in) = "global" Or LCase$(strNewSetName_in) = "global" Then
        strError_out = "The KVSet cannot be renamed from/to ""global""."
        RenameKVSet = False
        Exit Function
    End If

    'if no Gemus.ini file then it can't be renamed!
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then
        strError_out = "The KVSet could not be found."
        RenameKVSet = False
        Exit Function
    Else
        'read in the lines
        f = FreeFile()
        Open strFile For Input As #f
        Do Until EOF(f)
            Line Input #f, strLine
            strTrimmedLine = Trim$(strLine)
            If LCase$(strTrimmedLine) = "[" & LCase$(strCurrentSetName_in) & "]" Then
                strToWrite = strToWrite & "[" & strNewSetName_in & "]" & vbCrLf
            Else
                strToWrite = strToWrite & strLine & vbCrLf
            End If
        Loop
        Close #f
    End If

    'write the file
    f = FreeFile()
    Open strFile For Output As #f
    Print #f, strToWrite
    Close #f

    strError_out = ""
    RenameKVSet = True
    Exit Function

Err_Handler:
    strError_out = Err.Description & " (" & Err.Number & ")"
    RenameKVSet = False

End Function
Public Function DeleteKVSet(strSetName_in As String, strError_out As String) As Boolean

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim strLine As String
    Dim strTrimmedLine As String
    Dim strToWrite As String
    Dim blnInTHESet As Boolean
    Dim strLCaseSetName As String

    strToWrite = ""
    strError_out = ""
    strLCaseSetName = LCase$(strSetName_in)

    'if no Gemus.ini file then use the incoming kvset as the whole contents
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then
        DeleteKVSet = True
        Exit Function
    Else
        'read in the lines
        blnInTHESet = False 'we're not in THE set until we encounter it
        f = FreeFile()
        Open strFile For Input As #f
        Do Until EOF(f)
            Line Input #f, strLine
            strTrimmedLine = LCase$(Trim$(strLine))
            If strTrimmedLine = "[" & strLCaseSetName & "]" Then
                'We're in THE set! Don't write it back to the file!
                blnInTHESet = True
            Else
                'Are we still in the set?
                If Left$(strTrimmedLine, 1) = "[" And Right$(strTrimmedLine, 1) = "]" Then
                    blnInTHESet = False
                End If
                If Not blnInTHESet Then
                    'add line if not in THE set
                    strToWrite = strToWrite & strLine & vbCrLf
                End If
            End If
        Loop
        Close #f
    End If

    'write the file
    f = FreeFile()
    Open strFile For Output As #f
    Print #f, strToWrite
    Close #f

    strError_out = ""
    DeleteKVSet = True
    Exit Function

Err_Handler:
    strError_out = Err.Description & " (" & Err.Number & ")"
    DeleteKVSet = False

End Function
Public Function SaveKVSet(strSetName_in As String, strKVText_in As String, strError_out As String) As Boolean

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim strFile As String
    Dim strLine As String
    Dim strTrimmedLine As String
    Dim strToWrite As String
    Dim strText As String
    Dim blnInTHESet As Boolean
    Dim blnDoneTHESet As Boolean

    strToWrite = ""

    'the lines we want to save as a big long multiline string
    strText = "[" & LCase$(strSetName_in) & "]" & vbCrLf & strKVText_in

    'if no Gemus.ini file then use the incoming kvset as the whole contents
    strFile = m_strAppPathWithBackSlash & m_strGBName & "\Gemus.ini"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then
        strToWrite = strText
        blnDoneTHESet = True
    Else
        'read in the lines
        blnInTHESet = False 'we're not in THE set until we encounter it
        blnDoneTHESet = False
        f = FreeFile()
        Open strFile For Input As #f
        Do Until EOF(f)
            Line Input #f, strLine
            strTrimmedLine = Trim$(strLine)
            If LCase$(strTrimmedLine) = "[" & LCase$(strSetName_in) & "]" Then
                'We're in THE set!
                blnInTHESet = True
                blnDoneTHESet = True
                strToWrite = strToWrite & strText & vbCrLf
            Else
                'Are we still in the set?
                If Left$(strTrimmedLine, 1) = "[" And Right$(strTrimmedLine, 1) = "]" Then
                    blnInTHESet = False
                End If
                If Not blnInTHESet Then
                    'add line if not in THE set
                    strToWrite = strToWrite & strLine & vbCrLf
                End If
            End If
        Loop
        Close #f
    End If

    'add the set to the file if not already in there
    If Not blnDoneTHESet Then
        strToWrite = strToWrite & strText & vbCrLf
    End If

    'write the file
    f = FreeFile()
    Open strFile For Output As #f
    Print #f, strToWrite
    Close #f

    strError_out = ""
    SaveKVSet = True
    Exit Function

Err_Handler:
    strError_out = Err.Description & " (" & Err.Number & ")"
    SaveKVSet = False

End Function
Private Sub LoadPaths()

    Dim strSidPath As String            'sid path retrieval
    Dim strPhotoPath As String          'photo path retrieval
    Dim lngNumFiles As Long             'for emulator files retrieval
    Dim strFiles() As String            'for emulator files retrieval
    Dim strPaths() As String            'paths retrieval array
    Dim strNames() As String
    Dim strExes() As String             'exes retrieval array
    Dim strTypes As String              'for music supported filetypes
    Dim a_strTypes() As String          'for music supported filetypes
    Dim strCLP() As String              'command line parameters array
    Dim strScript() As String           'script string array
    Dim blnUseShort() As Boolean        'short filename flags
    Dim lngPathIdx As Long              'loop index
    Dim lngFileIdx As Long              'loop index
    Dim strPath As String
    Dim lngTotalTypes As Long
    Dim strSection As String
    Dim strSections() As String
    Dim lngPlayerIdx As Long

    'Music Path
    strSidPath = m_xobjFunc.INI_GetSetting("Misc", "Music", "", m_strPathFile)
    Call GetPathInfo(strSidPath, m_udtRegPathSid)

    'Photos Path
    strPhotoPath = m_xobjFunc.INI_GetSetting("Misc", "Photos", "", m_strPathFile)
    Call GetPathInfo(strPhotoPath, m_udtRegPathPhotos)
    
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
            Exit For
        End If
    Next lngPathIdx
        
    'picture paths
    lngPathIdx = 1
    m_lngTotalScrnShotPaths = 0
    ReDim strPaths(0)
    Do
        strPath = m_xobjFunc.INI_GetSetting("Pictures", Trim$(lngPathIdx), "", m_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        m_lngTotalScrnShotPaths = lngPathIdx - 1
        ReDim ma_udtRegPathScrnShot(m_lngTotalScrnShotPaths - 1)
        For lngPathIdx = 0 To (m_lngTotalScrnShotPaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ma_udtRegPathScrnShot(lngPathIdx))
        Next lngPathIdx
    End If

    'extras paths
    lngPathIdx = 1
    m_lngTotalExtrasPaths = 0
    ReDim strPaths(0)
    Do
        strPath = m_xobjFunc.INI_GetSetting("Extras", Trim$(lngPathIdx), "", m_strPathFile)
        If strPath <> "" Then
            ReDim Preserve strPaths(lngPathIdx - 1)
            strPaths(lngPathIdx - 1) = strPath
            lngPathIdx = lngPathIdx + 1
        Else
            Exit Do
        End If
    Loop
    If lngPathIdx > 1 Then
        m_lngTotalExtrasPaths = lngPathIdx - 1
        ReDim ma_udtRegPathExtras(m_lngTotalExtrasPaths - 1)
        For lngPathIdx = 0 To (m_lngTotalExtrasPaths - 1)
            Call GetPathInfo(strPaths(lngPathIdx), ma_udtRegPathExtras(lngPathIdx))
        Next lngPathIdx
    End If

End Sub
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
                .strRealPath = Left$(m_strAppPathWithBackSlash, 3) & Right(.strRealPath, (iLen - 1))
                iLen = Len(.strRealPath)
            End If
            
            'convert relative path to absolute
            If Mid$(.strRealPath, 2, 1) <> ":" Then
                .strRealPath = m_strAppPathWithBackSlash & .strRealPath
                iLen = Len(.strRealPath)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If m_xobjFunc.DirExists(.strRealPath) = True Then
                    .strRealPath = CurDir()
                End If
            End If
                    
            'if specified path does not exist, try to validate on the current drive
            If m_xobjFunc.DirExists(.strRealPath) = False Then
                strValidateLocal = Left$(m_strAppPathWithBackSlash, 3) & Right(.strRealPath, (iLen - 3))
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
Private Function ParseSendKeys(ByVal strSendKeys_in As String, a_udtSendBlocks_out() As SendKeysBlockType) As Long

    On Error GoTo 0

    Dim strKeys As String
    Dim lngOpenPos As Long
    Dim lngClosePos As Long
    Dim lngStart As Long
    Dim lngIdx As Long
    Dim lngNumSecs As Long
    Dim strRemainder As String

    ParseSendKeys = 0
    ReDim a_udtSendBlocks_out(0)
    If strSendKeys_in = "" Then Exit Function
    
    'replace gemus environment vars
    strKeys = strSendKeys_in
    Call ReplaceEnvironmentVars(strKeys, False, True)

    'replace double pipes
    strKeys = Replace(strKeys, "{DOUBLEPIPE}", "||")

    'look for all [[x]] "seconds to wait"
    lngStart = 1
    Do
        lngOpenPos = InStr(lngStart, strKeys, "[")
        If lngOpenPos > 0 Then
            'found next open "["
            If lngOpenPos > lngStart Then
                'grab from start up to the found "["
                ReDim Preserve a_udtSendBlocks_out(ParseSendKeys)
                a_udtSendBlocks_out(ParseSendKeys).lngSecsToWait = 0
                a_udtSendBlocks_out(ParseSendKeys).strKeysToSend = Mid$(strKeys, lngStart, lngOpenPos - lngStart)
                ParseSendKeys = ParseSendKeys + 1
            End If
            'find closing "]"
            lngClosePos = InStr(lngOpenPos + 1, strKeys, "]")
            If lngClosePos > 0 Then
                'found closing "]", grab the number of seconds to wait
                If lngClosePos > (lngOpenPos + 1) Then
                    lngNumSecs = Val(Mid$(strKeys, lngOpenPos + 1, (lngClosePos - lngOpenPos) - 1))
                    If lngNumSecs <> 0 Then
                        ReDim Preserve a_udtSendBlocks_out(ParseSendKeys)
                        a_udtSendBlocks_out(ParseSendKeys).lngSecsToWait = lngNumSecs
                        a_udtSendBlocks_out(ParseSendKeys).strKeysToSend = ""
                        ParseSendKeys = ParseSendKeys + 1
                    End If
                End If
                'start again at the first char after the closing "]"
                lngStart = lngClosePos + 1
            Else
                'no closing "]" - bad strSendKeys_in param!
                Exit Do
            End If
        Else
            'lngStart is start of last sendstring
            strRemainder = Right$(strKeys, Len(strKeys) - (lngStart - 1))
            If strRemainder <> "" Then
                ReDim Preserve a_udtSendBlocks_out(ParseSendKeys)
                a_udtSendBlocks_out(ParseSendKeys).lngSecsToWait = 0
                a_udtSendBlocks_out(ParseSendKeys).strKeysToSend = strRemainder
                ParseSendKeys = ParseSendKeys + 1
            End If
            Exit Do
        End If
    Loop

    'replace {OPENSQB} {CLOSESQB} with "[" and "]"
    For lngIdx = 0 To ParseSendKeys - 1
        With a_udtSendBlocks_out(lngIdx)
            If .lngSecsToWait = 0 Then
                .strKeysToSend = Replace(.strKeysToSend, "{OPENSQB}", "[")
                .strKeysToSend = Replace(.strKeysToSend, "{CLOSESQB}", "]")
                'MsgBox "SENDING: " & .strKeysToSend
            Else
                'MsgBox "WAITING: " & .lngSecsToWait
            End If
        End With
    Next lngIdx

End Function

Public Sub SetReZipLegacy(ByVal blnUseLegacy_in As Boolean)

    m_blnRezipLegacy = blnUseLegacy_in

End Sub

Public Sub SetRezipNotify(ByVal blnNotify_in As Boolean)

    m_blnNotifyUserIfRepackWasOK = blnNotify_in

End Sub

Public Sub SetSilentMode(ByVal blnSilentMode_in As Boolean)

    m_blnSilentMode = blnSilentMode_in

End Sub

Public Sub SetWaitIfNotRezipping(ByVal blnDoWait_in As Boolean)

    m_blnWaitIfNotRezipping = blnDoWait_in

End Sub

Private Function ShellExecuteGameFile(strFile_in As String, strCLP_in As String, lngHandle_out As Long) As Boolean

    Dim SEI As SHELLEXECUTEINFO     'api structure

    'Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI
        .hWnd = GetDesktopWindow()
        .lpVerb = "open"
        .lpFile = strFile_in
        If strCLP_in <> "" Then
            .lpParameters = strCLP_in
        Else
            .lpParameters = vbNullChar
        End If
        .lpDirectory = vbNullChar
        .nShow = SW_SHOW
        .hInstApp = 0
        .lpIDList = 0
   End With

    'run the file
    If ShellExecuteEx(SEI) = 0 Then
        ShellExecuteGameFile = False
    Else
        ShellExecuteGameFile = True
        'wait till finished?
        lngHandle_out = SEI.hProcess
    End If

End Function
Private Function GetUnzippedFiles(strPathWBS_in As String, a_UZF_out() As UnzippedFileType) As Long

    'returns the total number of files in the array, or -1 on error, -2 on invalid (zip contains a folder)

    Dim lngTotal As Long
    Dim lngBytes As Long
    Dim CRC As clsCRC
    Dim strFile As String

'    On Error GoTo Err_Handler
    
    Set CRC = New clsCRC
    lngTotal = 0
    ReDim a_UZF_out(0)

    'loop through the files in the folder
    strFile = Dir(strPathWBS_in & "*.*")
    While strFile <> ""
        If strFile <> "." And strFile <> ".." Then
            If Not (GetAttr(strPathWBS_in & strFile) And vbDirectory) Then
                'store the file details
                ReDim Preserve a_UZF_out(lngTotal)
                a_UZF_out(lngTotal).strName = strFile
                a_UZF_out(lngTotal).strUCaseName = UCase$(strFile)
                a_UZF_out(lngTotal).strCRC = CRC.CalculateFile(strPathWBS_in & strFile)
                a_UZF_out(lngTotal).blnChecked = False
                lngTotal = lngTotal + 1
            Else
                'no sub directories allowed
                GetUnzippedFiles = -2
                Exit Function
            End If
        End If
        strFile = Dir()
    Wend

    GetUnzippedFiles = lngTotal
    Set CRC = Nothing
    Exit Function

Err_Handler:
    GetUnzippedFiles = -1
    Set CRC = Nothing
    

End Function


Public Function SimulateRunEmulator(strKV_out As String, ByVal lngEmuIdx_in As Long, strGamePathFile_in As String, ByVal lngImageIndex_in As Long, _
                            strDefaultFile_inout As String, strExtractDestFolder_in As String, ByVal blnFromDB_in As Boolean, _
                            Optional ByVal varNumPlayersFrom_in As Variant, Optional ByVal varNumPlayersTo_in As Variant, _
                            Optional varControl_in As Variant, Optional varGameComment_in As Variant, _
                            Optional varVersionComment_in As Variant, Optional varPALNTSC_in As Variant, _
                            Optional ByVal varTrueDriveEmu_in As Variant, Optional varNotes_in As Variant) As String

    m_blnSimulateRun = True
    m_strSimulateKV = ""
    SimulateRunEmulator = RunEmulator(lngEmuIdx_in, strGamePathFile_in, lngImageIndex_in, _
                            strDefaultFile_inout, strExtractDestFolder_in, blnFromDB_in, _
                            varNumPlayersFrom_in, varNumPlayersTo_in, _
                            varControl_in, varGameComment_in, _
                            varVersionComment_in, varPALNTSC_in, _
                            varTrueDriveEmu_in, varNotes_in)

    m_blnSimulateRun = False
    strKV_out = m_strSimulateKV

End Function

Private Function UnzippedFilesChanged(a_Before_in() As UnzippedFileType, a_After_in() As UnzippedFileType, a_Rezip_out() As UnzippedFileType) As String

    'Compares the passed in 2 arrays of files and returns an empty string ("")
    'if they are different, "~NOCHANGE~" if there is no change, or an error message.
    'Returns the list of files to rezip in the a_Rezip_out() array.

    Dim lngIdxBefore As Long
    Dim lngIdxAfter As Long
    Dim lngTotalBefore As Long
    Dim lngTotalAfter As Long
    Dim lngTotalMatched As Long
    Dim blnChanged As Boolean
    Dim intFN As Integer

    On Error GoTo Err_Handler

    'are there less files in the after array
    lngTotalBefore = UBound(a_Before_in) + 1
    lngTotalAfter = UBound(a_After_in) + 1
    If lngTotalAfter < lngTotalBefore Then
        UnzippedFilesChanged = GetINIMessage("Running Games", "ErrorRezipFilesRemoved", "Some files in the ReZip folder have been removed!")
        Exit Function
    End If

    'match each file in the array
    For lngIdxAfter = 0 To (lngTotalAfter - 1)
        a_After_in(lngIdxAfter).blnChecked = False
    Next lngIdxAfter
    lngTotalMatched = 0
    blnChanged = False
    For lngIdxBefore = 0 To (lngTotalBefore - 1)
        With a_Before_in(lngIdxBefore)
            For lngIdxAfter = 0 To (lngTotalAfter - 1)
                If a_After_in(lngIdxAfter).blnChecked = False Then
                    If a_After_in(lngIdxAfter).strUCaseName = .strUCaseName Then
                        If a_After_in(lngIdxAfter).strCRC <> .strCRC Then
                            blnChanged = True
                        End If
                        a_After_in(lngIdxAfter).blnChecked = True
                        lngTotalMatched = lngTotalMatched + 1
                        Exit For
                    End If
                End If
            Next lngIdxAfter
        End With
    Next lngIdxBefore

    'what happened?
    If lngTotalMatched < lngTotalBefore Then
        'some source files were missing
        UnzippedFilesChanged = GetINIMessage("Running Games", "ErrorRezipFilesRemoved", "Some files in the ReZip folder have been removed!")
        Exit Function
    Else
        If (blnChanged = True) Or (lngTotalAfter > lngTotalBefore) Then
            'copy the rezip files to the output array
            ReDim a_Rezip_out(lngTotalAfter - 1)
            For lngIdxAfter = 0 To (lngTotalAfter - 1)
                a_Rezip_out(lngIdxAfter).strName = a_After_in(lngIdxAfter).strName
                a_Rezip_out(lngIdxAfter).blnChecked = a_After_in(lngIdxAfter).blnChecked
                a_Rezip_out(lngIdxAfter).strCRC = a_After_in(lngIdxAfter).strCRC
            Next lngIdxAfter
            UnzippedFilesChanged = ""
        Else
            UnzippedFilesChanged = "~NOCHANGE~"
        End If
    End If

    Exit Function

Err_Handler:
    UnzippedFilesChanged = Err.Description & "(Error " & Err.Number & ")"

End Function


Private Function DoRezip(a_Before_in() As UnzippedFileType, strRezipFolderWBS_in As String, strZipFile_in As String) As String

    'Performs a rezip operation.  Returns empty string on success, or an error message.
    'Assumes (if called) that the user wants to rezip the games.

    Dim a_After() As UnzippedFileType
    Dim a_Rezip() As UnzippedFileType
    Dim strError As String
    Dim lngRetVal As Long
    Dim lngTotalFiles As Long
    Dim lngIdx As Long
    Dim blnZip As Boolean
    Dim lngIdx2 As Long
    Dim strCRC As String
    Dim strTempFile As String
    On Error GoTo Err_Handler

    'get the latest file list from the rezip folder
    lngRetVal = GetUnzippedFiles(strRezipFolderWBS_in, a_After())
    If lngRetVal = -1 Then
        DoRezip = GetINIMessage("Running Games", "ErrorGettingPostRezipFiles", "Unable to get the post-file list from the ReZip folder")
        Exit Function
    ElseIf lngRetVal = -2 Then
        DoRezip = GetINIMessage("Running Games", "InvalidFilesInRezipFolder", "Rezipping cancelled - the zip file contains a subfolder which is currently not supported.")
        Exit Function
    End If

    'compare the files in the rezip folder to see if they have changed and need rezipping
    strError = UnzippedFilesChanged(a_Before_in(), a_After(), a_Rezip())
    If strError = "~NOCHANGE~" Then
        DoRezip = ""
        Exit Function
    ElseIf strError <> "" Then
        DoRezip = strError
        Exit Function
    End If
    'If here, the files in the rezip folder have changed and we need to rezip.
    'The files to rezip are stored in the a_Rezip() array.

    'is the user sure?
    If m_blnAskUserBeforeRepack = True Then
        If MsgBox(GetINIMessage("Running Games", "RezipConfirm", "There have been changes made to the game file(s).  Do you want to rezip them") & "?", vbQuestion Or vbYesNo, "GEMUS") = vbNo Then
            DoRezip = ""
            Exit Function
        End If
    End If

    'Are there additional files to rezip?
    If UBound(a_Before_in) < UBound(a_Rezip) Then
        If me_RepackAddtional = eRezipAdditionalZip Then
            'rezip all additional files
            lngTotalFiles = UBound(a_Rezip) + 1
            For lngIdx = 0 To (lngTotalFiles - 1)
                a_Rezip(lngIdx).blnChecked = True
            Next lngIdx
        ElseIf me_RepackAddtional = eRezipAdditionalAsk Then
            'ask the user which files to rezip
            Load frmRezipAsk
            With frmRezipAsk
                .lblPrompt = GetINIMessage("Running Games", "RezipAsk", "Please select the files you wish to Rezip") & ":"
                For lngIdx = 0 To UBound(a_Rezip)
                    With a_Rezip(lngIdx)
                        frmRezipAsk.AddFile .strName, .strUCaseName, .strCRC, .blnChecked
                    End With
                Next lngIdx
                .Show vbModal
                lngTotalFiles = .lngTotalFiles
                If lngTotalFiles > 0 Then
                    ReDim a_Rezip(lngTotalFiles - 1)
                    For lngIdx = 0 To (lngTotalFiles - 1)
                        With a_Rezip(lngIdx)
                            frmRezipAsk.GetFile lngIdx, .strName, .strUCaseName, .strCRC, .blnChecked
                        End With
                    Next lngIdx
                End If
            End With
            Unload frmRezipAsk
        Else
            'ignore additional files
            lngTotalFiles = UBound(a_Rezip) + 1
        End If
    Else
        'same number of files
        lngTotalFiles = UBound(a_Rezip) + 1
    End If

    'if there are some files to rezip then rezip'em!
    If lngTotalFiles > 0 Then
        blnZip = False
        For lngIdx = 0 To (lngTotalFiles - 1)
            If a_Rezip(lngIdx).blnChecked = True Then
                blnZip = True
                Exit For
            End If
        Next lngIdx
        'no duplicate crc's
'        If blnZip = True Then
'            If lngTotalFiles > 1 Then
'                For lngIdx = 0 To (lngTotalFiles - 1)
'                    If a_Rezip(lngIdx).blnChecked = True Then
'                        strCRC = a_Rezip(lngIdx).strCRC
'                        For lngIdx2 = 0 To (lngTotalFiles - 1)
'                            If lngIdx <> lngIdx2 Then
'                                If a_Rezip(lngIdx2).blnChecked = True Then
'                                    If a_Rezip(lngIdx2).strCRC = strCRC Then
'                                        blnZip = False
'                                        Exit For
'                                    End If
'                                End If
'                            End If
'                        Next lngIdx2
'                        If blnZip = False Then
'                            Exit For
'                        End If
'                    End If
'                Next lngIdx
'                If blnZip = False Then
'                    'duplicate crc's
'                    DoRezip = "The Rezip could not take place as there are multiple files to rezip with the same CRC."
'                    Exit Function
'                End If
'            End If
'        End If
        If blnZip = True Then
            'zip all files into the rezip folder under the temp filename
            strTempFile = strRezipFolderWBS_in & "gb_temp_rezip"
            strTempFile = strTempFile & Trim$(Int((1239833 - 10293 + 1) * Rnd + 10293))
            strTempFile = strTempFile & ".zip"
            Set m_objZipper = New cZip
            With m_objZipper
                .ZipFile = strTempFile
                .StoreFolderNames = False
                .RecurseSubDirs = False
                .ClearFileSpecs
                For lngIdx = 0 To (lngTotalFiles - 1)
                    If a_Rezip(lngIdx).blnChecked = True Then
                        .AddFileSpec strRezipFolderWBS_in & a_Rezip(lngIdx).strName
                    End If
                Next lngIdx
                .Zip
            End With
            Set m_objZipper = Nothing
            'rename existing gamefile to .old
            Call m_xobjFunc.DeletePathSilent(strZipFile_in & ".old")
            Name (strZipFile_in) As (strZipFile_in & ".old")
            'rename the temp zip file as the game zipfile
            Name (strTempFile) As (strZipFile_in)
            'delete the .old file
            Call m_xobjFunc.DeletePathSilent(strZipFile_in & ".old")
        End If
    End If

    'no error
    If m_blnNotifyUserIfRepackWasOK = True Then
        DoRezip = GetINIMessage("Running Games", "RezipSuccess", "The game was rezipped successfully!")
    Else
        DoRezip = ""
    End If

    Exit Function
    
Err_Handler:
    DoRezip = "Rezip operation Failed: " & vbCrLf & vbCrLf & Err.Description & "(Error " & Err.Number & ")"

End Function
Public Function EditEmulator(ByVal lngEmuIdx_in As Long) As Boolean

    Dim lngFileIdx As Long  'loop index
    Dim blnOK As Boolean

    With g_udtEmulatorProps
    
        EditEmulator = False
        .blnUseShort = ma_udtEmulators(lngEmuIdx_in).blnUseShortFilenames
        .lngNumFiles = ma_udtEmulators(lngEmuIdx_in).lngNumFiles
        .strFiles() = ma_udtEmulators(lngEmuIdx_in).strFiles()
        .strFullPath = ma_udtEmulators(lngEmuIdx_in).strFullPathAndFilename
        .strName = ma_udtEmulators(lngEmuIdx_in).strName
        .strScript = ma_udtEmulators(lngEmuIdx_in).strScript
        .blnUseDefaultTypes = ma_udtEmulators(lngEmuIdx_in).blnUseDefaultTypes
        .strTypes = ma_udtEmulators(lngEmuIdx_in).strTypes
        .lngWait = ma_udtEmulators(lngEmuIdx_in).lngWait
        .strGBName = m_strGBName
        .strEmuFile = m_strEmuFile
        frmEmulatorProps.Show vbModal
        If .strName <> "" Then
            'delete old section
            m_xobjFunc.INI_DeleteSection ma_udtEmulators(lngEmuIdx_in).strName, m_strEmuFile
            'save changed data
            ma_udtEmulators(lngEmuIdx_in).blnUseShortFilenames = .blnUseShort
            ma_udtEmulators(lngEmuIdx_in).lngNumFiles = .lngNumFiles
            ma_udtEmulators(lngEmuIdx_in).strFiles() = .strFiles()
            ma_udtEmulators(lngEmuIdx_in).strFullPathAndFilename = .strFullPath
            ma_udtEmulators(lngEmuIdx_in).strName = .strName
            ma_udtEmulators(lngEmuIdx_in).strScript = .strScript
            ma_udtEmulators(lngEmuIdx_in).blnUseDefaultTypes = .blnUseDefaultTypes
            ma_udtEmulators(lngEmuIdx_in).strTypes = .strTypes
            ma_udtEmulators(lngEmuIdx_in).lngWait = .lngWait
            Call GetEmulatorInfo(.strScript, .blnUseShort, .strName, .strFullPath, ma_udtEmulators(lngEmuIdx_in), .strFiles(), .lngNumFiles, .blnUseDefaultTypes, .strTypes, .lngWait)
            With m_xobjFunc
                blnOK = .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "UseShortFilenames", IIf((ma_udtEmulators(lngEmuIdx_in).blnUseShortFilenames = True), "1", "0"), m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "UseDefaultTypes", IIf((ma_udtEmulators(lngEmuIdx_in).blnUseDefaultTypes = True), "1", "0"), m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "SupportedTypes", ma_udtEmulators(lngEmuIdx_in).strTypes, m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "EmulatorPathFile", ma_udtEmulators(lngEmuIdx_in).strFullPathAndFilename, m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "ScriptFile", ma_udtEmulators(lngEmuIdx_in).strScript, m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, "WaitFinish", Trim$(ma_udtEmulators(lngEmuIdx_in).lngWait), m_strEmuFile)
                If ma_udtEmulators(lngEmuIdx_in).lngNumFiles > 0 Then
                    For lngFileIdx = 0 To (ma_udtEmulators(lngEmuIdx_in).lngNumFiles - 1)
                        blnOK = blnOK And .INI_WriteSetting(ma_udtEmulators(lngEmuIdx_in).strName, Trim$(lngFileIdx + 1), ma_udtEmulators(lngEmuIdx_in).strFiles(lngFileIdx), m_strEmuFile)
                    Next lngFileIdx
                End If
                blnOK = blnOK And .INI_DeleteNumberedValues(ma_udtEmulators(lngEmuIdx_in).strName, ma_udtEmulators(lngEmuIdx_in).lngNumFiles + 1, m_strEmuFile)
                blnOK = blnOK And .INI_WriteSetting("Emulators", Trim$(lngEmuIdx_in + 1), ma_udtEmulators(lngEmuIdx_in).strName, m_strEmuFile)
            End With
            If blnOK = False Then
                MsgBox "There was an error saving the Emulator information to " & m_strEmuFile, vbExclamation, "GEMUS"
            End If
            EditEmulator = True
        End If
    End With

End Function

Public Sub EmulatorFiles(ByVal lngEmulatorIdx_in As Long, strFiles_out() As String)

    strFiles_out() = ma_udtEmulators(lngEmulatorIdx_in).strFiles()

End Sub

Public Sub ExportEmulatorDefinitions()

    'export the current setup to GEMUS.DAT 1.0 format emulator definitions

    Dim a_strEmus() As String
    Dim lngTotalEmus As Long
    Dim lngIdx As Long
    Dim strEmu As String
    Dim strExport As String
    Dim objCRC As clsCRC
    Dim strPath As String
    Dim strFile As String
    Dim strTemp As String
    Dim strAssFile As String
    Dim lngAssFileIdx As Long
    Dim strEmuPath As String
    Dim strExportFile As String
    Dim intFN As Integer
    Dim strScript As String

    If m_lngTotalEmulators = 0 Then Exit Sub
  
    On Error GoTo Err_Handler
    lngTotalEmus = 0
    ReDim a_strEmus(0)

    With m_xobjFunc
        'kill the existing file if it exists
        strExportFile = m_strAppPathWithBackSlash & m_strGBName & "\emudefs.txt"
        If .FileExists(strExportFile, False, "") = True Then
            Call .DeletePathSilent(strExportFile)
        End If
        'get list of emulators in the right order
        strEmu = .INI_GetSetting("Emulators", Trim$(lngTotalEmus + 1), "", m_strEmuFile)
        Do Until strEmu = ""
            ReDim Preserve a_strEmus(lngTotalEmus)
            a_strEmus(lngTotalEmus) = strEmu
            lngTotalEmus = lngTotalEmus + 1
            strEmu = .INI_GetSetting("Emulators", Trim$(lngTotalEmus + 1), "", m_strEmuFile)
        Loop
        If lngTotalEmus = 0 Then Exit Sub
        'Add each line to the export string
        strExport = ""
        Set objCRC = New clsCRC
        For lngIdx = 0 To (lngTotalEmus - 1)
            strEmu = a_strEmus(lngIdx)
            '¬name¬url¬comment¬
            strExport = strExport & "¬" & strEmu & "¬¬¬"
            'filename of emulator executable
            strTemp = .INI_GetSetting(strEmu, "EmulatorPathFile", "", m_strEmuFile)
            Call .GetFileAndPath(strTemp, strPath, strFile)
            strExport = strExport & strFile & "¬"
            'crc
            If .FileExists(strTemp, False, "") = True Then
                strExport = strExport & Hex(objCRC.CalculateFile(strTemp)) & "¬"
            Else
                strExport = strExport & "*¬"
            End If
            'emulator filetypes
            strExport = strExport & .INI_GetSetting(strEmu, "SupportedTypes", "", m_strEmuFile) & "¬"
            'use short flag
            strExport = strExport & IIf((Val(.INI_GetSetting(strEmu, "UseShortFilenames", "1", m_strEmuFile)) = 1), "Y", "N") & "¬"
            'script file
            strScript = .INI_GetSetting(strEmu, "ScriptFile", "", m_strEmuFile)
            strExport = strExport & strScript & "¬"
            'associated files
            lngAssFileIdx = 1
            strEmuPath = strPath
            strAssFile = .INI_GetSetting(strEmu, Trim$(lngAssFileIdx), "", m_strEmuFile)
            Do Until strAssFile = ""
                'get the associated file
                Call .GetFileAndPath(strAssFile, strPath, strFile)
                If UCase$(Left$(strPath, Len(strEmuPath))) = UCase$(strEmuPath) Then
                    If Len(strEmuPath) = Len(strPath) Then
                        'ass file path is emu path
                        strExport = strExport & strFile & "¬"
                    Else
                        If Mid$(strPath, Len(strEmuPath) + 1, 1) = "\" Then
                            'associated file path is sub-folder of emu path
                            strPath = Right$(strPath, Len(strPath) - Len(strEmuPath))
                            strPath = Right$(strPath, Len(strPath) - 1)
                            If Right$(strPath, 1) <> "\" Then
                                strPath = strPath & "\"
                            End If
                            strExport = strExport & strPath & strFile & "¬"
                        End If
                    End If
                End If
                'goto the next one
                lngAssFileIdx = lngAssFileIdx + 1
                strAssFile = .INI_GetSetting(strEmu, Trim$(lngAssFileIdx), "", m_strEmuFile)
            Loop
            'next line
            strExport = strExport & vbCrLf
        Next lngIdx
        Set objCRC = Nothing
        'write the file and open it in notepad or default viewer
        If Len(strExport) > 0 Then
            intFN = FreeFile()
            Open strExportFile For Output As #intFN
            Print #intFN, strExport
            Close #intFN
            Call .ShellExecuteFile(strExportFile, "", False)
        End If
    End With

    Exit Sub
    
Err_Handler:
    Set objCRC = Nothing

End Sub

Private Function FileTypeSupported(strFileTypes_in As String, strExt_in As String) As Boolean

    Dim a_strTypes() As String
    Dim a_strValidatedTypes() As String
    Dim lngFTIdx As Long
    Dim lngTotalTypes As Long
    Dim strUpper As String

    lngTotalTypes = 0
    ReDim a_strValidatedTypes(0)

    If strFileTypes_in <> "" Then
        'construct the array to check
        a_strTypes = Split(strFileTypes_in, ";")
        For lngFTIdx = 0 To UBound(a_strTypes)
            a_strTypes(lngFTIdx) = Trim$(a_strTypes(lngFTIdx))
            If ValidateEmuFileType(a_strTypes(lngFTIdx)) = True Then
                ReDim Preserve a_strValidatedTypes(lngTotalTypes)
                a_strValidatedTypes(lngTotalTypes) = UCase$(a_strTypes(lngFTIdx))
                lngTotalTypes = lngTotalTypes + 1
            End If
        Next lngFTIdx
        'check the array
        strUpper = UCase$(strExt_in)
        For lngFTIdx = 0 To UBound(a_strValidatedTypes)
            If a_strValidatedTypes(lngFTIdx) = strUpper Then
                FileTypeSupported = True
                Exit Function
            End If
        Next lngFTIdx
        FileTypeSupported = False
    Else
        FileTypeSupported = False
    End If

End Function

Public Sub GetManualFieldsDefaults(udtFields_out As ManualGEMUSFields, strGBName_in As String)

    With m_xobjRegistry
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "TDE", udtFields_out.blnTDE, False, False
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Control", udtFields_out.lngControl, False, 0
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "PalNTSC", udtFields_out.lngPalNTSC, False, 0
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Players", udtFields_out.lngPlayers, False, 1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Comment", udtFields_out.strComment, True, ""
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "VComment", udtFields_out.strVComment, True, ""
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Notes", udtFields_out.strNotes, True, ""
    End With

End Sub

Public Function GetManualFieldsFromUser(udtFields_out As ManualGEMUSFields, blnSaveAsDefaults_in As Boolean, strGBName_in As String) As Boolean

    '-----------------------------------------------------------------
    'Lets the user specify the script fields to use when running games
    'Returns True if the user pressed OK, False otherwise.
    '-----------------------------------------------------------------

    If blnSaveAsDefaults_in = True Then
        With m_xobjRegistry
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "TDE", g_udtManualFields.blnTDE, False, False
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Control", g_udtManualFields.lngControl, False, 0
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "PalNTSC", g_udtManualFields.lngPalNTSC, False, 0
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Players", g_udtManualFields.lngPlayers, False, 1
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Comment", g_udtManualFields.strComment, True, ""
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "VComment", g_udtManualFields.strVComment, True, ""
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Notes", g_udtManualFields.strNotes, True, ""
        End With
    Else
        With udtFields_out
            g_udtManualFields.blnTDE = .blnTDE
            g_udtManualFields.lngControl = .lngControl
            g_udtManualFields.lngPalNTSC = .lngPalNTSC
            g_udtManualFields.lngPlayers = .lngPlayers
            g_udtManualFields.strComment = .strComment
            g_udtManualFields.strVComment = .strVComment
            g_udtManualFields.strNotes = .strNotes
        End With
    End If
    g_udtManualFields.blnDefaults = blnSaveAsDefaults_in
    frmManualFields.Show vbModal
    If g_udtManualFields.blnOKClicked = True Then
        GetManualFieldsFromUser = True
        With udtFields_out
            .blnTDE = g_udtManualFields.blnTDE
            .lngControl = g_udtManualFields.lngControl
            .lngPalNTSC = g_udtManualFields.lngPalNTSC
            .lngPlayers = g_udtManualFields.lngPlayers
            .strComment = g_udtManualFields.strComment
            .strVComment = g_udtManualFields.strVComment
            .strNotes = g_udtManualFields.strNotes
        End With
        If blnSaveAsDefaults_in = True Then
            With m_xobjRegistry
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "TDE", g_udtManualFields.blnTDE, False
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Control", g_udtManualFields.lngControl, False
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "PalNTSC", g_udtManualFields.lngPalNTSC, False
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Players", g_udtManualFields.lngPlayers, False
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Comment", g_udtManualFields.strComment, True
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "VComment", g_udtManualFields.strVComment, True
                .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & strGBName_in & "\GEMUSFields", "Notes", g_udtManualFields.strNotes, True
            End With
        End If
    Else
        GetManualFieldsFromUser = False
    End If

End Function
Public Sub Init(strGBName_in)
    
    Dim strFileTypes As String

    m_strEmuFile = m_strAppPathWithBackSlash & strGBName_in & "\Emulators.ini"
    m_strGBName = strGBName_in
    
    m_strPathFile = m_strAppPathWithBackSlash & m_strGBName & "\Paths.ini"
    Call LoadPaths

    'load the script data
    Call LoadScriptData

    'load the file types
    strFileTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", "", m_strEmuFile)
    Call SetEmuFileTypes(strFileTypes)

    'read in the emulators
    Call ReadInEmulators

    'read in language specific text
    m_strLang_EndOfLineExpected = GetINIMessage("Emulator Script", "Error_EndOfLineExpected", "End of line expected!")
    m_strLang_ScriptError = GetINIMessage("Emulator Script", "GEMUSError", "GEMUS Error")
    m_strLang_Line = GetINIMessage("Emulator Script", "Line", "Line")
    m_strLang_UnknownCommand = GetINIMessage("Emulator Script", "Error_UnknownCommand", "Unknown function or statement!")
    m_strLang_MissingParams = GetINIMessage("Emulator Script", "Error_MissingParams", "Too few function parameters!")
    m_strLang_TooManyParams = GetINIMessage("Emulator Script", "Error_TooManyParams", "Too many function parameters!")
    m_strLang_Expected = GetINIMessage("Emulator Script", "Error_Expected", "Expected")
    m_strLang_GameNotSupported = GetINIMessage("Emulator Script", "Error_GameNotSupported", "This game type is not supported by this emulator!")
    m_strLang_InvalidFunctionParam = GetINIMessage("Emulator Script", "Error_InvalidFunctionParam", "Invalid function parameter")
    m_strLang_Number = GetINIMessage("Emulator Script", "Error_Number", "Number")
    m_strLang_ExpressionReqd = GetINIMessage("Emulator Script", "Error_ExpressionRequired", "Expression Required!")
    m_strLang_UnknownField = GetINIMessage("Emulator Script", "Error_UnknownField", "Unknown comparison field!")
    m_strLang_InvalidOp = GetINIMessage("Emulator Script", "Error_InvalidOp", "Missing or invalid comparison operator!")
    m_strLang_InvalidValue = GetINIMessage("Emulator Script", "Error_InvalidValue", "Missing or invalid comparison value!")
    m_strLang_Without = GetINIMessage("Emulator Script", "Error_Without", "without")
    m_strLang_TooMany = GetINIMessage("Emulator Script", "Error_TooMany", "Too many")
    m_strLang_Statements = GetINIMessage("Emulator Script", "Error_Statements", "statements")

End Sub
Private Function ConvertStringToNumber(strNumber_in As String) As Long

    Dim lngCharIdx As Long
    
    If Len(strNumber_in) > 4 Then
        'too big
        ConvertStringToNumber = -1
    Else
        'check for invalid chars
        For lngCharIdx = 1 To Len(strNumber_in)
            If InStr(mc_strNumbersAllowed, Mid$(strNumber_in, lngCharIdx, 1)) = 0 Then
                ConvertStringToNumber = -1
                Exit Function
            End If
        Next lngCharIdx
        'not too big and just digits
        ConvertStringToNumber = Int(Val(strNumber_in))
    End If

End Function
Private Function EvaluateContainsText(strHaystack As String, strNeedle As String) As Boolean

    '-----------------------------------------------------------------
    'Evaluates the line for CONTAINS(TEXT[||TEXT...]) comparison lines
    '
    'Inputs:    strHaystack:    String to search in (field)
    '           strNeedle:      String to search for (values in line)
    '
    'Returns:   True if any value in strNeedle is in strHaystack
    '           False otherwise
    '
    'Assumes:   strNeedle <> ""
    '-----------------------------------------------------------------

    Dim lngArgIdx As Long       'loop index
    Dim strArgs() As String     'array for split()
    Dim strUpper As String      'for comparison match
    Dim lngArgLen As Long       'length of argument

    strUpper = UCase$(strHaystack)
    strArgs() = Split(strNeedle, "||")

    For lngArgIdx = 0 To UBound(strArgs)
        strArgs(lngArgIdx) = UCase$(strArgs(lngArgIdx))
        lngArgLen = Len(strArgs(lngArgIdx))
        If strArgs(lngArgIdx) <> "" Then
            If (strArgs(lngArgIdx) = "*") Or (strArgs(lngArgIdx) = "**") Then
                'field is either empty or has characters
                EvaluateContainsText = (strHaystack <> "")
                Exit Function
            ElseIf InStr(strArgs(lngArgIdx), "*") <> 0 Then
                'field is partially matching (left/right or instr) - find out which
                If (Left$(strArgs(lngArgIdx), 1) = "*") And (Right$(strArgs(lngArgIdx), 1) = "*") Then
                    'use instr() to test for substring arg in field
                    If InStr(strUpper, Mid$(strArgs(lngArgIdx), 2, Len(strArgs(lngArgIdx)) - 2)) <> 0 Then
                        EvaluateContainsText = True
                        Exit Function
                    End If
                ElseIf Left$(strArgs(lngArgIdx), 1) = "*" Then
                    'use right$() to test last x chars in arg match field
                    If Right$(strUpper, lngArgLen - 1) = Right$(strArgs(lngArgIdx), lngArgLen - 1) Then
                        EvaluateContainsText = True
                        Exit Function
                    End If
                ElseIf Left$(strArgs(lngArgIdx), 1) = "*" Then
                    'use left$() to test last x chars in arg match field
                    If Left$(strUpper, lngArgLen - 1) = Left$(strArgs(lngArgIdx), lngArgLen - 1) Then
                        EvaluateContainsText = True
                        Exit Function
                    End If
                Else
                    'no stars at edges so must test for exact match
                    If strUpper = strArgs(lngArgIdx) Then
                        EvaluateContainsText = True
                        Exit Function
                    End If
                End If
            Else
                'field exactly matches argument?
                If strUpper = strArgs(lngArgIdx) Then
                    EvaluateContainsText = True
                    Exit Function
                End If
            End If
        End If
    Next lngArgIdx

    'none of the arguments matched
    EvaluateContainsText = False

End Function

Private Function EvaluateLine(strLine_in As String) As Boolean

    Dim strLineTest As String
    Dim strField As String
    Dim strOp As String
    Dim strData As String
    Dim strError As String
    Dim lngStatementLen As Long
    Dim lngIdx As Long
    Dim lngKeyIdx As Long

    'validate the line and get field, operator and data
    strLineTest = strLine_in
    If UCase$(Left$(strLineTest, 3)) = "IF " Then
        lngStatementLen = 3
    ElseIf UCase$(Left$(strLineTest, 7)) = "ELSEIF " Then
        lngStatementLen = 7
    ElseIf UCase$(Left$(strLineTest, 8)) = "ELSE IF " Then
        lngStatementLen = 8
    End If

    strError = ValidateComparisonLine(lngStatementLen, strLineTest, strField, strOp, strData)
    If strError = "" Then
        'find out which field
        If Left$(UCase$(strField), 4) = "KEY_" Then
            'notes text key=value comparison
            If m_lngTotalNotesKeys = 0 Then
                EvaluateLine = False
            Else
                strField = Trim$(strField)
                If Len(strField) < 5 Then
                    EvaluateLine = False
                Else
                    strField = Right$(strField, Len(strField) - 4)
                    strField = UCase$(strField)
                    lngKeyIdx = -1
                    For lngIdx = 0 To (m_lngTotalNotesKeys - 1)
                        If strField = UCase$(ma_strNotesKeys(lngIdx).strKey) Then
                            lngKeyIdx = lngIdx
                            Exit For
                        End If
                    Next lngIdx
                End If
                If lngKeyIdx = -1 Then
                    EvaluateLine = False
                Else
                    EvaluateLine = EvaluateContainsText(ma_strNotesKeys(lngKeyIdx).strValue, strData)
                    If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
                End If
            End If
        ElseIf Left$(strField, 1) = "%" Then
            'CONTAINS function on a script environment variable
            Call ReplaceEnvironmentVars(strField, False, False)
            EvaluateLine = EvaluateContainsText(strField, strData)
            If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
        Else
            Select Case strField
                Case "QUESTION"
                    Call ReplaceEnvironmentVars(strOp, False, False)
                    If MsgBox(strOp, vbQuestion Or vbYesNo, m_strAppName) = vbYes Then
                        EvaluateLine = (strData = "YES")
                    Else
                        EvaluateLine = (strData = "NO")
                    End If
                Case "FILESIZE"
                    'not implemented
                Case "NumPlayers"
                    If strOp = "=" Then
                        EvaluateLine = (m_lngFldNumPlayersTo = Val(strData))
                    ElseIf strOp = "<" Then
                        If m_lngFldNumPlayersTo <> gc_lngPLAYERSUNKNOWN Then
                            EvaluateLine = (m_lngFldNumPlayersTo < Val(strData))
                        Else
                            EvaluateLine = False
                        End If
                    ElseIf strOp = ">" Then
                        If m_lngFldNumPlayersFrom = gc_lngPLAYERSUNKNOWN Then
                            If m_lngFldNumPlayersTo = gc_lngPLAYERSUNKNOWN Then
                                'From AND To is unknown
                                EvaluateLine = False
                            Else
                                'only From is unknown
                                EvaluateLine = (m_lngFldNumPlayersTo > Val(strData))
                            End If
                        ElseIf m_lngFldNumPlayersTo = gc_lngPLAYERSUNKNOWN Then
                            'only To is unknown
                            EvaluateLine = (m_lngFldNumPlayersFrom > Val(strData))
                        Else
                            'from and to are known : use to
                            EvaluateLine = (m_lngFldNumPlayersTo > Val(strData))
                        End If
                    Else
                        EvaluateLine = (m_lngFldNumPlayersTo <> Val(strData))
                    End If
    
                Case "ImageIndex"
                    If strOp = "=" Then
                        EvaluateLine = (m_lngImageIndex = Val(strData))
                    ElseIf strOp = "<" Then
                        EvaluateLine = (m_lngImageIndex < Val(strData))
                    ElseIf strOp = ">" Then
                        EvaluateLine = (m_lngImageIndex > Val(strData))
                    Else
                        EvaluateLine = (m_lngImageIndex <> Val(strData))
                    End If
    
                Case "ImageName"
                    If m_strImageName <> "" Then
                        EvaluateLine = EvaluateContainsText(m_strImageName, strData)
                        If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
                    Else
                        EvaluateLine = False
                    End If
    
                Case "Control"
                    If strData = "JoyPort2" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 0)
                        Else
                            EvaluateLine = (m_lngFldControl <> 0)
                        End If
                    ElseIf strData = "JoyPort1" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 1)
                        Else
                            EvaluateLine = (m_lngFldControl <> 1)
                        End If
                    ElseIf strData = "Keyboard" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 2)
                        Else
                            EvaluateLine = (m_lngFldControl <> 2)
                        End If
                    ElseIf strData = "PaddlePort2" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 3)
                        Else
                            EvaluateLine = (m_lngFldControl <> 3)
                        End If
                    ElseIf strData = "PaddlePort1" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 4)
                        Else
                            EvaluateLine = (m_lngFldControl <> 4)
                        End If
                    ElseIf strData = "Mouse" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 5)
                        Else
                            EvaluateLine = (m_lngFldControl <> 5)
                        End If
                    ElseIf strData = "LightPen" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 6)
                        Else
                            EvaluateLine = (m_lngFldControl <> 6)
                        End If
                    ElseIf strData = "KoalaPad" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 7)
                        Else
                            EvaluateLine = (m_lngFldControl <> 7)
                        End If
                    ElseIf strData = "LightGun" Then
                        If strOp = "=" Then
                            EvaluateLine = (m_lngFldControl = 8)
                        Else
                            EvaluateLine = (m_lngFldControl <> 8)
                        End If
                    Else
                        If strOp = "=" Then
                            EvaluateLine = False
                        Else
                            EvaluateLine = True
                        End If
                    End If
                
                Case "GameType"
                    EvaluateLine = EvaluateContainsText(m_strFldGameType, strData)
                    If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
                
                Case "GameComment"
                    If m_strFldGameComment <> "" Then
                        EvaluateLine = EvaluateContainsText(m_strFldGameComment, strData)
                        If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
                    Else
                        EvaluateLine = False
                    End If
                
                Case "VersionComment"
                    If m_strFldVersionComment <> "" Then
                        EvaluateLine = EvaluateContainsText(m_strFldVersionComment, strData)
                        If strOp = "EXCLUDES" Then EvaluateLine = Not EvaluateLine
                    Else
                        EvaluateLine = False
                    End If
                
                Case "PalNTSC"
                    If m_lngFldPALNTSC = 0 Then
                        If strOp = "=" Then
                            EvaluateLine = (UCase$(strData) = "PAL")
                        Else
                            EvaluateLine = (UCase$(strData) <> "PAL")
                        End If
                    ElseIf m_lngFldPALNTSC = 1 Then
                        If strOp = "=" Then
                            EvaluateLine = (UCase$(strData) = "BOTH")
                        Else
                            EvaluateLine = (UCase$(strData) <> "BOTH")
                        End If
                    ElseIf m_lngFldPALNTSC = 2 Then
                        If strOp = "=" Then
                            EvaluateLine = (UCase$(strData) = "NTSC")
                        Else
                            EvaluateLine = (UCase$(strData) <> "NTSC")
                        End If
                    Else
                        If strOp = "=" Then
                            EvaluateLine = False
                        Else
                            EvaluateLine = True
                        End If
                    End If
                
                Case "TrueDriveEmu"
                    EvaluateLine = ((UCase$(strData) = "YES") And (m_blnFldTrueDriveEmu = True)) Or ((UCase$(strData) = "NO") And (m_blnFldTrueDriveEmu = False))
    
                Case "NumGameFiles"
                    If strOp = "=" Then
                        EvaluateLine = (m_lngFldNumGameFiles = Val(strData))
                    ElseIf strOp = "<" Then
                        EvaluateLine = (m_lngFldNumGameFiles < Val(strData))
                    ElseIf strOp = ">" Then
                        EvaluateLine = (m_lngFldNumGameFiles > Val(strData))
                    Else
                        EvaluateLine = (m_lngFldNumGameFiles <> Val(strData))
                    End If
                
                Case Else
                    'should NEVER happen as we have pre-checked the script by now
                    EvaluateLine = False
    
            End Select
        End If
    Else
        'should NEVER happen as we have pre-checked the script by now
        EvaluateLine = False
    End If

End Function
Private Function Process_SHOW_KV() As String

    '-------------------------------------------------
    'processes the SHOW_KV() function
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim lngIdx As Long
    
    g_strShowKV = ""
    For lngIdx = 0 To m_lngTotalNotesKeys - 1
        g_strShowKV = g_strShowKV & ma_strNotesKeys(lngIdx).strKey & "=" & ma_strNotesKeys(lngIdx).strValue & vbCrLf
    Next lngIdx
    g_strShowKV = m_xobjFunc.RTrimWhitespace(g_strShowKV)

    'show the KVs
    frmShowKV.Show vbModal

    'return ok
    Process_SHOW_KV = ""
    
End Function
Public Function GetImageNameFromGameFile(strFile_in As String, ByVal lngImageIndex_in As Long) As String

    '-----------------------------------------------------------
    'gets the image name from the passed in file and image index
    '-----------------------------------------------------------

    Dim strExt As String
    Dim D64 As clsD64
    Dim T64 As clsT64

    strExt = UCase$(Right$(strFile_in, 4))
    
    If strExt = ".D64" Then
        'D64
        Set D64 = New clsD64
        If D64.Load(strFile_in) = "" Then
            If lngImageIndex_in < D64.Entries.Count Then
                GetImageNameFromGameFile = D64.Entries(lngImageIndex_in + 1).Filename
            Else
                GetImageNameFromGameFile = ""
            End If
        Else
            GetImageNameFromGameFile = ""
        End If
        Set D64 = Nothing
    
    ElseIf strExt = ".T64" Then
        'T64
        Set T64 = New clsT64
        If T64.Load(strFile_in) = "" Then
            If lngImageIndex_in < T64.Entries.Count Then
                GetImageNameFromGameFile = T64.Entries(lngImageIndex_in + 1).C64Filename
            Else
                GetImageNameFromGameFile = ""
            End If
        Else
            GetImageNameFromGameFile = ""
        End If
        Set T64 = Nothing
    
    Else
        'IMAGE NAME DOESN'T MATTER
        GetImageNameFromGameFile = ""
    End If

End Function
Private Function GetScriptFunctionArgs(ByVal lngLenFuncName As Long, strLine_in As String, strArgs_out() As String, strError_out As String, ByVal lngReqdArgs_in As Long) As Long

    '-----------------------------------------------------------
    'Parses out the function arguments from the passed
    'in function line
    '
    'Inputs:    strLine_in:             The function line string
    '           lngReqdArgs_in:         Required number of args
    '           lngLenFuncName:         length of name of function
    '
    'Outputs:   strArgs_out():          Array of function params
    '           strError_out:           Error message if any
    '
    'Returns:   The number of params in the string
    '-----------------------------------------------------------

    Dim strArgs As String           'test string for args

    'strip function name and trim
    strArgs = Trim$(Right$(strLine_in, Len(strLine_in) - lngLenFuncName))
    'make sure first char is "("
    If Left$(strArgs, 1) <> "(" Then
        'no opening bracket
        GetScriptFunctionArgs = 0
        strError_out = m_strLang_Expected & ": ""("""
    Else
        'make sure last char is ")"
        If Right$(strArgs, 1) <> ")" Then
            'no closing bracket
            strError_out = m_strLang_Expected & ": "")"""
        Else
            'NO MORE ERRORS!
            strError_out = ""
            'test if any args
            If Len(strArgs) = 2 Then
                'no args
                GetScriptFunctionArgs = 0
            Else
                'strip brackets
                strArgs = Mid$(strArgs, 2, Len(strArgs) - 2)
                'split into params
                strArgs_out = Split(strArgs, "||")
                GetScriptFunctionArgs = UBound(strArgs_out) + 1
            End If
            'check if what is required
            If GetScriptFunctionArgs > lngReqdArgs_in Then
                strError_out = m_strLang_TooManyParams & vbCrLf & m_strLang_Expected & ": " & lngReqdArgs_in
            ElseIf GetScriptFunctionArgs < lngReqdArgs_in Then
                strError_out = m_strLang_MissingParams & vbCrLf & m_strLang_Expected & ": " & lngReqdArgs_in
            End If
        End If
    End If
    

End Function
Public Function IsFileTypeRecognised(strFT_in As String) As Boolean

    Dim lngFTIdx As Long        'loop index
    Dim strUpper As String      'upper case param

    strUpper = UCase(strFT_in)
    For lngFTIdx = 0 To (m_lngTotalEmulatorFileTypes - 1)
        If strUpper = ma_strEmulatorFileTypes(lngFTIdx) Then
            'recognised
            IsFileTypeRecognised = True
            Exit Function
        End If
    Next lngFTIdx
    'not recognised
    IsFileTypeRecognised = False

End Function

Public Function ManageEmulators() As Boolean

    ManageEmulators = False
    g_strManageEmuGBName = m_strGBName
    g_strManageEmuFile = m_strEmuFile
    frmManageEmus.Show vbModal
    If g_blnEmusManaged = True Then
        're-read in the emu details
        Call ReadInEmulators
        ManageEmulators = True
    End If

End Function

Private Function NextNonWhiteSpaceCharIs(strString_in As String, lngStartPos_in As Long, strChar_in As String) As Boolean

    Dim lngLen As Long
    Dim lngIdx As Long
    Dim strChar As String

    lngLen = Len(strString_in)
    For lngIdx = lngStartPos_in To lngLen
        strChar = Mid$(strString_in, lngIdx, 1)
        If (strChar <> " ") And (strChar <> vbTab) Then
            If Mid$(strString_in, lngIdx, Len(strChar_in)) = strChar_in Then
                NextNonWhiteSpaceCharIs = True
            Else
                NextNonWhiteSpaceCharIs = False
            End If
            Exit Function
        End If
    Next lngIdx
    NextNonWhiteSpaceCharIs = False

End Function

Private Function Process_RUN_EMULATOR(ByVal blnDOSClose_in As Boolean) As String

    Dim strExeLine As String
    Dim strPath As String
    Dim strFile As String
    Dim strCommandLine As String
    Dim dblProcessID As Double
    Dim lngHandle As Long
    Dim lngExitCode As Long
    
    On Error GoTo Err_Handler:

    Call PushLogFileTrace("clsGemus.Process_RUN_EMULATOR")
    
    'get exe line
    strExeLine = m_xobjFunc.Long2Short(ma_udtEmulators(m_lngEmuIdx).strShortPathAndFilename)

    'check it exists
    If m_xobjFunc.FileExists(strExeLine, False, "") = True Then
        'change directory to it!
        Call m_xobjFunc.GetFileAndPath(strExeLine, strPath, strFile)
        Call ChDrive(Left$(strPath, 3))
        Call ChDir(strPath)
        Call LogStatus("Changed directory to [" & strPath & "]")
        'attempt to run it!
        If blnDOSClose_in = True Then
            strCommandLine = m_strCommandCom & strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
        Else
            strCommandLine = strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
        End If
        Call LogStatus("Running emulator with command line [" & strCommandLine & "]")
        dblProcessID = Shell(strCommandLine, vbNormalFocus)
        'repack?
        If m_blnAttemptRepack = True And m_blnSilentMode = False Then
            Call LogStatus("Started to wait for emulator to finish (for repacking) at " & Format(Now(), "HH:mm:ss"))
            Call RezipWait(m_lngEmuIdx, dblProcessID, 0)
            Call LogStatus("Finished waiting for emulator to finish at " & Format(Now(), "HH:mm:ss"))
        ElseIf m_blnAttemptRepack = True Or m_blnWaitIfNotRezipping = True Then
            'wait until the process finished
            Call LogStatus("Started to wait for emulator to finish at " & Format(Now(), "HH:mm:ss"))
            lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dblProcessID)
            Do
                Call GetExitCodeProcess(lngHandle, lngExitCode)
                DoEvents
                If lngExitCode <> STATUS_PENDING Then
                    Exit Do
                End If
            Loop
            Call CloseHandle(lngHandle)
            Call LogStatus("Finished waiting for emulator to finish at " & Format(Now(), "HH:mm:ss"))
        End If
    Else
        'emulator not found
        Process_RUN_EMULATOR = GetINIMessage("Running Games", "FileNotFound", "The following file was not found") & ":" & vbCrLf & vbCrLf & strExeLine
    End If
    
    Call PopLogFileTrace
    Exit Function

Err_Handler:
    Process_RUN_EMULATOR = GetINIMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") & ":" & vbCrLf & vbCrLf & strExeLine
    Call PopLogFileTrace
    
End Function
Private Function Process_RUN_EMULATOR_SEND_KEYS(ByVal strSendKeys_in As String, ByVal strKeyDownDuration_in As String, ByVal blnDOSClose_in As Boolean) As String

    Dim strExeLine As String
    Dim strPath As String
    Dim strFile As String
    Dim strCommandLine As String
    Dim dblProcessID As Double
    Dim lngHandle As Long
    Dim lngExitCode As Long
    Dim lngTotalSendBlocks As Long
    Dim a_udtSendBlocks() As SendKeysBlockType
    Dim lngIdx As Long
    Dim lngKeyDownDuration As Long
    
    On Error GoTo Err_Handler
    
    Call PushLogFileTrace("clsGemus.Process_RUN_EMULATOR_SEND_KEYS")

    'get exe line
    strExeLine = m_xobjFunc.Long2Short(ma_udtEmulators(m_lngEmuIdx).strShortPathAndFilename)

    'check it exists
    If m_xobjFunc.FileExists(strExeLine, False, "") = True Then
    
        'change directory to it!
        Call m_xobjFunc.GetFileAndPath(strExeLine, strPath, strFile)
        Call ChDrive(Left$(strPath, 3))
        Call ChDir(strPath)
        Call LogStatus("Changed directory to [" & strPath & "]")
        'attempt to run it!
        If blnDOSClose_in = True Then
            strCommandLine = m_strCommandCom & strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
        Else
            strCommandLine = strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
        End If
        Call LogStatus("Running emulator with command line [" & strCommandLine & "]")
        dblProcessID = Shell(strCommandLine, vbNormalFocus)
        
        'send keys
        Call LogStatus("Started to send keystrokes at " & Format(Now(), "HH:mm:ss"))
        lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dblProcessID)
        lngTotalSendBlocks = ParseSendKeys(strSendKeys_in, a_udtSendBlocks)
        Call ReplaceEnvironmentVars(strKeyDownDuration_in, False, False)
        lngKeyDownDuration = Val(strKeyDownDuration_in)
        If lngKeyDownDuration < 0 Then lngKeyDownDuration = 0
        For lngIdx = 0 To lngTotalSendBlocks - 1
            With a_udtSendBlocks(lngIdx)
                If .lngSecsToWait = 0 Then
                    Call GetExitCodeProcess(lngHandle, lngExitCode)
                    If lngExitCode <> STATUS_PENDING Then
                        Call LogStatus("Aborted sending any remaining keystrokes as emulator was closed")
                        Exit For
                    End If
                    Call VbSendKeys(.strKeysToSend, lngKeyDownDuration)
                    If g_strVBSendKeysError <> "" Then
                        Call CloseHandle(lngHandle)
                        Process_RUN_EMULATOR_SEND_KEYS = "Failed Sending Keys """ & .strKeysToSend & """ to " & strExeLine & vbCrLf & vbCrLf & "Please check the keys you are sending are in the valid format."
                        Call PopLogFileTrace
                        Exit Function
                    End If
                ElseIf .lngSecsToWait < 0 Then
                    'treat as new keypress duration ms
                    lngKeyDownDuration = Abs(.lngSecsToWait)
                ElseIf .lngSecsToWait > 0 Then
                    Call WaitSecondsWhileProcessOpen(.lngSecsToWait, lngHandle)
                End If
            End With
        Next lngIdx
        Call CloseHandle(lngHandle)
        Call LogStatus("Finished sending keystrokes at " & Format(Now(), "HH:mm:ss"))

        'rezip?
        If m_blnAttemptRepack = True And m_blnSilentMode = False Then
            Call LogStatus("Started to wait for emulator to finish (for repacking) at " & Format(Now(), "HH:mm:ss"))
            Call RezipWait(m_lngEmuIdx, dblProcessID, 0)
            Call LogStatus("Finished waiting for emulator to finish at " & Format(Now(), "HH:mm:ss"))
        ElseIf m_blnAttemptRepack = True Or m_blnWaitIfNotRezipping = True Then
            'wait until the process finished
            Call LogStatus("Started to wait for emulator to finish at " & Format(Now(), "HH:mm:ss"))
            lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dblProcessID)
            Do
                Call GetExitCodeProcess(lngHandle, lngExitCode)
                DoEvents
                If lngExitCode <> STATUS_PENDING Then
                    Exit Do
                End If
            Loop
            Call CloseHandle(lngHandle)
            Call LogStatus("Finished waiting for emulator to finish at " & Format(Now(), "HH:mm:ss"))
        End If
    Else
        'emulator not found
        Process_RUN_EMULATOR_SEND_KEYS = GetINIMessage("Running Games", "FileNotFound", "The following file was not found") & ":" & vbCrLf & vbCrLf & strExeLine
    End If
    
    Call PopLogFileTrace
    Exit Function

Err_Handler:
    Process_RUN_EMULATOR_SEND_KEYS = GetINIMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") & ":" & vbCrLf & vbCrLf & strExeLine
    Call PopLogFileTrace
    
End Function
Private Function VBSendKeysEscape(ByVal strText_in As String, ByVal blnDoEscape_in As Boolean) As String

    VBSendKeysEscape = strText_in
    If Not blnDoEscape_in Then Exit Function

    ' {
    VBSendKeysEscape = Replace(VBSendKeysEscape, "{", "{{}")
    
    ' }
    VBSendKeysEscape = Replace(VBSendKeysEscape, "{{}", "~~~~~TEMP_REPLACEMENT~~~~~")
    VBSendKeysEscape = Replace(VBSendKeysEscape, "}", "{}}")
    VBSendKeysEscape = Replace(VBSendKeysEscape, "~~~~~TEMP_REPLACEMENT~~~~~", "{{}")

    ' +
    VBSendKeysEscape = Replace(VBSendKeysEscape, "+", "{+}")

    ' ^
    VBSendKeysEscape = Replace(VBSendKeysEscape, "^", "{^}")

    ' %
    VBSendKeysEscape = Replace(VBSendKeysEscape, "%", "{%}")

    ' ~
    VBSendKeysEscape = Replace(VBSendKeysEscape, "~", "{~}")

End Function

Private Sub WaitSecondsWhileProcessOpen(ByVal lngNumSeconds_in As Long, ByVal lngHandle_in As Long)

    Dim lngFinishTime As Long
    Dim lngExitCode As Long

    lngFinishTime = GetTickCount + (lngNumSeconds_in * 1000)
    While GetTickCount() < lngFinishTime
        Call GetExitCodeProcess(lngHandle_in, lngExitCode)
        If lngExitCode <> STATUS_PENDING Then
            Exit Sub
        End If
        DoEvents
    Wend

End Sub
Private Function Process_RUN_GAMEFILE(ByVal blnDOSClose_in As Boolean) As String

    Dim strExeLine As String
    Dim strPath As String
    Dim strFile As String
    Dim strCommandLine As String
    Dim strExt As String
    Dim lngHandle As Long
    Dim dblProcessID As Double

    On Error GoTo Err_Handler:

    Call PushLogFileTrace("clsGemus.Process_RUN_GAMEFILE")

    'get exe line
    strExeLine = m_xobjFunc.Long2Short(m_strEnvGFFull)

    'check it exists
    If m_xobjFunc.FileExists(strExeLine, False, "") = True Then
        'change directory to it!
        Call m_xobjFunc.GetFileAndPath(strExeLine, strPath, strFile)
        Call ChDrive(Left$(strPath, 3))
        Call ChDir(strPath)
        Call LogStatus("Changed directory to [" & strPath & "]")
        'attempt to run it!
        strExt = UCase$(Right$(strFile, 4))
        If strExt = ".EXE" Or strExt = ".COM" Or strExt = ".BAT" Then
            'use shell
            If blnDOSClose_in = True Then
                strCommandLine = m_strCommandCom & strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
            Else
                strCommandLine = strFile & IIf((m_strCommandLine = ""), "", " " & m_strCommandLine)
            End If
            Call LogStatus("Running gamefile with command line [" & strCommandLine & "]")
            dblProcessID = Shell(strCommandLine, vbNormalFocus)
            lngHandle = 0
        Else
            'use shell execute and ignore DOSCLOSE and command line
            Call LogStatus("Running gamefile [" & strExeLine & "]")
            Call ShellExecuteGameFile(strExeLine, "", lngHandle)
            Call LogStatus("Started to wait for gamefile to finish at " & Format(Now(), "HH:mm:ss"))
            Call RezipWait(m_lngEmuIdx, 0, lngHandle)
            Call LogStatus("Finished waiting for gamefile to finish at " & Format(Now(), "HH:mm:ss"))
        End If
    Else
        'game file not found
        Process_RUN_GAMEFILE = GetINIMessage("Running Games", "FileNotFound", "The following file was not found") & ":" & vbCrLf & vbCrLf & strExeLine
    End If
    
    Call PopLogFileTrace
    Exit Function

Err_Handler:
    Process_RUN_GAMEFILE = GetINIMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") & ":" & vbCrLf & vbCrLf & strExeLine
    Call PopLogFileTrace
    
End Function
Private Function Process_SHOW_MESSAGE(strMessage_in As String) As String

    '-------------------------------------------------
    'processes the SHOW_MESSAGE() function
    '
    'Inputs:    strMessage_in:  The message to display
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim strMessage As String

    'replace environment variables in the message
    strMessage = strMessage_in
    Call ReplaceEnvironmentVars(strMessage, False, False)

    'show the message
    Call MsgBox(strMessage, vbInformation, m_strAppName)

    'return ok
    Process_SHOW_MESSAGE = ""
    
End Function

Private Function Process_WAIT(strMillisecondsToWait_in As String) As String

    '-------------------------------------------------
    'processes the WAIT() function
    '
    'Inputs:    lngMillisecondsToWait:  The number of seconds to wait
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim strMillisecondsToWait As String
    Dim lngMillisecondsToWait&
    Dim lngFinishTime&

    'catch overflow errors
    On Error GoTo Err_Handler

    Call PushLogFileTrace("clsGemus.Process_WAIT")

    'replace environment variables in the argument
    strMillisecondsToWait = strMillisecondsToWait_in
    Call ReplaceEnvironmentVars(strMillisecondsToWait, False, False)

    'convert it to a number
    lngMillisecondsToWait = Val(strMillisecondsToWait)

    'log status
    Call LogStatus("Waiting for [" & lngMillisecondsToWait & "] milliseconds...")

    'wait for number of milliseconds specified
    lngFinishTime = GetTickCount + (lngMillisecondsToWait)
    While GetTickCount() < lngFinishTime
        DoEvents
    Wend

    'log status
    Call LogStatus("Finished waiting!")

Err_Handler:

    'return ok
    Call PopLogFileTrace
    Process_WAIT = ""

End Function


Private Function Process_KILL_PROCESS(strFilename_in As String) As String

    '-------------------------------------------------
    'processes the KILL_PROCESS() function
    '
    'Inputs:    strFilename_in:  filename to kill
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim strFilename As String
    Dim lngNumProcessesToKill As Long
    Dim a_lngProcessIDsToKill() As Long
    Dim lngIdx As Long
    Dim lngProcessHandle As Long

    Call PushLogFileTrace("clsGemus.Process_KILL_PROCESS")

    'replace environment variables in the argument
    strFilename = strFilename_in
    Call ReplaceEnvironmentVars(strFilename, False, False)

    'log status
    Call LogStatus("Attempting to kill processes matching [" & strFilename & "]...")

    'get the process ids of matching running processes
    lngNumProcessesToKill = GetRunningProcessIDsByFilename(strFilename, a_lngProcessIDsToKill)

    'kill'em
    If lngNumProcessesToKill > 0 Then
        For lngIdx = 0 To lngNumProcessesToKill - 1
            If a_lngProcessIDsToKill(lngIdx) > 0 Then
                lngProcessHandle = OpenProcess(PROCESS_TERMINATE, 0, a_lngProcessIDsToKill(lngIdx))
                Call TerminateProcess(lngProcessHandle, 0&)
                Call CloseHandle(lngProcessHandle)
            End If
        Next lngIdx
    End If

    'log status
    Call LogStatus("Killed X matching process(es)")

    'return ok
    Call PopLogFileTrace
    Process_KILL_PROCESS = ""

End Function


Private Function Process_ADD_CLP(strCLP_in As String) As String

    '-------------------------------------------------
    'processes the ADD_CLP() function
    '
    'Inputs:    strCLP_in:  command line text to add
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim strCLP As String

    'replace environment variables in the command line
    strCLP = strCLP_in
    Call ReplaceEnvironmentVars(strCLP, True, False)

    'add to command line
    m_strCommandLine = m_strCommandLine & strCLP

    'return ok
    Process_ADD_CLP = ""
    
    Call PushLogFileTrace("clsGemus.Process_ADD_CLP")
    Call LogStatus("Adding to command line: [" & strCLP & "]")
    Call LogStatus("Command line is now: [" & m_strCommandLine & "]")
    Call PopLogFileTrace

End Function
Private Function Process_ADD_CLP2(strCLP_in As String) As String

    '-------------------------------------------------
    'processes the ADD_CLP2() function
    '
    'Inputs:    strCLP_in:  command line text to add
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    Dim strCLP As String

    'replace environment variables in the command line
    strCLP = strCLP_in
    Call ReplaceEnvironmentVars(strCLP, False, False)

    'add to command line
    m_strCommandLine = m_strCommandLine & strCLP

    'return ok
    Process_ADD_CLP2 = ""

    Call PushLogFileTrace("clsGemus.Process_ADD_CLP2")
    Call LogStatus("Adding to command line: [" & strCLP & "]")
    Call LogStatus("Command line is now: [" & m_strCommandLine & "]")
    Call PopLogFileTrace

End Function
Private Function Process_EDIT_CLP(strPrompt_in As String) As String

    '-------------------------------------------------
    'processes the EDIT_CLP() function.  Brings up a
    'message box and allows the user to change the
    'current contents of the command line parameters
    '
    'Inputs:    strPrompt_in: Message prompt
    '
    'Returns:   None
    '-------------------------------------------------

    Dim strMessage As String

    'replace environment variables in the message
    strMessage = strPrompt_in
    Call ReplaceEnvironmentVars(strMessage, False, False)

    'show the edit window
    With g_udtInputBox
        .blnAllowEmptyReturn = True
        .blnOKClicked = False
        .lngMaxLength = 8092
        .strCancelButtonText = GetINIMessage("General Buttons", "Cancel", "&Cancel")
        .strCaption = "GEMUS"
        .strInvalidChars = ""
        .strOKButtonText = GetINIMessage("General Buttons", "OK", "&OK")
        .strPrompt = strMessage & ":"
        .strReturnedText = ""
        .strStartingText = m_strCommandLine
    End With
    frmInputBox.Show vbModal
    If g_udtInputBox.blnOKClicked = True Then
        m_strCommandLine = g_udtInputBox.strReturnedText
    End If

    'return ok
    Process_EDIT_CLP = ""

    Call PushLogFileTrace("clsGemus.Process_EDIT_CLP")
    Call LogStatus("Command line (after edit) is now: [" & m_strCommandLine & "]")
    Call PopLogFileTrace

End Function
Private Function Process_CLEAR_CLP() As String

    '-------------------------------------------------
    'processes the CLEAR_CLP() function
    '
    'Returns:   Error message or "" if ok
    '-------------------------------------------------

    'clear command line
    m_strCommandLine = ""
    
    'return ok
    Process_CLEAR_CLP = ""

    Call PushLogFileTrace("clsGemus.Process_CLEAR_CLP2")
    Call LogStatus("Command line cleared")
    Call PopLogFileTrace

End Function
Private Function Process_RUN_PROGRAM(lngFileNum_in As Long, strCLP_in As String, blnWait_in As Boolean, strFilename_in As String) As String

    '----------------------------------------------------------
    'processes the RUN_PROGRAM() function
    '
    'Inputs:    lngFileNum_in:  file number of assoc script file to run if strFilename_in is empty
    '           strCLP_in:      command line text to add
    '           blnWait_in:     wait till program finishes flag
    '           strFilename_in: filename to run if not empty
    '
    'Returns:   Error message or "" if ok
    '----------------------------------------------------------

    Dim strCLP As String
    Dim strFile As String
    Dim blnFound As Boolean
    Dim lngFileNum As Long
    Dim strFilename As String
    Dim strPath As String
    
    On Error GoTo Err_Handler
    
    Call PushLogFileTrace("clsGemus.Process_RUN_PROGRAM")
    
    'get filename
    If strFilename_in <> "" Then
        strFile = strFilename_in
        Call ReplaceEnvironmentVars(strFile, True, False)
        blnFound = True
    Else
        'check the file exists
        lngFileNum = lngFileNum_in - 1
        If lngFileNum < 0 Or lngFileNum > (ma_udtEmulators(m_lngEmuIdx).lngNumFiles - 1) Then
            blnFound = False
            'invalid file number
            If MsgBox(GetINIMessage("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function.") & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName & " - " & m_strLang_ScriptError) = vbNo Then
                Process_RUN_PROGRAM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                Call PopLogFileTrace
                Exit Function
            End If
        Else
            'check file exists
            strFile = ma_udtEmulators(m_lngEmuIdx).strFiles(lngFileNum)
            blnFound = True
        End If

    End If

    Call LogStatus("Determined program to run to be [" & strFile & "]")

    'make sure it exists
    If m_xobjFunc.FileExists(strFile, False, "") = False Then
        blnFound = False
        If MsgBox(GetINIMessage("Running Games", "FileNotFound", "The following file was not found") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
            Process_RUN_PROGRAM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
            Call PopLogFileTrace
            Exit Function
        End If
    End If

    'if here we can run the program if the file was found!
    If blnFound = True Then
    
        'replace environment variables in the command line parameters
        strCLP = strCLP_in
        Call ReplaceEnvironmentVars(strCLP, True, False)
        'make the command line string
        If m_blnUseShort = True Then
            strFile = m_xobjFunc.Long2Short(strFile)
        End If
        'run the program
        Call m_xobjFunc.GetFileAndPath(strFile, strPath, strFilename)
        Call ChDrive(Left$(strPath, 3))
        Call ChDir(strPath)
        Call LogStatus("Changed directory to [" & strPath & "]")
        Call LogStatus("Running program [" & strFile & "] with command line arguments [" & strCLP & "]")
        If m_xobjFunc.ShellExecuteFile(strFile, strCLP, blnWait_in) = False Then
            If MsgBox(GetINIMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") & ": " & strCLP & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                Process_RUN_PROGRAM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                Call PopLogFileTrace
                Exit Function
            Else
                'error but continue with script
                Process_RUN_PROGRAM = ""
            End If
        Else
            'all ok
            Process_RUN_PROGRAM = ""
        End If
    End If

    Call PopLogFileTrace
    Exit Function

Err_Handler:
    Process_RUN_PROGRAM = GetINIMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") & ":" & vbCrLf & vbCrLf & strFile
    Call PopLogFileTrace

End Function

Private Function Process_SET_INI_VALUE(lngFileNum_in As Long, strSection_in As String, strKey_in As String, strValue_in As String, strFilename_in As String) As String

    Dim strFile As String
    Dim blnFound As Boolean
    Dim lngFileNum As Long
    Dim strSection As String
    Dim strKey As String
    Dim strValue As String

    Call PushLogFileTrace("clsGemus.Process_SET_INI_VALUE")

    'check the file exists
    If strFilename_in = "" Then
        lngFileNum = lngFileNum_in - 1
        If lngFileNum < 0 Or lngFileNum > (ma_udtEmulators(m_lngEmuIdx).lngNumFiles - 1) Then
            blnFound = False
            'invalid file number
            Call LogStatus("Associated script file [" & lngFileNum_in & "] does not point to valid file")
            If m_blnScriptFileErrorDone = False Then
                m_blnScriptFileErrorDone = True
                If MsgBox(GetINIMessage("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function.") & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName & " - " & m_strLang_ScriptError) = vbNo Then
                    Process_SET_INI_VALUE = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                    Call PopLogFileTrace
                    Exit Function
                Else
                    Call LogStatus("Unable to set ini value but user chose to continue anyway.")
                End If
            End If
        Else
            'should be ok
            strFile = ma_udtEmulators(m_lngEmuIdx).strFiles(lngFileNum)
            Call LogStatus("Associated script file [" & lngFileNum_in & "] determined to be [" & strFile & "]")
            blnFound = True
        End If
    Else
        blnFound = True
        strFile = strFilename_in
        Call ReplaceEnvironmentVars(strFile, False, False)
    End If

    'if here we can add the value to the ini file
    If blnFound = True Then
        'replace environment variables in the section, key and value vars
        strSection = strSection_in
        strKey = strKey_in
        strValue = strValue_in
        If Left$(strSection, 1) = "[" And Right$(strSection, 1) = "]" Then
            If Len(strSection) > 2 Then
                strSection = Mid$(strSection, 2, Len(strSection) - 2)
            End If
        End If
        Call ReplaceEnvironmentVars(strSection, False, False)
        Call ReplaceEnvironmentVars(strKey, False, False)
        Call ReplaceEnvironmentVars(strValue, False, False)
        'attempt to write the string to the ini file
        If Trim$(strSection) = "" Or Trim$(strKey) = "" Then
            'don't bother
            Process_SET_INI_VALUE = ""
        Else
            
            Call LogStatus("Setting ini value [" & strKey & "=" & strValue & "] in section [" & strSection & "] in file [" & strFile & "]")
            If WritePrivateProfileString(strSection, strKey, strValue, strFile) = 0 Then
                If m_blnScriptFileErrorDone = False Then
                    m_blnScriptFileErrorDone = True
                    If MsgBox(GetINIMessage("Running Games", "ErrorSetINIValue", "Could not set the following INI File value") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & "[" & strSection & "]" & vbCrLf & strKey & "=" & strValue & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                        Process_SET_INI_VALUE = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                        Call PopLogFileTrace
                        Exit Function
                    Else
                        'error but continue anyway
                        Call LogStatus("Error setting ini value but user chose to continue anyway.")
                        Process_SET_INI_VALUE = ""
                    End If
                End If
            Else
                'written ok!
                Process_SET_INI_VALUE = ""
            End If
        End If
    End If
    
    Call PopLogFileTrace

End Function
Private Function Process_SET_CFG_VALUE(lngFileNum_in As Long, strKey_in As String, strValue_in As String, strFilename_in As String) As String

    Dim strFile As String
    Dim blnFound As Boolean
    Dim lngFileNum As Long
    Dim strKey As String
    Dim strValue As String

    Call PushLogFileTrace("clsGemus.Process_SET_CFG_VALUE")

    'check the file exists
    If strFilename_in = "" Then
        lngFileNum = lngFileNum_in - 1
        If lngFileNum < 0 Or lngFileNum > (ma_udtEmulators(m_lngEmuIdx).lngNumFiles - 1) Then
            blnFound = False
            'invalid file number
            Call LogStatus("Associated script file [" & lngFileNum_in & "] does not point to valid file")
            If m_blnScriptFileErrorDone = False Then
                m_blnScriptFileErrorDone = True
                If MsgBox(GetINIMessage("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function.") & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName & " - " & m_strLang_ScriptError) = vbNo Then
                    Process_SET_CFG_VALUE = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                    Call PopLogFileTrace
                    Exit Function
                Else
                    Call LogStatus("Unable to set cfg value but user chose to continue anyway.")
                End If
            End If
        Else
            'should be ok
            strFile = ma_udtEmulators(m_lngEmuIdx).strFiles(lngFileNum)
            Call LogStatus("Associated script file [" & lngFileNum_in & "] determined to be [" & strFile & "]")
            blnFound = True
        End If
    Else
        blnFound = True
        strFile = strFilename_in
        Call ReplaceEnvironmentVars(strFile, False, False)
    End If

    'is the file on a CD?
    If GetTypeOfDrive(Left$(strFile, 3)) = e_PathLocation_CD Then
        Call LogStatus("Unable to set cfg value due to readonly media (CD/DVD).")
        If m_blnScriptFileErrorDone = False Then
            m_blnScriptFileErrorDone = True
            If MsgBox(GetINIMessage("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & strKey & "=" & strValue & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                Process_SET_CFG_VALUE = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                Call PopLogFileTrace
                Exit Function
            Else
                'error but continue anyway
                Process_SET_CFG_VALUE = ""
            End If
        End If
    Else
        'if here we can add the value to the cfg file
        If blnFound = True Then
            'replace environment variables in the section, key and value vars
            strKey = strKey_in
            strValue = strValue_in
            Call ReplaceEnvironmentVars(strKey, False, False)
            Call ReplaceEnvironmentVars(strValue, False, False)
            'attempt to write the value to the cfg file
            If Trim$(strKey) = "" Then
                'don't bother
                Process_SET_CFG_VALUE = ""
            Else
                Call LogStatus("Setting cfg value [" & strKey & "=" & strValue & "] in file [" & strFile & "]")
                If WriteCFGValue(strFile, strKey, strValue) = False Then
                    If m_blnScriptFileErrorDone = False Then
                        m_blnScriptFileErrorDone = True
                        If MsgBox(GetINIMessage("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & strKey & "=" & strValue & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                            Process_SET_CFG_VALUE = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                            Call PopLogFileTrace
                            Exit Function
                        Else
                            'error but continue anyway
                            Call LogStatus("Error setting cfg value but user chose to continue anyway.")
                            Process_SET_CFG_VALUE = ""
                        End If
                    End If
                Else
                    'written ok!
                    Process_SET_CFG_VALUE = ""
                End If
            End If
        End If
    End If

    Call PopLogFileTrace

End Function

Private Function Process_SET_CFG_ITEM(lngFileNum_in As Long, strKey_in As String, strDelim_in As String, strValue_in As String, strFilename_in As String) As String

    Dim strFile As String
    Dim blnFound As Boolean
    Dim lngFileNum As Long
    Dim strKey As String
    Dim strValue As String
    Dim strDelim As String

    Call PushLogFileTrace("clsGemus.Process_SET_CFG_ITEM")

    'check the file exists
    If strFilename_in = "" Then
        lngFileNum = lngFileNum_in - 1
        If lngFileNum < 0 Or lngFileNum > (ma_udtEmulators(m_lngEmuIdx).lngNumFiles - 1) Then
            blnFound = False
            'invalid file number
            Call LogStatus("Associated script file [" & lngFileNum_in & "] does not point to valid file")
            If m_blnScriptFileErrorDone = False Then
                m_blnScriptFileErrorDone = True
                If MsgBox(GetINIMessage("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function.") & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName & " - " & m_strLang_ScriptError) = vbNo Then
                    Process_SET_CFG_ITEM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                    Call PopLogFileTrace
                    Exit Function
                Else
                    Call LogStatus("Unable to set cfg item but user chose to continue anyway.")
                End If
            End If
        Else
            'should be ok
            strFile = ma_udtEmulators(m_lngEmuIdx).strFiles(lngFileNum)
            Call LogStatus("Associated script file [" & lngFileNum_in & "] determined to be [" & strFile & "]")
            blnFound = True
        End If
    Else
        blnFound = True
        strFile = strFilename_in
        Call ReplaceEnvironmentVars(strFile, False, False)
    End If

    'is the file on a CD?
    If GetTypeOfDrive(Left$(strFile, 3)) = e_PathLocation_CD Then
        Call LogStatus("Unable to set cfg item due to readonly media (CD/DVD).")
        If m_blnScriptFileErrorDone = False Then
            m_blnScriptFileErrorDone = True
            If MsgBox(GetINIMessage("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & strKey & "=" & strValue & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                Process_SET_CFG_ITEM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                Call PopLogFileTrace
                Exit Function
            Else
                'error but continue anyway
                Process_SET_CFG_ITEM = ""
            End If
        End If
    Else
        'if here we can add the value to the cfg file
        If blnFound = True Then
            'replace environment variables in the key, delim and value vars
            strKey = strKey_in
            strDelim = strDelim_in
            strValue = strValue_in
            Call ReplaceEnvironmentVars(strKey, False, False)
            If Len(strDelim) > 0 Then
                Call ReplaceEnvironmentVars(strDelim, False, False)
            End If
            If Len(strValue) > 0 Then
                Call ReplaceEnvironmentVars(strValue, False, False)
            End If
            'attempt to write the value to the cfg file
            If Trim$(strKey) = "" Then
                'don't bother
                Process_SET_CFG_ITEM = ""
            Else
                Call LogStatus("Setting cfg item [" & strKey & strDelim & strValue & "] in file [" & strFile & "]")
                If WriteCFGValue2(strFile, strKey, strDelim, strValue) = False Then
                    If m_blnScriptFileErrorDone = False Then
                        m_blnScriptFileErrorDone = True
                        If MsgBox(GetINIMessage("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value") & ":" & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf & strKey & strDelim & strValue & vbCrLf & vbCrLf & GetINIMessage("Running Games", "Continue", "Do you want to continue running the game anyway?"), vbYesNo Or vbQuestion, m_strAppName) = vbNo Then
                            Process_SET_CFG_ITEM = GetINIMessage("Running Games", "ScriptAborted", "Script aborted!")
                            Call PopLogFileTrace
                            Exit Function
                        Else
                            'error but continue anyway
                            Call LogStatus("Error setting cfg item but user chose to continue anyway.")
                            Process_SET_CFG_ITEM = ""
                        End If
                    End If
                Else
                    'written ok!
                    Process_SET_CFG_ITEM = ""
                End If
            End If
        End If
    End If

    Call PopLogFileTrace

End Function

Private Function ProcessScriptLine(strLine_in As String) As String

    '-------------------------------------------------------------------
    'finds out which function to execute, validates it, then executes it
    '
    'Inputs:    strLine_in:     The script line to process
    '
    'Returns:   Error message or "" if processed ok
    '-------------------------------------------------------------------

    Dim strError As String      'for validation
    Dim strLine As String       'for validation
    Dim strUpper As String      'for testing which function
    Dim strCLP As String        'for ADD_CLP()
    Dim lngFileNum As Long      'for SET_INI/CFG_VALUE() and RUN_PROGRAM()
    Dim strSection As String    'for SET_INI_VALUE()
    Dim strKey As String        'for SET_INI/CFG_VALUE/CFG_ITEM()
    Dim strValue As String      'for SET_INI/CFG_VALUE/CFG_ITEM()
    Dim strDelim As String      'for SET_CFG_ITEM()
    Dim blnWait As Boolean      'for RUN_PROGRAM()
    Dim strFilename As String   'for RUN_PROGRAM(), SET_CFG_ITEM(), SET_INI_VALUE(), SET_CFG_VALUE(), KILL_PROCESS()
    Dim strMessage As String    'for SHOW_MESSAGE()/EDIT_CLP()
    Dim blnDOSClose As Boolean  'for RUN_EMULATOR()
    Dim strKeyDownDuration As String 'for RUN_EMULATOR_SEND_KEYS()
    Dim strSendKeys As String   'for RUN_EMULATOR_SEND_KEYS()
    Dim strMillisecondsToWait As String   'for WAIT()

    'validate the function line
    strLine = strLine_in
    strUpper = UCase$(strLine)
    If Left$(strUpper, 8) = "ADD_CLP2" Then
        
        'ADD_CLP2()
        
        strError = Validate_ADD_CLP2(strLine, strCLP)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_ADD_CLP2(strCLP)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If
        
    ElseIf Left$(strUpper, 7) = "ADD_CLP" Then
        
        'ADD_CLP()
        
        strError = Validate_ADD_CLP(strLine, strCLP)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_ADD_CLP(strCLP)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 8) = "EDIT_CLP" Then
        
        'EDIT_CLP()
        
        strError = Validate_EDIT_CLP(strLine, strMessage)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_EDIT_CLP(strMessage)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 9) = "CLEAR_CLP" Then
        
        'CLEAR_CLP()
        
        strError = Validate_CLEAR_CLP(strLine)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_CLEAR_CLP()
        Else
            'error validating function
            ProcessScriptLine = strError
        End If
    
    ElseIf Left$(strUpper, 7) = "SHOW_KV" Then
        
        'SHOW_KV()
        
        strError = Validate_SHOW_KV(strLine)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_SHOW_KV()
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 4) = "WAIT" Then
    
        'WAIT()
        
        strError = Validate_WAIT(strLine, strMillisecondsToWait)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_WAIT(strMillisecondsToWait)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 12) = "KILL_PROCESS" Then
    
        'KILL_PROCESS()
        
        strError = Validate_KILL_PROCESS(strLine, strFilename)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_KILL_PROCESS(strFilename)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 13) = "SET_INI_VALUE" Then
        
        'SET_INI_VALUE
        
        strError = Validate_SET_INI_VALUE(strLine, lngFileNum, strSection, strKey, strValue, strFilename)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_SET_INI_VALUE(lngFileNum, strSection, strKey, strValue, strFilename)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 12) = "SET_CFG_ITEM" Then
        
        'SET_CFG_ITEM
        
        strError = Validate_SET_CFG_ITEM(strLine, lngFileNum, strKey, strDelim, strValue, strFilename)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_SET_CFG_ITEM(lngFileNum, strKey, strDelim, strValue, strFilename)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 13) = "SET_CFG_VALUE" Then
        
        'SET_CFG_VALUE
        
        strError = Validate_SET_CFG_VALUE(strLine, lngFileNum, strKey, strValue, strFilename)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_SET_CFG_VALUE(lngFileNum, strKey, strValue, strFilename)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If
        
    ElseIf Left$(strUpper, 11) = "RUN_PROGRAM" Then
        
        'RUN_PROGRAM()
        
        strError = Validate_RUN_PROGRAM(strLine, lngFileNum, strCLP, blnWait, strFilename)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_RUN_PROGRAM(lngFileNum, strCLP, blnWait, strFilename)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 22) = "RUN_EMULATOR_SEND_KEYS" Then
        
        'RUN_EMULATOR_SEND_KEYS()

        strError = Validate_RUN_EMULATOR_SEND_KEYS(strLine, strSendKeys, strKeyDownDuration, blnDOSClose)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_RUN_EMULATOR_SEND_KEYS(strSendKeys, strKeyDownDuration, blnDOSClose)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 12) = "RUN_EMULATOR" Then
        
        'RUN_EMULATOR()
        
        strError = Validate_RUN_EMULATOR(strLine, blnDOSClose)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_RUN_EMULATOR(blnDOSClose)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If

    ElseIf Left$(strUpper, 12) = "RUN_GAMEFILE" Then
        
        'RUN_GAMEFILE()
        
        strError = Validate_RUN_GAMEFILE(strLine, blnDOSClose)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_RUN_GAMEFILE(blnDOSClose)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If
        
    ElseIf Left$(strUpper, 12) = "SHOW_MESSAGE" Then
        
        'SHOW_MESSAGE()
        
        strError = Validate_SHOW_MESSAGE(strLine, strMessage)
        If strError = "" Then
            'process the function
            ProcessScriptLine = Process_SHOW_MESSAGE(strMessage)
        Else
            'error validating function
            ProcessScriptLine = strError
        End If
        
    Else
        'not a recognised function
        ProcessScriptLine = m_strLang_UnknownCommand
    End If

End Function
Private Sub ReadInEmulators()

    Dim strSections() As String
    Dim strNames() As String
    Dim strExes() As String
    Dim strScript() As String
    Dim strFiles() As String
    Dim blnUseShort() As Boolean
    Dim blnUseDefaultTypes() As Boolean
    Dim lngWaits() As Long
    Dim strTypes() As String
    Dim lngNumFiles As Long
    Dim lngPathIdx As Long
    Dim lngFileIdx As Long
    Dim lngEmuIdx As Long
    Dim strSection As String
    Dim strFile As String

    With m_xobjFunc
        'get list of emulator names (sections)
        m_lngTotalEmulators = 0
        ReDim strSections(0)
        lngEmuIdx = 1
        Do
            strSection = .INI_GetSetting("Emulators", Trim$(lngEmuIdx), "", m_strEmuFile)
            If strSection <> "" Then
                ReDim Preserve strSections(lngEmuIdx - 1)
                strSections(lngEmuIdx - 1) = strSection
                lngEmuIdx = lngEmuIdx + 1
            Else
                Exit Do
            End If
        Loop
        'read in their details
        If lngEmuIdx > 1 Then
            m_lngTotalEmulators = lngEmuIdx - 1
            ReDim ma_udtEmulators(m_lngTotalEmulators - 1)
            ReDim strNames(m_lngTotalEmulators - 1)
            ReDim strExes(m_lngTotalEmulators - 1)
            ReDim blnUseShort(m_lngTotalEmulators - 1)
            ReDim strScript(m_lngTotalEmulators - 1)
            ReDim strTypes(m_lngTotalEmulators - 1)
            ReDim blnUseDefaultTypes(m_lngTotalEmulators - 1)
            ReDim lngWaits(m_lngTotalEmulators - 1)
            lngPathIdx = 0
            For lngEmuIdx = 0 To (m_lngTotalEmulators - 1)
                blnUseShort(lngPathIdx) = (.INI_GetSetting(strSections(lngEmuIdx), "UseShortFilenames", "0", m_strEmuFile) = "1")
                blnUseDefaultTypes(lngPathIdx) = (.INI_GetSetting(strSections(lngEmuIdx), "UseDefaultTypes", "0", m_strEmuFile) = "1")
                strNames(lngPathIdx) = strSections(lngEmuIdx)
                strExes(lngPathIdx) = .INI_GetSetting(strSections(lngEmuIdx), "EmulatorPathFile", "", m_strEmuFile)
                strScript(lngPathIdx) = .INI_GetSetting(strSections(lngEmuIdx), "ScriptFile", "", m_strEmuFile)
                strTypes(lngPathIdx) = .INI_GetSetting(strSections(lngEmuIdx), "SupportedTypes", "", m_strEmuFile)
                lngWaits(lngPathIdx) = Val(.INI_GetSetting(strSections(lngEmuIdx), "WaitFinish", "1", m_strEmuFile))
                'get the associated files
                ReDim strFiles(0)
                lngFileIdx = 1
                lngNumFiles = 0
                Do
                    strFile = .INI_GetSetting(strSections(lngEmuIdx), Trim$(lngFileIdx), "", m_strEmuFile)
                    If strFile <> "" Then
                        ReDim Preserve strFiles(lngFileIdx - 1)
                        strFiles(lngFileIdx - 1) = strFile
                        lngFileIdx = lngFileIdx + 1
                    Else
                        Exit Do
                    End If
                Loop
                If lngFileIdx > 1 Then
                    lngNumFiles = lngFileIdx - 1
                End If
                Call GetEmulatorInfo(strScript(lngPathIdx), blnUseShort(lngPathIdx), strNames(lngPathIdx), strExes(lngPathIdx), ma_udtEmulators(lngPathIdx), strFiles(), lngNumFiles, blnUseDefaultTypes(lngPathIdx), strTypes(lngPathIdx), lngWaits(lngPathIdx))
                lngPathIdx = lngPathIdx + 1
            Next lngEmuIdx
        End If
    End With

End Sub

Private Function ReadInWizardDATFile(strDATFile_in As String, intFileNumber_in As Integer) As String

    'read in the passed in GEMUS.DAT file into memory (g_udtWizard)
    'note: the file is assumed to be already open for input and the filepointer is
    'assumed to be at the beginning of the file
    'returns "" on success or an error message to display to the user

    Dim strLine As String
    Dim a_strFields() As String
    Dim lngTotalFields As Long
    Dim blnInvalidSection As Boolean
    Dim strLineError As String
    Dim strNoScriptFileError As String
    Dim lngLineNumber As Long
    Dim blnError As Boolean
    Dim lngCurrentSection As Long
    Dim a_strValues() As String
    Dim lngTotalValues As Long
    Dim lngIdx As Long
    Dim lngTotalEmulators As Long

    On Error GoTo Err_Handler

    strLineError = GetINIMessage("Wizard", "Error_DATLineError", "There was missing or incorrect data in the DAT file!") & vbCrLf & vbCrLf & strDATFile_in & " (" & GetINIMessage("Wizard", "Error_Line", "Line") & " "
    strNoScriptFileError = GetINIMessage("Wizard", "Error_NoScriptFile", "A GEMUS Script file could not be found in the \Scripts folder!") & vbCrLf & vbCrLf & strDATFile_in & " (" & GetINIMessage("Wizard", "Error_Line", "Line") & " "
    lngLineNumber = 0
    lngTotalEmulators = 0
    With g_udtWizard
        ReDim .a_udtSections(0)
        .lngTotalSections = 0
        lngCurrentSection = 0
        blnInvalidSection = True
        Do Until EOF(intFileNumber_in)
            Line Input #intFileNumber_in, strLine
            lngLineNumber = lngLineNumber + 1
            strLine = Trim$(strLine)
            If Left$(strLine, 1) = "[" And Right$(strLine, 1) = "]" Then
                If Len(strLine) > 3 Then
                    If UCase$(strLine) <> "[CREDITS]" And UCase$(strLine) <> "[DAT]" Then
                        'new section
                        blnInvalidSection = False
                        lngCurrentSection = .lngTotalSections
                        ReDim Preserve .a_udtSections(.lngTotalSections)
                        With .a_udtSections(.lngTotalSections)
                            .strName = Mid$(strLine, 2, Len(strLine) - 2)
                            .lngTotalEmulators = 0
                            ReDim .a_udtEmulators(0)
                        End With
                        .lngTotalSections = .lngTotalSections + 1
                    Else
                        'invalid (reserved) section
                        blnInvalidSection = True
                    End If
                End If
            ElseIf Left$(strLine, 1) = "¬" And Right$(strLine, 1) = "¬" Then
                If Not blnInvalidSection Then
                    'new emulator
                    blnError = False
                    a_strFields() = Split(strLine, "¬")
                    If UBound(a_strFields) > 8 Then
                        'name
                        If Trim$(a_strFields(1)) = "" Then
                            blnError = True
                        End If
                        If blnError = False Then
                            'url
                            If Trim$(a_strFields(2)) = "" Then
                                blnError = True
                            End If
                        End If
                        If blnError = False Then
                            'filename
                            If Trim$(a_strFields(4)) = "" Then
                                blnError = True
                            End If
                        End If
                        If blnError = False Then
                            'crc
                            If Trim$(a_strFields(5)) = "" Then
                                blnError = True
                            End If
                        End If
                        If blnError = False Then
                            'filetypes
                            If Trim$(a_strFields(6)) = "" Then
                                blnError = True
                            End If
                        End If
                        If blnError = False Then
                            'use short filenames
                            If UCase$(a_strFields(7)) <> "Y" And UCase$(a_strFields(7)) <> "N" Then
                                blnError = True
                            End If
                        End If
                        If blnError = False Then
                            'script file
                            If Trim$(a_strFields(8)) = "" Then
                                blnError = True
                            ElseIf m_xobjFunc.FileExists(.strScriptFolder & "\" & Trim$(a_strFields(8)), False, "") = False Then
                                ReadInWizardDATFile = strNoScriptFileError & lngLineNumber & ")"
                                Exit Function
                            End If
                        End If
                    Else
                        blnError = True
                    End If
                    If Not blnError Then
                        'save data in array
                        ReDim Preserve .a_udtSections(lngCurrentSection).a_udtEmulators(.a_udtSections(lngCurrentSection).lngTotalEmulators)
                        With .a_udtSections(lngCurrentSection).a_udtEmulators(.a_udtSections(lngCurrentSection).lngTotalEmulators)
                            lngTotalFields = UBound(a_strFields)
                            'crc values
                            a_strValues() = Split(a_strFields(5), "|")
                            lngTotalValues = UBound(a_strValues) + 1
                            .lngTotalCRCs = 0
                            For lngIdx = 0 To (lngTotalValues - 1)
                                Call Trim$(a_strValues(lngIdx))
                                If a_strValues(lngIdx) <> "" Then
                                    ReDim Preserve .a_strCRC32(.lngTotalCRCs)
                                    .a_strCRC32(.lngTotalCRCs) = a_strValues(lngIdx)
                                    .lngTotalCRCs = .lngTotalCRCs + 1
                                End If
                            Next lngIdx
                            If .lngTotalCRCs = 0 Then
                                'corrupt data in emulator line
                                ReadInWizardDATFile = strLineError & lngLineNumber & ")"
                                Exit Function
                            End If
                            'filetypes
                            a_strValues() = Split(a_strFields(6), ";")
                            lngTotalValues = UBound(a_strValues) + 1
                            .lngTotalFileTypes = 0
                            For lngIdx = 0 To (lngTotalValues - 1)
                                Call Trim$(a_strValues(lngIdx))
                                If a_strValues(lngIdx) <> "" Then
                                    If Left$(a_strValues(lngIdx), 1) = "." Then
                                        a_strValues(lngIdx) = Right$(a_strValues(lngIdx), Len(a_strValues(lngIdx)) - 1)
                                    End If
                                    If a_strValues(lngIdx) <> "" Then
                                        If Left$(a_strValues(lngIdx), 1) <> "." Then
                                            ReDim Preserve .a_strFileTypes(.lngTotalFileTypes)
                                            .a_strFileTypes(.lngTotalFileTypes) = a_strValues(lngIdx)
                                            .lngTotalFileTypes = .lngTotalFileTypes + 1
                                        End If
                                    End If
                                End If
                            Next lngIdx
                            If .lngTotalFileTypes = 0 Then
                                'corrupt data in emulator line
                                ReadInWizardDATFile = strLineError & lngLineNumber & ")"
                                Exit Function
                            End If
                            'use short filenames flag
                            .blnUseShort = (UCase$(a_strFields(7)) = "Y")
                            'emulator name
                            .strEmulatorName = a_strFields(1)
                            'url
                            .strURL = a_strFields(2)
                            'emulator filename
                            .strFilename = a_strFields(4)
                            'emulator info (optional)
                            .strInfo = a_strFields(3)
                            'gemus script file
                            .strScriptFile = a_strFields(8)
                            'init the found path
                            .strFound = ""
                            'associated files (optional)
                            .lngTotalAssFiles = 0
                            For lngIdx = 9 To (lngTotalFields - 1)
                                Call Trim$(a_strFields(lngIdx))
                                If a_strFields(lngIdx) <> "" Then
                                    a_strFields(lngIdx) = Replace(a_strFields(lngIdx), "/", "\")
                                    If Left$(a_strFields(lngIdx), 1) = "\" Then
                                        a_strFields(lngIdx) = Right$(a_strFields(lngIdx), Len(a_strFields(lngIdx)) - 1)
                                    End If
                                    If a_strFields(lngIdx) <> "" Then
                                        ReDim Preserve .a_strAssFiles(.lngTotalAssFiles)
                                        .a_strAssFiles(.lngTotalAssFiles) = a_strFields(lngIdx)
                                        .lngTotalAssFiles = .lngTotalAssFiles + 1
                                    End If
                                End If
                            Next lngIdx
                            'we have at least one emulator!
                            lngTotalEmulators = lngTotalEmulators + 1
                        End With
                        .a_udtSections(lngCurrentSection).lngTotalEmulators = .a_udtSections(lngCurrentSection).lngTotalEmulators + 1
                    Else
                        'corrupt data in emulator line
                        ReadInWizardDATFile = strLineError & lngLineNumber & ")"
                        Exit Function
                    End If
                End If
            End If
        Loop
        
    End With

    'have we got at least one section with one emulator?
    If lngTotalEmulators > 0 Then
        ReadInWizardDATFile = ""
        Exit Function
    Else
        ReadInWizardDATFile = GetINIMessage("Wizard", "Error_NoEmulatorsFound", "No valid emulator definitions could be found in the DAT file!" & vbCrLf & vbCrLf & strDATFile_in)
        Exit Function
    End If

Err_Handler:
    ReadInWizardDATFile = Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Sub ReplaceEnvironmentVars(strText_inout As String, ByVal blnAddQuotes_in As Boolean, ByVal blnDoVBSendKeys As Boolean)

    '-----------------------------------------------------
    'replaces all environment variables in the text string
    'passed in to their actual values.
    '
    'Modify:    strText_inout:  The text string to change
    '----------------------------------------------------

    Dim strShortPath As String          'for short filenames
    Dim strShortFile As String          'for short filenames
    Dim strShortPathFile As String      'for short filenames
    Dim strShortFileNoExt As String     'for short filenames
    Dim lngFileIdx As Long              'loop index
    Dim lngIdx As Long                  'loop index
    Dim strFind As String
    Dim strReplace As String

    'db drive letter
    strText_inout = Replace(strText_inout, "%dbdriveletter%", Left$(m_strAppPathWithBackSlash, 1))

    If Not m_blnUseShort Then
        'music path
        With m_udtRegPathSid
            If .ePathType <> e_PathLocation_Invalid Then
                strText_inout = Replace(strText_inout, "%storedmusic%", VBSendKeysEscape(m_udtRegPathSid.strRealPath, blnDoVBSendKeys))
            End If
        End With
        'photos path
        With m_udtRegPathPhotos
            If .ePathType <> e_PathLocation_Invalid Then
                strText_inout = Replace(strText_inout, "%storedphotos%", VBSendKeysEscape(m_udtRegPathPhotos.strRealPath, blnDoVBSendKeys))
            End If
        End With
        'games paths
        For lngIdx = 0 To m_lngTotalGamePaths - 1
            With ma_udtRegPathGame(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedgames(" & lngIdx & ")%", VBSendKeysEscape(.strRealPath, blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'picture paths
        For lngIdx = 0 To m_lngTotalScrnShotPaths - 1
            With ma_udtRegPathScrnShot(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedpictures(" & lngIdx & ")%", VBSendKeysEscape(.strRealPath, blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'extras paths
        For lngIdx = 0 To m_lngTotalExtrasPaths - 1
            With ma_udtRegPathExtras(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedextras(" & lngIdx & ")%", VBSendKeysEscape(.strRealPath, blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'gbgame path
        strText_inout = Replace(strText_inout, "%gbgamepath%", VBSendKeysEscape(m_strGBGamePath, blnDoVBSendKeys))
        'db path
        strText_inout = Replace(strText_inout, "%dbpath%", VBSendKeysEscape(m_strAppPathWithBackSlash & m_strGBName, blnDoVBSendKeys))
        'emulator path
        strText_inout = Replace(strText_inout, "%emupath%", VBSendKeysEscape(ma_udtEmulators(m_lngEmuIdx).strFullPathOnly, blnDoVBSendKeys))
        'full game path\filename
        If InStr(strText_inout, "%gamepathfile%") <> 0 Then
            If blnAddQuotes_in = True Then
                If InStr(m_strEnvGFFull, " ") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamepathfile%", """" & VBSendKeysEscape(m_strEnvGFFull, blnDoVBSendKeys) & """")
                Else
                    strText_inout = Replace(strText_inout, "%gamepathfile%", VBSendKeysEscape(m_strEnvGFFull, blnDoVBSendKeys))
                End If
            Else
                strText_inout = Replace(strText_inout, "%gamepathfile%", VBSendKeysEscape(m_strEnvGFFull, blnDoVBSendKeys))
            End If
        End If
        'full filename
        If InStr(strText_inout, "%gamefile%") <> 0 Then
            If blnAddQuotes_in = True Then
                If InStr(m_strEnvGFFile, " ") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamefile%", """" & VBSendKeysEscape(m_strEnvGFFile, blnDoVBSendKeys) & """")
                Else
                    strText_inout = Replace(strText_inout, "%gamefile%", VBSendKeysEscape(m_strEnvGFFile, blnDoVBSendKeys))
                End If
            Else
                strText_inout = Replace(strText_inout, "%gamefile%", VBSendKeysEscape(m_strEnvGFFile, blnDoVBSendKeys))
            End If
        End If
        'full path
        If InStr(strText_inout, "%gamepath%") <> 0 Then
            If blnAddQuotes_in = True Then
                If InStr(m_strEnvGFPath, " ") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamepath%", """" & VBSendKeysEscape(m_strEnvGFPath, blnDoVBSendKeys) & """")
                Else
                    strText_inout = Replace(strText_inout, "%gamepath%", VBSendKeysEscape(m_strEnvGFPath, blnDoVBSendKeys))
                End If
            Else
                strText_inout = Replace(strText_inout, "%gamepath%", VBSendKeysEscape(m_strEnvGFPath, blnDoVBSendKeys))
            End If
        End If
        'filename without extension
        If InStr(strText_inout, "%gamefilenoext%") <> 0 Then
            If blnAddQuotes_in = True Then
                If InStr(m_strEnvGFNoExt, " ") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamefilenoext%", """" & VBSendKeysEscape(m_strEnvGFNoExt, blnDoVBSendKeys) & """")
                Else
                    strText_inout = Replace(strText_inout, "%gamefilenoext%", VBSendKeysEscape(m_strEnvGFNoExt, blnDoVBSendKeys))
                End If
            Else
                strText_inout = Replace(strText_inout, "%gamefilenoext%", VBSendKeysEscape(m_strEnvGFNoExt, blnDoVBSendKeys))
            End If
        End If
        'extra files
        If m_lngFldNumGameFiles > 1 Then
            For lngFileIdx = 1 To (m_lngFldNumGameFiles - 1)
                'full game path\filename
                If InStr(strText_inout, "%gamepathfile(" & lngFileIdx & ")%") <> 0 Then
                    If blnAddQuotes_in = True Then
                        If InStr(ma_strExtraGFFull(lngFileIdx), " ") <> 0 Then
                            strText_inout = Replace(strText_inout, "%gamepathfile(" & lngFileIdx & ")%", """" & VBSendKeysEscape(ma_strExtraGFFull(lngFileIdx), blnDoVBSendKeys) & """")
                        Else
                            strText_inout = Replace(strText_inout, "%gamepathfile(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFFull(lngFileIdx), blnDoVBSendKeys))
                        End If
                    Else
                        strText_inout = Replace(strText_inout, "%gamepathfile(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFFull(lngFileIdx), blnDoVBSendKeys))
                    End If
                End If
                'full filename
                If InStr(strText_inout, "%gamefile(" & lngFileIdx & ")%") <> 0 Then
                    If blnAddQuotes_in = True Then
                        If InStr(ma_strExtraGFFile(lngFileIdx), " ") <> 0 Then
                            strText_inout = Replace(strText_inout, "%gamefile(" & lngFileIdx & ")%", """" & VBSendKeysEscape(ma_strExtraGFFile(lngFileIdx), blnDoVBSendKeys) & """")
                        Else
                            strText_inout = Replace(strText_inout, "%gamefile(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFFile(lngFileIdx), blnDoVBSendKeys))
                        End If
                    Else
                        strText_inout = Replace(strText_inout, "%gamefile(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFFile(lngFileIdx), blnDoVBSendKeys))
                    End If
                End If
                'full path
                If InStr(strText_inout, "%gamepath(" & lngFileIdx & ")%") <> 0 Then
                    If blnAddQuotes_in = True Then
                        If InStr(ma_strExtraGFPath(lngFileIdx), " ") <> 0 Then
                            strText_inout = Replace(strText_inout, "%gamepath(" & lngFileIdx & ")%", """" & VBSendKeysEscape(ma_strExtraGFPath(lngFileIdx), blnDoVBSendKeys) & """")
                        Else
                            strText_inout = Replace(strText_inout, "%gamepath(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFPath(lngFileIdx), blnDoVBSendKeys))
                        End If
                    Else
                        strText_inout = Replace(strText_inout, "%gamepath(" & lngFileIdx & ")%", VBSendKeysEscape(ma_strExtraGFPath(lngFileIdx), blnDoVBSendKeys))
                    End If
                End If
                'filename without extension
                If InStr(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%") <> 0 Then
                    If blnAddQuotes_in = True Then
                        If InStr(m_strExtraGFNoExt(lngFileIdx), " ") <> 0 Then
                            strText_inout = Replace(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%", """" & VBSendKeysEscape(m_strExtraGFNoExt(lngFileIdx), blnDoVBSendKeys) & """")
                        Else
                            strText_inout = Replace(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%", VBSendKeysEscape(m_strExtraGFNoExt(lngFileIdx), blnDoVBSendKeys))
                        End If
                    Else
                        strText_inout = Replace(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%", VBSendKeysEscape(m_strExtraGFNoExt(lngFileIdx), blnDoVBSendKeys))
                    End If
                End If
            Next lngFileIdx
        End If
    Else
        'music path
        With m_udtRegPathSid
            If .ePathType <> e_PathLocation_Invalid Then
                strText_inout = Replace(strText_inout, "%storedmusic%", VBSendKeysEscape(m_xobjFunc.Long2Short(.strRealPath), blnDoVBSendKeys))
            End If
        End With
        'photos path
        With m_udtRegPathPhotos
            If .ePathType <> e_PathLocation_Invalid Then
                strText_inout = Replace(strText_inout, "%storedphotos%", VBSendKeysEscape(m_xobjFunc.Long2Short(.strRealPath), blnDoVBSendKeys))
            End If
        End With
        'games paths
        For lngIdx = 0 To m_lngTotalGamePaths - 1
            With ma_udtRegPathGame(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedgames(" & lngIdx & ")%", VBSendKeysEscape(m_xobjFunc.Long2Short(.strRealPath), blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'picture paths
        For lngIdx = 0 To m_lngTotalScrnShotPaths - 1
            With ma_udtRegPathScrnShot(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedpictures(" & lngIdx & ")%", VBSendKeysEscape(m_xobjFunc.Long2Short(.strRealPath), blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'extras paths
        For lngIdx = 0 To m_lngTotalExtrasPaths - 1
            With ma_udtRegPathExtras(lngIdx)
                If .ePathType <> e_PathLocation_Invalid Then
                    strText_inout = Replace(strText_inout, "%storedextras(" & lngIdx & ")%", VBSendKeysEscape(m_xobjFunc.Long2Short(.strRealPath), blnDoVBSendKeys))
                End If
            End With
        Next lngIdx
        'gbgame path
        strText_inout = Replace(strText_inout, "%gbgamepath%", VBSendKeysEscape(m_xobjFunc.Long2Short(m_strGBGamePath), blnDoVBSendKeys))
        'db path
        strText_inout = Replace(strText_inout, "%dbpath%", VBSendKeysEscape(m_xobjFunc.Long2Short(m_strAppPathWithBackSlash & m_strGBName), blnDoVBSendKeys))
        'emulator path
        strText_inout = Replace(strText_inout, "%emupath%", VBSendKeysEscape(m_xobjFunc.Long2Short(ma_udtEmulators(m_lngEmuIdx).strFullPathOnly), blnDoVBSendKeys))
        'get short vars
        strShortPathFile = m_xobjFunc.Long2Short(m_strEnvGFFull)
        Call m_xobjFunc.GetFileAndPath(strShortPathFile, strShortPath, strShortFile)
        strShortFileNoExt = Left$(strShortFile, InStrRev(strShortFile, ".") - 1)
        'short game path\filename
        If InStr(strText_inout, "%gamepathfile%") <> 0 Then
            strText_inout = Replace(strText_inout, "%gamepathfile%", VBSendKeysEscape(strShortPathFile, blnDoVBSendKeys))
        End If
        'short filename
        If InStr(strText_inout, "%gamefile%") <> 0 Then
            strText_inout = Replace(strText_inout, "%gamefile%", VBSendKeysEscape(strShortFile, blnDoVBSendKeys))
        End If
        'short path
        If InStr(strText_inout, "%gamepath%") <> 0 Then
            strText_inout = Replace(strText_inout, "%gamepath%", VBSendKeysEscape(strShortPath, blnDoVBSendKeys))
        End If
        'short filename without extension
        If InStr(strText_inout, "%gamefilenoext%") <> 0 Then
            strText_inout = Replace(strText_inout, "%gamefilenoext%", VBSendKeysEscape(strShortFileNoExt, blnDoVBSendKeys))
        End If
        'extra files
        If m_lngFldNumGameFiles > 1 Then
            For lngFileIdx = 1 To (m_lngFldNumGameFiles - 1)
                'get short vars
                strShortPathFile = m_xobjFunc.Long2Short(ma_strExtraGFFull(lngFileIdx))
                Call m_xobjFunc.GetFileAndPath(strShortPathFile, strShortPath, strShortFile)
                strShortFileNoExt = Left$(strShortFile, InStrRev(strShortFile, ".") - 1)
                'short game path\filename
                If InStr(strText_inout, "%gamepathfile(" & lngFileIdx & ")%") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamepathfile(" & lngFileIdx & ")%", VBSendKeysEscape(strShortPathFile, blnDoVBSendKeys))
                End If
                'short filename
                If InStr(strText_inout, "%gamefile(" & lngFileIdx & ")%") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamefile(" & lngFileIdx & ")%", VBSendKeysEscape(strShortFile, blnDoVBSendKeys))
                End If
                'short path
                If InStr(strText_inout, "%gamepath(" & lngFileIdx & ")%") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamepath(" & lngFileIdx & ")%", VBSendKeysEscape(strShortPath, blnDoVBSendKeys))
                End If
                'short filename without extension
                If InStr(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%") <> 0 Then
                    strText_inout = Replace(strText_inout, "%gamefilenoext(" & lngFileIdx & ")%", VBSendKeysEscape(strShortFileNoExt, blnDoVBSendKeys))
                End If
            Next lngFileIdx
        End If
    End If

    'image index
    If InStr(strText_inout, "%imageindex%") <> 0 Then
        strText_inout = Replace(strText_inout, "%imageindex%", Trim$(m_lngImageIndex))
    End If
    'image name
    If InStr(strText_inout, "%imagename%") <> 0 Then
        strText_inout = Replace(strText_inout, "%imagename%", VBSendKeysEscape(m_strImageName, blnDoVBSendKeys))
    End If
    'c64 image name (lower case)
    If InStr(strText_inout, "%c64imagename%") <> 0 Then
        strText_inout = Replace(strText_inout, "%c64imagename%", VBSendKeysEscape(m_strC64ImageName, blnDoVBSendKeys))
    End If
    'command line
    If InStr(strText_inout, "%commandline%") <> 0 Then
        strText_inout = Replace(strText_inout, "%commandline%", VBSendKeysEscape(m_strCommandLine, blnDoVBSendKeys))
    End If
    'number of game files
    If InStr(strText_inout, "%numfiles%") <> 0 Then
        strText_inout = Replace(strText_inout, "%numfiles%", m_lngFldNumGameFiles)
    End If

    'carriage return and line feed
    If InStr(strText_inout, "%crlf%") <> 0 Then
        strText_inout = Replace(strText_inout, "%crlf%", vbCrLf)
    End If
    '2 x carriage return and line feed (i.e. a blank line)
    If InStr(strText_inout, "%crlfx2%") <> 0 Then
        strText_inout = Replace(strText_inout, "%crlfx2%", vbCrLf & vbCrLf)
    End If
    'tab character
    If InStr(strText_inout, "%tab%") <> 0 Then
        strText_inout = Replace(strText_inout, "%tab%", vbTab)
    End If

    'gemus key=value pairs
    For lngIdx = 0 To (m_lngTotalNotesKeys - 1)
        strFind = "%" & LCase$(ma_strNotesKeys(lngIdx).strKey) & "_value%"
        strReplace = ma_strNotesKeys(lngIdx).strValue
        strText_inout = Replace(strText_inout, strFind, VBSendKeysEscape(strReplace, False))
    Next lngIdx

End Sub

Private Sub RezipWait(ByVal lngEmuIdx_in As Long, ByVal dblProcessID_in As Double, ByVal lngHandle_in As Long)

    Dim blnStopped As Boolean

    If m_blnAttemptRepack Then
        If ma_udtEmulators(m_lngEmuIdx).lngWait = 1 Then
            'wait for it to finish
            Load frmRezipWait
            With frmRezipWait
                .MousePointer = vbHourglass
                .cmdStop.Enabled = True
                .m_blnStopped = False
                .lblPrompt = GetINIMessage("Running Games", "RezipWaitMessage", "Waiting for the emulator to close") & "..."
                .Caption = GetINIMessage("Running Games", "RunningPrompt", "Running") & " " & ma_udtEmulators(m_lngEmuIdx).strName & "..."
                .cmdStop.Caption = GetINIMessage("Running Games", "RezipWaitStop", "&Stop Waiting")
                .m_dblProcessID = dblProcessID_in
                .m_lngHandle = lngHandle_in
                .Show vbModal
            End With
            Unload frmRezipWait
        Else
            'don't wait
            MsgBox GetINIMessage("Running Games", "RezipNoWaitMessage", "Click OK when the emulator has finished."), vbInformation Or vbOKOnly, GetINIMessage("Running Games", "RunningPrompt", "Running") & " " & ma_udtEmulators(m_lngEmuIdx).strName & "..."
        End If
    End If

End Sub

Public Function RunEmulator(ByVal lngEmuIdx_in As Long, strGamePathFile_in As String, ByVal lngImageIndex_in As Long, _
                            strDefaultFile_inout As String, strExtractDestFolder_in As String, ByVal blnFromDB_in As Boolean, _
                            Optional ByVal varNumPlayersFrom_in As Variant, Optional ByVal varNumPlayersTo_in As Variant, _
                            Optional varControl_in As Variant, Optional varGameComment_in As Variant, _
                            Optional varVersionComment_in As Variant, Optional varPALNTSC_in As Variant, _
                            Optional ByVal varTrueDriveEmu_in As Variant, Optional varNotes_in As Variant) As String

    '----------------------------------------------------------------------------------------------
    'Runs the script to run the passed in game with the passed in emulator
    '
    'Inputs:    lngEmuIdx_in:           Index of emulator to attempt to run
    '           strGamePathFile_in:     full path and filename of game to run
    '           lngImageIndex_in:       index of image in gamefile to run
    '           strDefaultFile_inout:   file within archive
    '           strExtractDestFolder_in: gbgame folder
    '           blnFromDB_in:           flag specifying whether the following db fields are present
    '           varNumPlayersFrom_in:   DB Field: Number of players From
    '           varNumPlayersTo_in:     DB Field: Number of players To
    '           varControl_in:          DB Field: Control method
    '           varGameComment_in:      DB Field: Game Comment
    '           varVersionComment_in:   DB Field: Version Comment
    '           varPALNTSC_in:          DB Field: PAL/NTSC/BOTH
    '           varTrueDriveEmu_in:     DB Field: True Drive Emulation (YES/NO)
    '
    'Returns:   Error string or "" if successful
    '
    'Assumptions:   Assumes passed in gamefile exists
    '----------------------------------------------------------------------------------------------

    On Error GoTo Err_Handler:

    Dim lngFileIdx As Long          'loop index
    Dim strScript As String         'the emulator's script
    Dim strError As String          'function result string
    Dim lngErrorLine As Long        'error line
    Dim lngLineIdx As Long          'loop index
    Dim lngNextLine As Long         'for recursive call
    Dim strLastChar As String       'for extra gamefiles
    Dim lngExtraIdx As Long         'for extra gamefiles
    Dim strPathWBS As String        'for extra gamefiles
    Dim objExtract As clsExtract    'for extracting archives
    Dim strFileToRun As String      'after extraction
    Dim lngEmuIdx As Long           'loop index
    Dim a_strTypes() As String      'for getting the supported file types for an emulator
    Dim a_strTempTypes() As String  'for getting the supported file types for an emulator
    Dim lngTotalTypes As Long       'total supported emu-file-types
    Dim lngFTIdx As Long            'file-type loop index
    Dim strUpper As String          'upper case file-type for comparison
    Dim strExt As String            'game-type filename extension
    Dim strNotes As String
    Dim a_strArrTemp() As String
    Dim f As Integer
    Dim lngIdx As Long
    Dim lngPos As Long
    Dim strKey As String
    Dim strVal As String
    Dim strTestName As String
    Dim strTestNameDisplay As String
    Dim blnFound As Boolean
    Dim a_Before() As UnzippedFileType
    Dim lngTotalBefore As Long
    Dim strRezipWBS As String
    Dim a_strExtractedTypes() As String
    Dim lngTotalExtractedTypes As Long
    Dim a_udtSetKV() As NotesKeysType
    Dim lngTotalSetKV As Long
    Dim lngFindIdx As Long
    Dim lngSetKVIdx As Long
    Dim blnSuccess As Boolean
    Dim strNotUsed As String
    Dim blnUserPick As Boolean
    Dim strDestFolderWBS As String
    Dim blnIsRepackPossible As Boolean
    Dim blnWasRepacked As Boolean
    Dim blnRepackAsk As Boolean
    Dim blnRepackNotify As Boolean

    Dim Unpacker As clsUnpack

    Dim Length, i, l As Integer
    Dim strSrc, strDst, c As String
    
    m_strSimulateKV = ""

    Call CreateLogFile
    Call PushLogFileTrace("clsGemus.RunEmulator")
    Call LogStatus("Arg Silent Mode: [" & IIf(m_blnSilentMode, "True", "False") & "]")
    Call LogStatus("Arg strGamePathFile_in: [" & strGamePathFile_in & "]")
    Call LogStatus("Arg lngImageIndex_in: [" & lngImageIndex_in & "]")
    Call LogStatus("Arg strDefaultFile_inout: [" & strDefaultFile_inout & "]")
    Call LogStatus("Arg strExtractDestFolder_in: [" & strExtractDestFolder_in & "]")
    Call LogStatus("Arg blnFromDB_in: [" & IIf(blnFromDB_in, "True", "False") & "]")
    Call LogStatus("Arg varNumPlayersFrom_in: [" & IIf(IsMissing(varNumPlayersFrom_in), "--not-specified--", varNumPlayersFrom_in) & "]")
    Call LogStatus("Arg varNumPlayersTo_in: [" & IIf(IsMissing(varNumPlayersTo_in), "--not-specified--", varNumPlayersTo_in) & "]")
    Call LogStatus("Arg varControl_in: [" & IIf(IsMissing(varControl_in), "--not-specified--", varControl_in) & "]")
    Call LogStatus("Arg varGameComment_in: [" & IIf(IsMissing(varGameComment_in), "--not-specified--", varGameComment_in) & "]")
    Call LogStatus("Arg varVersionComment_in: [" & IIf(IsMissing(varVersionComment_in), "--not-specified--", varVersionComment_in) & "]")
    Call LogStatus("Arg varPALNTSC_in: [" & IIf(IsMissing(varPALNTSC_in), "--not-specified--", varPALNTSC_in) & "]")
    Call LogStatus("Arg varTrueDriveEmu_in: [" & IIf(IsMissing(varTrueDriveEmu_in), "--not-specified--", varTrueDriveEmu_in) & "]")
    Call LogStatus("Arg varNotes_in: [" & IIf(IsMissing(varNotes_in), "--not-specified--", varNotes_in) & "]")

    '-------------------------
    'unpack the game if needed
    '-------------------------

    Set Unpacker = New clsUnpack
    m_strGBGamePath = strExtractDestFolder_in
    m_blnScriptFileErrorDone = False
    strDestFolderWBS = strExtractDestFolder_in
    If Right$(strDestFolderWBS, 1) <> "\" Then strDestFolderWBS = strDestFolderWBS & "\"
    lngTotalExtractedTypes = GetAllValidFileTypesArray(a_strExtractedTypes())
    Call LogStatus("GetAllValidFileTypesArray: " & Join(a_strExtractedTypes(), ";"))

    'get the actual runnable game file by running the unpacker
    Call LogStatus("Starting Unpacker...")
    blnSuccess = Unpacker.Unpack(strGamePathFile_in, strDefaultFile_inout, a_strExtractedTypes, lngTotalExtractedTypes, _
                                strDestFolderWBS, blnIsRepackPossible, strFileToRun, blnUserPick, strNotUsed, strError)
    If Not blnSuccess Then
        Call LogStatus("Unpacker failed with error [" & strError & "]")
        RunEmulator = strError
        Set Unpacker = Nothing
        Call CloseLogFile
        Exit Function
    ElseIf blnUserPick Then
        'ask user for runnable file if not in silent mode, else return "no runnable file" error
        Call LogStatus("User needs to pick runnable file within archive")
        If Not m_blnSilentMode Then
            'to do, ask for file, then try to unpack again
            blnSuccess = Unpacker.GetRunnableFileInArchiveFromUser(strGamePathFile_in, strDestFolderWBS, _
                                a_strExtractedTypes, strFileToRun)
            If blnSuccess Then
                'save it to the output var so it can be saved after playing
                strDefaultFile_inout = strFileToRun
                Call LogStatus("User picked runnable file [" & strDefaultFile_inout & "]")
                'got the runnable file, try again to run it
                Call LogStatus("Starting Unpacker (again)...")
                blnSuccess = Unpacker.Unpack(strGamePathFile_in, strDefaultFile_inout, a_strExtractedTypes, _
                                lngTotalExtractedTypes, strDestFolderWBS, blnIsRepackPossible, strFileToRun, _
                                blnUserPick, strNotUsed, strError)
                If Not blnSuccess Then
                    'unable to run the selected runnable file (should never happen if the archive is sound)
                    RunEmulator = strError
                    Set Unpacker = Nothing
                    Call LogStatus("Unpacker failed with error [" & strError & "]")
                    Call CloseLogFile
                    Exit Function
                ElseIf blnUserPick Then
                    RunEmulator = GetINIMessage("Running Games", "RunnableFileNotDetermined", "Runnable file within archive could not be determined")
                    Set Unpacker = Nothing
                    Call LogStatus("Unpacker unable to determine runnable file within archive. Maybe user did not pick a file.")
                    Call CloseLogFile
                    Exit Function
                End If
            Else
                'user cancelled picking runnable file, quit
                Call LogStatus("Unpacker failed with error [" & strError & "]")
                RunEmulator = GetINIMessage("Running Games", "AbortedRunningGamePleaseSelect", "Aborted running game. Please select a runnable file for this archive")
                Set Unpacker = Nothing
                Call LogStatus("Unpacker unable to determine runnable file within archive. User aborted running game.")
                Call CloseLogFile
                Exit Function
            End If
        Else
            'silent mode, so we can't ask the user for the runnable file within the archive
            RunEmulator = GetINIMessage("Running Games", "RunnableFileNotAutoDetermined", "Runnable file within archive could not be automatically determined")
            Set Unpacker = Nothing
            Call LogStatus("Unpacker unable to auto-determine runnable file within archive. User not asked as running in silent mode.")
            Call CloseLogFile
            Exit Function
        End If
    Else
        Call LogStatus("Runnable file determined [" & strFileToRun & "]")
    End If

    'see if repacking is desired and possible
    m_blnAttemptRepack = False
    If m_blnUserWantsToRepack And blnIsRepackPossible Then
        If Not m_xobjFunc.DriveTypeIsCDRom(Left$(strGamePathFile_in, 2)) Then
            m_blnAttemptRepack = True
            Call LogStatus("Will attempt to repack after emulator closes")
            Call Unpacker.StorePreUnpackedFilesSnapshot(Unpacker.GetFinalFolder())
            Call LogStatus("Stored unpacked files snapshot from [" & Unpacker.GetFinalFolder() & "]")
        End If
    End If

    '-----------------------------
    'get the notes key=value array
    '-----------------------------

    If blnFromDB_in = True Then
        strNotes = CStr(varNotes_in)
    Else
        m_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "Notes", strNotes, True, ""
    End If

    'read in any global pairs first, then the per-game pairs
    m_lngTotalNotesKeys = GetKVSet(ma_strNotesKeys, "global")
    Call LogStatus("Loaded " & m_lngTotalNotesKeys & " global k=v pairs")

    If strNotes <> "" Then
        'look for the "kvset" keys
        Call LogStatus("Looking for kvsets to load...")
        a_strArrTemp = Split(strNotes, vbCrLf)
        For lngIdx = 0 To UBound(a_strArrTemp)
            a_strArrTemp(lngIdx) = Trim$(a_strArrTemp(lngIdx))
            lngPos = InStr(a_strArrTemp(lngIdx), "=")
            If lngPos > 1 Then
                strKey = Left$(a_strArrTemp(lngIdx), lngPos - 1)
                strKey = Trim$(strKey)
                If strKey = "kvset" Then
                    strVal = Right$(a_strArrTemp(lngIdx), Len(a_strArrTemp(lngIdx)) - lngPos)
                    strVal = Trim$(strVal)
                    If strVal <> "" And LCase$(strVal) <> "global" Then
                        lngTotalSetKV = GetKVSet(a_udtSetKV(), strVal)
                        Call LogStatus("Loaded " & lngTotalSetKV & " k=v pairs in kvset """ & strVal & """")
                        'merge these kv set pairs into the existing ones?
                        If lngTotalSetKV > 0 Then
                            For lngSetKVIdx = 0 To lngTotalSetKV - 1
                                blnFound = False
                                For lngFindIdx = 0 To m_lngTotalNotesKeys - 1
                                    If LCase$(ma_strNotesKeys(lngFindIdx).strKey) = LCase$(a_udtSetKV(lngSetKVIdx).strKey) Then
                                        'overwrite existing value with the set value
                                        ma_strNotesKeys(lngFindIdx).strValue = a_udtSetKV(lngSetKVIdx).strValue
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngFindIdx
                                If Not blnFound Then
                                    'add to main key=value pairs
                                    ReDim Preserve ma_strNotesKeys(m_lngTotalNotesKeys)
                                    ma_strNotesKeys(m_lngTotalNotesKeys).strKey = a_udtSetKV(lngSetKVIdx).strKey
                                    ma_strNotesKeys(m_lngTotalNotesKeys).strValue = a_udtSetKV(lngSetKVIdx).strValue
                                    m_lngTotalNotesKeys = m_lngTotalNotesKeys + 1
                                End If
                            Next lngSetKVIdx
                        End If
                    End If
                End If
            End If
        Next lngIdx

        'merge/add the per-game key=value pairs into the existing ones
        For lngIdx = 0 To UBound(a_strArrTemp)
            a_strArrTemp(lngIdx) = Trim$(a_strArrTemp(lngIdx))
            lngPos = InStr(a_strArrTemp(lngIdx), "=")
            If lngPos > 1 Then
                strKey = Left$(a_strArrTemp(lngIdx), lngPos - 1)
                strKey = Trim$(strKey)
                If strKey <> "" And LCase$(strKey) <> "kvset" And Left$(strKey, 1) <> ";" Then
                    strVal = Right$(a_strArrTemp(lngIdx), Len(a_strArrTemp(lngIdx)) - lngPos)
                    strVal = Trim$(strVal)
                    blnFound = False
                    For lngFindIdx = 0 To m_lngTotalNotesKeys - 1
                        If LCase$(ma_strNotesKeys(lngFindIdx).strKey) = LCase$(strKey) Then
                            'overwrite existing value with the more specific one
                            ma_strNotesKeys(lngFindIdx).strValue = strVal
                            blnFound = True
                            Exit For
                        End If
                    Next lngFindIdx
                    If Not blnFound Then
                        'add to main key=value pairs
                        ReDim Preserve ma_strNotesKeys(m_lngTotalNotesKeys)
                        ma_strNotesKeys(m_lngTotalNotesKeys).strKey = strKey
                        ma_strNotesKeys(m_lngTotalNotesKeys).strValue = strVal
                        m_lngTotalNotesKeys = m_lngTotalNotesKeys + 1
                    End If
                End If
            End If
        Next lngIdx
        
        Call LogStatus("Finished loading all k=v pairs. Total loaded: " & m_lngTotalNotesKeys)
    Else
        Call LogStatus("No game-level k=v pairs to process")
    End If

    '----------------
    'get the emulator
    '----------------

    'if not a specific emulator requested, check the key=value pairs for a specific emulator
    If lngEmuIdx_in = -1 Then
        'find emu=xxx pair and grab the emu name
        strTestName = ""
        For lngIdx = 0 To (m_lngTotalNotesKeys - 1)
            If UCase$(ma_strNotesKeys(lngIdx).strKey) = "EMU" Then
                If ma_strNotesKeys(lngIdx).strValue <> "" Then
                    strTestName = UCase$(ma_strNotesKeys(lngIdx).strValue)
                    strTestNameDisplay = ma_strNotesKeys(lngIdx).strValue
                    Call LogStatus("Found [emu=" & strTestNameDisplay & "] k=v pair")
                    Exit For
                End If
            End If
        Next lngIdx
        'find a matching emulator
        If strTestName <> "" Then
            blnFound = False
            For lngEmuIdx = 0 To (m_lngTotalEmulators - 1)
                With ma_udtEmulators(lngEmuIdx)
                    If UCase$(.strName) = strTestName Then
                        'found a match! Use this emulator!
                        blnFound = True
                        lngEmuIdx_in = lngEmuIdx
                        Exit For
                    End If
                End With
            Next lngEmuIdx
            If Not blnFound Then
                RunEmulator = GetINIMessage("Running Games", "NoMatchingEmulator", "No emulator name was found to match your 'emu=' key=value pair.") & vbCrLf & vbCrLf & "[" & "emu=" & strTestNameDisplay & "]"
                Set Unpacker = Nothing
                Call CloseLogFile
                Exit Function
            End If
        End If
    End If
    
    'grab and validate the emulator details!
    strExt = Right$(strFileToRun, Len(strFileToRun) - InStrRev(strFileToRun, "."))
    If (lngEmuIdx_in > -1) And (lngEmuIdx_in < m_lngTotalEmulators) Then
        'specific emulator requested
        With ma_udtEmulators(lngEmuIdx_in)
            If .blnUseDefaultTypes = False Then
                If FileTypeSupported(.strTypes, strExt) = True Then
                    m_lngEmuIdx = lngEmuIdx_in
                Else
                    RunEmulator = GetINIMessage("Running Games", "Status_NotSupportedByEmulator", "This game type is not supported by this emulator.") & " [" & strExt & "]"
                    Set Unpacker = Nothing
                    Call CloseLogFile
                    Exit Function
                End If
            Else
                'already checked in extractor class
                m_lngEmuIdx = lngEmuIdx_in
            End If
        End With
        Call LogStatus("Using emulator with index " & m_lngEmuIdx & " [" & ma_udtEmulators(m_lngEmuIdx).strName & "]")
    Else
        'find first emulator that supports game type
        Call LogStatus("Finding first emulator that supports gamefile type [" & strExt & "]")
        m_lngEmuIdx = -1
        For lngEmuIdx = 0 To (m_lngTotalEmulators - 1)
            With ma_udtEmulators(lngEmuIdx)
                If .blnUseDefaultTypes = False Then
                    If FileTypeSupported(.strTypes, strExt) = True Then
                        m_lngEmuIdx = lngEmuIdx
                        Exit For
                    End If
                Else
                    'already checked in extractor class
                    m_lngEmuIdx = lngEmuIdx
                    Exit For
                End If
            End With
        Next lngEmuIdx
        'ok?
        If m_lngEmuIdx = -1 Then
            RunEmulator = GetINIMessage("Running Games", "Status_NotSupportedByAnyEmulator", "An emulator that supports this game type was not found.") & " [" & strExt & "]"
            Set Unpacker = Nothing
            Call CloseLogFile
            Exit Function
        End If
        Call LogStatus("Using emulator with index " & m_lngEmuIdx & " [" & ma_udtEmulators(m_lngEmuIdx).strName & "]")
    End If

    '------------------------------------------
    'set all the field variables for the script
    '------------------------------------------
    
    m_blnUseShort = ma_udtEmulators(m_lngEmuIdx).blnUseShortFilenames

    'Database Fields
    m_lngImageIndex = lngImageIndex_in
    m_strImageName = GetImageNameFromGameFile(strFileToRun, m_lngImageIndex)
    'm_strC64ImageName = LCase$(m_strImageName)
        
    'Convert upper to lower case limited to just Alpha A-Z and [] characters (LCase had modified certain C64 characters)
    strDst = ""
    If m_strImageName <> "" Then
        strSrc = m_strImageName
    
        Length = Len(strSrc)
        If Length > 0 Then
            Length = Length - 1
        End If
        
        For i = 0 To Length
            c = Mid$(strSrc, i + 1, 1)
            l = Asc(c)
            If l > 64 Then
                If l < 94 Then
                    c = Chr$(l + 32)
                ElseIf l = 95 Then
                    c = "?"
                End If
            End If
            strDst = strDst & c
        Next i
    End If
    m_strC64ImageName = strDst

    If blnFromDB_in = True Then
        'store passed in field variables for database fields
        m_lngFldNumPlayersFrom = CLng(varNumPlayersFrom_in)
        m_lngFldNumPlayersTo = CLng(varNumPlayersTo_in)
        m_lngFldControl = CLng(varControl_in)
        m_strFldGameComment = CStr(varGameComment_in)
        m_strFldVersionComment = CStr(varVersionComment_in)
        m_lngFldPALNTSC = CLng(varPALNTSC_in)
        m_blnFldTrueDriveEmu = varTrueDriveEmu_in
    Else
        'set default field variables for database fields
        With m_xobjRegistry
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "TDE", m_blnFldTrueDriveEmu, False, False
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "Control", m_lngFldControl, False, 0
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "PalNTSC", m_lngFldPALNTSC, False, 0
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "Players", m_lngFldNumPlayersTo, False, 1
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "Comment", m_strFldGameComment, True, ""
            .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\GEMUSFields", "VComment", m_strFldVersionComment, True, ""
        End With
        m_lngFldNumPlayersFrom = 1
    End If
    Call LogStatus("Will be using these game field values: ")
    Call LogStatus(" m_lngFldNumPlayersFrom [" & m_lngFldNumPlayersFrom & "]")
    Call LogStatus(" m_lngFldNumPlayersTo [" & m_lngFldNumPlayersTo & "]")
    Call LogStatus(" m_lngFldControl [" & m_lngFldControl & "]")
    Call LogStatus(" m_strFldGameComment [" & m_strFldGameComment & "]")
    Call LogStatus(" m_strFldVersionComment [" & m_strFldVersionComment & "]")
    Call LogStatus(" m_lngFldPALNTSC [" & m_lngFldNumPlayersFrom & "]")
    Call LogStatus(" m_blnFldTrueDriveEmu [" & m_lngFldNumPlayersFrom & "]")

    'Environment variables/fields
    m_strEnvGFFull = LCase$(strFileToRun)
    m_strFldGameType = UCase$(Right$(strFileToRun, Len(strFileToRun) - InStrRev(strFileToRun, ".")))
    m_lngFldNumGameFiles = 1
    With m_xobjFunc
        Call .GetFileAndPath(m_strEnvGFFull, m_strEnvGFPath, m_strEnvGFFile)
        m_strEnvGFPath = LCase$(m_strEnvGFPath)
        m_strEnvGFFile = LCase$(m_strEnvGFFile)
        strPathWBS = m_strEnvGFPath
        If Right$(strPathWBS, 1) <> "\" Then strPathWBS = strPathWBS & "\"
        m_strEnvGFNoExt = Left$(m_strEnvGFFile, InStrRev(m_strEnvGFFile, ".") - 1)
        strLastChar = UCase$(Right$(m_strEnvGFNoExt, 1))
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", strLastChar) > 0 Then
            'alphabetical extra filenames
            lngExtraIdx = Asc(strLastChar) + 1
            For lngFileIdx = 1 To 20
                ma_strExtraGFFull(lngFileIdx) = strPathWBS & Left$(m_strEnvGFNoExt, Len(m_strEnvGFNoExt) - 1) & LCase$(Chr$(lngExtraIdx)) & "." & LCase$(m_strFldGameType)
                If .FileExists(ma_strExtraGFFull(lngFileIdx), False, "") = True Then
                    m_lngFldNumGameFiles = m_lngFldNumGameFiles + 1
                    lngExtraIdx = lngExtraIdx + 1
                    Call .GetFileAndPath(ma_strExtraGFFull(lngFileIdx), ma_strExtraGFPath(lngFileIdx), ma_strExtraGFFile(lngFileIdx))
                    m_strExtraGFNoExt(lngFileIdx) = Left$(ma_strExtraGFFile(lngFileIdx), InStrRev(ma_strExtraGFFile(lngFileIdx), ".") - 1)
                Else
                    Exit For
                End If
            Next lngFileIdx
        ElseIf InStr("0123456789", strLastChar) > 0 Then
            'numerical extra filenames
            lngExtraIdx = Val(strLastChar) + 1
            For lngFileIdx = 1 To 20
                ma_strExtraGFFull(lngFileIdx) = strPathWBS & Left$(m_strEnvGFNoExt, Len(m_strEnvGFNoExt) - 1) & Trim$(lngExtraIdx) & "." & LCase$(m_strFldGameType)
                If .FileExists(ma_strExtraGFFull(lngFileIdx), False, "") = True Then
                    m_lngFldNumGameFiles = m_lngFldNumGameFiles + 1
                    lngExtraIdx = lngExtraIdx + 1
                    Call .GetFileAndPath(ma_strExtraGFFull(lngFileIdx), ma_strExtraGFPath(lngFileIdx), ma_strExtraGFFile(lngFileIdx))
                    m_strExtraGFNoExt(lngFileIdx) = Left$(ma_strExtraGFFile(lngFileIdx), InStrRev(ma_strExtraGFFile(lngFileIdx), ".") - 1)
                Else
                    Exit For
                End If
            Next lngFileIdx
        Else
            'no extra files - just set to first file
            For lngFileIdx = 1 To 20
                ma_strExtraGFFull(lngFileIdx) = m_strEnvGFFull
                ma_strExtraGFPath(lngFileIdx) = m_strEnvGFPath
                ma_strExtraGFFile(lngFileIdx) = m_strEnvGFFile
                m_strExtraGFNoExt(lngFileIdx) = m_strEnvGFNoExt
            Next lngFileIdx
        End If
    End With
    Call LogStatus("All script environment vars created")

    '-------------------------------------------------
    'replace environment vars in gemus key=value pairs
    '-------------------------------------------------

    Call LogStatus("Final k=v pairs are:")
    For lngIdx = 0 To m_lngTotalNotesKeys - 1
        Call ReplaceEnvironmentVars(ma_strNotesKeys(lngIdx).strValue, False, False)
        Call LogStatus(" " & ma_strNotesKeys(lngIdx).strKey & "=" & ma_strNotesKeys(lngIdx).strValue)
    Next lngIdx

    'quit if simulating
    If m_blnSimulateRun = True Then
        m_strSimulateKV = ""
        For lngIdx = 0 To m_lngTotalNotesKeys - 1
            With ma_strNotesKeys(lngIdx)
                m_strSimulateKV = m_strSimulateKV & .strKey & "=" & .strValue & vbCrLf
            End With
        Next lngIdx
        Set Unpacker = Nothing
        RunEmulator = ""
        m_blnSimulateRun = False
        Call LogStatus("Simulating run. Aborting.")
        Call CloseLogFile
        Exit Function
    End If

    '-------------------------
    'check the emulator exists
    '-------------------------

    If (m_lngEmuIdx >= 0) And (m_lngEmuIdx < m_lngTotalEmulators) Then
        If m_xobjFunc.FileExists(ma_udtEmulators(m_lngEmuIdx).strFullPathAndFilename, False, "") = True Then

            Call LogStatus("Emulator executable found: [" & ma_udtEmulators(m_lngEmuIdx).strFullPathAndFilename & "]")

            '---------------------------------------------------------------
            'Gamefile and Emulator Exist! - Validate and process the script!
            '---------------------------------------------------------------

            'get and validate the script lines
            strScript = ma_udtEmulators(m_lngEmuIdx).strScript
            If strScript <> "" Then
                If m_xobjFunc.FileExists(m_strAppPathWithBackSlash & m_strGBName & "\Scripts\" & strScript, False, "GEMUS") = True Then
                    f = FreeFile()
                    Open (m_strAppPathWithBackSlash & m_strGBName & "\Scripts\" & strScript) For Input As #f
                    strScript = Input$(LOF(f), f)
                    Close #f
                    Call LogStatus("Loaded GEMUS Script [" & m_strAppPathWithBackSlash & m_strGBName & "\Scripts\" & ma_udtEmulators(m_lngEmuIdx).strScript & "]")
                    strError = ValidateAndGetScriptLines(strScript)
                    If strError = "" Then
                        Call LogStatus("Validated Script lines")
                        'validate the script structure
                        strError = ValidateScriptStructure(strScript, lngErrorLine)
                        If strError = "" Then
                            Call LogStatus("Validated Script structure")
                            '--------------------------------------------------
                            'Initial script tests passed - Get lines to process
                            '--------------------------------------------------
    
                            'evaluate all expressions and get list of functions to process
                            m_lngTotalScriptLinesToProcess = 0
                            ReDim ma_strScriptLinesToProcess(0)
                            Call GetScriptFunctionsToProcess(0, lngNextLine, 0, True)
                            If m_lngTotalScriptLinesToProcess > 0 Then
                                Call LogStatus("Processing " & m_lngTotalScriptLinesToProcess & " script functions...")
                                
                                '------------------------
                                'Process the script lines
                                '------------------------
    
                                m_strCommandLine = ""
                                For lngLineIdx = 0 To (m_lngTotalScriptLinesToProcess - 1)
                                    Call LogStatus((lngLineIdx + 1) & " " & ma_strScriptLinesToProcess(lngLineIdx))
                                    strError = ProcessScriptLine(ma_strScriptLinesToProcess(lngLineIdx))
                                    If strError <> "" Then
                                        'error processing or executing a function
                                        RunEmulator = strError
                                        Call LogStatus("Error returned: [" & RunEmulator & "]")
                                        Exit For
                                    End If
                                Next lngLineIdx

                                If strError = "" Then
                                    'do a repack if required
                                    If Not m_blnSilentMode Then
                                        blnRepackAsk = m_blnAskUserBeforeRepack
                                        blnRepackNotify = m_blnNotifyUserIfRepackWasOK
                                    Else
                                        blnRepackAsk = False
                                        blnRepackNotify = False
                                    End If
                                    If m_blnAttemptRepack Then
                                        Call LogStatus("Attempting repack...")
                                        Screen.MousePointer = vbHourglass
                                        DoEvents
                                        If Unpacker.Repack( _
                                            blnRepackAsk, _
                                            blnRepackNotify, _
                                            strDestFolderWBS, _
                                            strGamePathFile_in, _
                                            me_RepackAddtional, _
                                            blnWasRepacked, _
                                            strError _
                                        ) Then
                                            RunEmulator = ""
                                        Else
                                            RunEmulator = strError
                                            Call LogStatus("Error returned: [" & RunEmulator & "]")
                                        End If
                                        Screen.MousePointer = vbDefault
                                        DoEvents
                                    End If
                                End If

                            Else
                                'no script lines to process!
                                RunEmulator = GetINIMessage("Running Games", "NoFunctions", "The script for this game has no functions to process!")
                                Call LogStatus("Error returned: [" & RunEmulator & "]")
                            End If
                        Else
                            'invalid script structure
                            If lngErrorLine > -1 Then
                                RunEmulator = m_strLang_ScriptError & vbCrLf & vbCrLf & strError & vbCrLf & vbCrLf & m_strLang_Line & " " & lngErrorLine
                            Else
                                RunEmulator = m_strLang_ScriptError & vbCrLf & vbCrLf & strError
                            End If
                            Call LogStatus("Error returned: [" & RunEmulator & "]")
                        End If
                    ElseIf m_lngTotalScriptLines = 0 Then
                        'no script
                        RunEmulator = GetINIMessage("Running Games", "NoScript", "There is no script set for this emulator!")
                        Call LogStatus("Error returned: [" & RunEmulator & "]")
                    Else
                        'bad line(s) found in script
                        RunEmulator = m_strLang_ScriptError & vbCrLf & vbCrLf & strError
                        Call LogStatus("Error returned: [" & RunEmulator & "]")
                    End If
                Else
                    'no script file
                    RunEmulator = GetINIMessage("Running Games", "NoScript", "There is no script set for this emulator!")
                    Call LogStatus("Error returned: [" & RunEmulator & "]")
                End If
            Else
                'no script
                RunEmulator = GetINIMessage("Running Games", "NoScript", "There is no script set for this emulator!")
                Call LogStatus("Error returned: [" & RunEmulator & "]")
            End If
        Else
            'emulator doesn't exist
            RunEmulator = GetINIMessage("Running Games", "EmulatorNotFound", "The emulator file could not be found.") & vbCrLf & vbCrLf & "'" & ma_udtEmulators(m_lngEmuIdx).strFullPathAndFilename & "'"
            Call LogStatus("Error returned: [" & RunEmulator & "]")
        End If
    Else
        'something very bad - invalid emulator index
        RunEmulator = GetINIMessage("Running Games", "InvalidEmulator", "Invalid Emulator!")
        Call LogStatus("Error returned: [" & RunEmulator & "]")
    End If
    
    Set Unpacker = Nothing
    Call CloseLogFile
    Exit Function

Err_Handler:
    RunEmulator = "Unable to run the Emulator at this time: " & Err.Description
    Call LogStatus("Error returned: [" & RunEmulator & "]")
    Set Unpacker = Nothing
    Call CloseLogFile

End Function

Friend Property Get ScriptLines(ByVal Index As Long) As String

    ScriptLines = ma_strScriptLines(Index)

End Property
Friend Property Get ScriptLinesPre(ByVal Index As Long) As String

    ScriptLinesPre = ma_strScriptLinesPre(Index)

End Property
Private Sub SetEmuFileTypes(strFileTypes_in As String)

    Dim a_strTypes() As String
    Dim lngFTIdx As Long

    m_lngTotalEmulatorFileTypes = 0
    ReDim ma_strEmulatorFileTypes(0)

    If strFileTypes_in <> "" Then
        a_strTypes = Split(strFileTypes_in, ";")
        For lngFTIdx = 0 To UBound(a_strTypes)
            a_strTypes(lngFTIdx) = Trim$(a_strTypes(lngFTIdx))
            If ValidateEmuFileType(a_strTypes(lngFTIdx)) = True Then
                ReDim Preserve ma_strEmulatorFileTypes(m_lngTotalEmulatorFileTypes)
                ma_strEmulatorFileTypes(m_lngTotalEmulatorFileTypes) = UCase$(a_strTypes(lngFTIdx))
                m_lngTotalEmulatorFileTypes = m_lngTotalEmulatorFileTypes + 1
            End If
        Next lngFTIdx
    End If

End Sub
Public Sub SetRezip(ByVal blnDoRezip_in As Boolean, ByVal blnUserConfirm_in As Boolean, e_AdditionalFilesRule As EnumRezipAdditionalFiles)

    m_blnUserWantsToRepack = blnDoRezip_in
    m_blnAskUserBeforeRepack = blnUserConfirm_in
    me_RepackAddtional = e_AdditionalFilesRule

End Sub

Public Sub ShowFileTypeDialog()

    Dim strFileTypes As String

    g_strFileTypesGBName = m_strGBName
    g_strFileTypesEmuFile = m_strEmuFile
    frmFileTypes.Show vbModal
    If g_blnFileTypesChanged = True Then
        strFileTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", "", m_strEmuFile)
        Call SetEmuFileTypes(strFileTypes)
    End If

End Sub

Public Function ShowWizard() As Boolean

    Dim blnGlobalDAT As Boolean
    Dim strDATFile As String
    Dim intFileNum As Integer
    Dim strError As String

    'get the DAT file to use
    If m_xobjFunc.FileExists(m_strAppPathWithBackSlash & m_strGBName & "\" & "gemus.dat", False, "") = True Then
        blnGlobalDAT = False
        strDATFile = m_strAppPathWithBackSlash & m_strGBName & "\" & "gemus.dat"
    ElseIf m_xobjFunc.FileExists(m_strAppPathWithBackSlash & "gemus.dat", False, "") = True Then
        blnGlobalDAT = True
        strDATFile = m_strAppPathWithBackSlash & "gemus.dat"
    Else
        'no dat file
        MsgBox GetINIMessage("Wizard", "Error_NoDATFile", "The Wizard could not find GEMUS.DAT in your GameBase folder! You may need to download it from the GameBase home page!"), vbExclamation, m_strAppName
        ShowWizard = False
        Exit Function
    End If
    
    'read the file into memory
    With g_udtWizard
        'init the dat file
        .blnAborted = True
        .lngTotalSections = 0
        .strDATFile = strDATFile
        .blnGlobalDAT = blnGlobalDAT
        .strMessagesFile = m_strLangINIPathFile
        .strGBFolder = m_strAppPathWithBackSlash & m_strGBName
        If blnGlobalDAT = True Then
            .strScriptFolder = m_strAppPathWithBackSlash & "Scripts"
        Else
            .strScriptFolder = m_strAppPathWithBackSlash & m_strGBName & "\Scripts"
        End If
        .strEmuFile = m_strEmuFile
        ReDim .a_udtSections(0)
        'get the header info
        .sngDATVersion = Val(m_xobjFunc.INI_GetSetting("DAT", "Version", "0", strDATFile))
        If .sngDATVersion <> 1# Then
            MsgBox GetINIMessage("Wizard", "Error_WrongDATVersion", "Wrong DAT file version. Expected") & ": " & "1.0", vbExclamation, m_strAppName
            ShowWizard = False
            Exit Function
        End If
        .strLastUpdated = m_xobjFunc.INI_GetSetting("CREDITS", "LastUpdated", "", strDATFile)
        .strAuthor = m_xobjFunc.INI_GetSetting("CREDITS", "Author", "", strDATFile)
        .strHomePage = m_xobjFunc.INI_GetSetting("CREDITS", "Contact", "http://www.bu22.com/", strDATFile)
        'now get the sections (systems) and emulators
        intFileNum = FreeFile()
        Open strDATFile For Input As #intFileNum
        strError = ReadInWizardDATFile(strDATFile, intFileNum)
        Close #intFileNum
        If strError <> "" Then
            MsgBox strError, vbExclamation, m_strAppName
            ShowWizard = False
            Exit Function
        End If
        'show the wizard
        frmWizard.Show vbModal
        'what happened?
        If .blnAborted = False Then
            'read in the emulators again!
            Call ReadInEmulators
        End If
    End With

    'all ok - signify client to update its GEMUS menu
    ShowWizard = (g_udtWizard.blnAborted = False)

End Function

Friend Property Get TotalScriptLines() As Long

    TotalScriptLines = m_lngTotalScriptLines

End Property

Public Function ValidateAndGetScriptLines(strScript_in As String) As String

    Dim strLine As String
    Dim lngLineIdx As Long

    ma_strScriptLines() = Split(strScript_in, vbCrLf)
    m_lngTotalScriptLines = UBound(ma_strScriptLines) + 1
    ReDim ma_strScriptLinesPre(m_lngTotalScriptLines - 1)
    If m_lngTotalScriptLines > 0 Then
        For lngLineIdx = 0 To (m_lngTotalScriptLines - 1)
            ValidateAndGetScriptLines = ValidateScriptLine(ma_strScriptLines(lngLineIdx))
            If ValidateAndGetScriptLines <> "" Then
                 ValidateAndGetScriptLines = ValidateAndGetScriptLines & vbCrLf & vbCrLf & m_strLang_Line & " " & (lngLineIdx + 1)
                Exit Function
            End If
        Next lngLineIdx
    End If

End Function
Private Sub LoadScriptData()

    '-----------------------------------------------------
    'Loads the script data for the GEMUS(tm) script engine
    'GEMUS - Gamebase Emulator Script
    '-----------------------------------------------------

    Dim lngIdx As Long      'loop index
    Dim lngVar As Long      'environment var position

    '========================
    'script comparison fields
    '========================

    m_lngMaxScriptFields = 11
    ReDim ma_udtScriptFields(m_lngMaxScriptFields)

    'GAMETYPE
    '=,!=
    'T64,D64,G64,G41,TAP,X64,C64,P00,PRG,CRT,*any*
    With ma_udtScriptFields(0)
        .strName = "GameType"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 2
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With
    
    'NUMPLAYERS
    '=,!=,<,>
    '*SINGLE NUMBER*
    With ma_udtScriptFields(1)
        .strName = "NumPlayers"
        .eValueType = e_ScriptValType_Number
        .lngTotalOperators = 4
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .strOperators(2) = "<"
        .strOperators(3) = ">"
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'CONTROL
    '=,!=
    'JoyPort1,JoyPort2,Keyboard,PaddlePort1,PaddlePort2,Mouse,LightPen,KoalaPad,LightGun
    With ma_udtScriptFields(2)
        .strName = "Control"
        .eValueType = e_ScriptValType_List
        .lngTotalOperators = 2
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .lngTotalListValues = 9
        ReDim .strListValues(.lngTotalListValues - 1)
        .strListValues(0) = "JoyPort2"
        .strListValues(1) = "JoyPort1"
        .strListValues(2) = "Keyboard"
        .strListValues(3) = "PaddlePort2"
        .strListValues(4) = "PaddlePort1"
        .strListValues(5) = "Mouse"
        .strListValues(6) = "LightPen"
        .strListValues(7) = "KoalaPad"
        .strListValues(8) = "LightGun"
    End With

    'GAMECOMMENT
    'CONTAINS
    '(*any text*)
    With ma_udtScriptFields(3)
        .strName = "GameComment"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = ""
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'PALNTSC
    '=,!=
    'PAL,NTSC,BOTH
    With ma_udtScriptFields(4)
        .strName = "PalNTSC"
        .eValueType = e_ScriptValType_List
        .lngTotalOperators = 2
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .lngTotalListValues = 3
        ReDim .strListValues(.lngTotalListValues - 1)
        .strListValues(0) = "PAL"
        .strListValues(1) = "NTSC"
        .strListValues(2) = "BOTH"
    End With

    'TRUEDRIVEEMU
    '=
    'YES,NO
    With ma_udtScriptFields(5)
        .strName = "TrueDriveEmu"
        .eValueType = e_ScriptValType_List
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .lngTotalListValues = 2
        ReDim .strListValues(.lngTotalListValues - 1)
        .strListValues(0) = "YES"
        .strListValues(1) = "NO"
    End With

    'VERSIONCOMMENT
    'CONTAINS
    '(*any text*)
    With ma_udtScriptFields(6)
        .strName = "VersionComment"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = ""
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'NUMGAMEFILES
    '=,!=,<,>
    '*SINGLE NUMBER*
    With ma_udtScriptFields(7)
        .strName = "NumGameFiles"
        .eValueType = e_ScriptValType_Number
        .lngTotalOperators = 4
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .strOperators(2) = "<"
        .strOperators(3) = ">"
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'IMAGEINDEX
    '=,!=,<,>
    '*SINGLE NUMBER*
    With ma_udtScriptFields(8)
        .strName = "ImageIndex"
        .eValueType = e_ScriptValType_Number
        .lngTotalOperators = 4
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = "="
        .strOperators(1) = "!="
        .strOperators(2) = "<"
        .strOperators(3) = ">"
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'IMAGENAME
    'CONTAINS
    '(*any text*)
    With ma_udtScriptFields(9)
        .strName = "ImageName"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = ""
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'KEY_
    'CONTAINS
    '(*any text*)
    With ma_udtScriptFields(10)
        .strName = "Key_"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = ""
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    '%script_env_var%
    'CONTAINS
    '(*any text*)
    With ma_udtScriptFields(11)
        .strName = "%"
        .eValueType = e_ScriptValType_Text
        .lngTotalOperators = 1
        ReDim .strOperators(.lngTotalOperators - 1)
        .strOperators(0) = ""
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    '================
    'script functions
    '================

    m_lngMaxScriptFunctions = 14
    ReDim ma_udtScriptFunctions(m_lngMaxScriptFunctions)

    'ADD_CLP2(*anytext*)
    With ma_udtScriptFunctions(0)
        .strName = "Add_CLP2"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With
    
    'ADD_CLP(*anytext*)
    With ma_udtScriptFunctions(1)
        .strName = "Add_CLP"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'SET_INI_VALUE(filenum||section||key||value)
    With ma_udtScriptFunctions(2)
        .strName = "Set_INI_Value"
        .lngTotalParams = 4
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_FileNumber
        .eParamType(1) = e_ScriptFuncParamType_INISection
        .eParamType(2) = e_ScriptFuncParamType_Text
        .eParamType(3) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'SET_CFG_ITEM(filenum||key||delim||value)
    With ma_udtScriptFunctions(3)
        .strName = "Set_CFG_Item"
        .lngTotalParams = 4
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_FileNumber
        .eParamType(1) = e_ScriptFuncParamType_Text
        .eParamType(2) = e_ScriptFuncParamType_Text
        .eParamType(3) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With
    
    'SET_CFG_VALUE(filenum||key||value)
    With ma_udtScriptFunctions(4)
        .strName = "Set_CFG_Value"
        .lngTotalParams = 3
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_FileNumber
        .eParamType(1) = e_ScriptFuncParamType_Text
        .eParamType(2) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'RUN_PROGRAM(filenum||clp||wait)
    With ma_udtScriptFunctions(5)
        .strName = "Run_Program"
        .lngTotalParams = 3
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_FileNumber
        .eParamType(1) = e_ScriptFuncParamType_Text
        .eParamType(2) = e_ScriptFuncParamType_List
        .lngTotalListValues = 2
        ReDim .strListValues(.lngTotalListValues - 1)
        .strListValues(0) = "WAIT"
        .strListValues(1) = "NOWAIT"
    End With

    'RUN_EMULATOR()
    With ma_udtScriptFunctions(6)
        .strName = "Run_Emulator"
        .lngTotalParams = 0
        ReDim .eParamType(0)
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'SHOW_MESSAGE(GAMENOTSUPPORTED|*custom message*)
    With ma_udtScriptFunctions(7)
        .strName = "Show_Message"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'EDIT_CLP()
    With ma_udtScriptFunctions(8)
        .strName = "Edit_CLP"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'RUN_GAMEFILE()
    With ma_udtScriptFunctions(9)
        .strName = "Run_GameFile"
        .lngTotalParams = 0
        ReDim .eParamType(0)
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'CLEAR_CLP()
    With ma_udtScriptFunctions(10)
        .strName = "Clear_CLP"
        .lngTotalParams = 0
        ReDim .eParamType(0)
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'RUN_EMULATOR_SEND_KEYS()
    With ma_udtScriptFunctions(11)
        .strName = "Run_Emulator_Send_Keys"
        .lngTotalParams = 2
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .eParamType(1) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'SHOW_KV()
    With ma_udtScriptFunctions(12)
        .strName = "Show_KV"
        .lngTotalParams = 0
        ReDim .eParamType(0)
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'WAIT(ms)
    With ma_udtScriptFunctions(13)
        .strName = "Wait"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    'KILL_PROCESS(filename)
    With ma_udtScriptFunctions(14)
        .strName = "Kill_Process"
        .lngTotalParams = 1
        ReDim .eParamType(.lngTotalParams - 1)
        .eParamType(0) = e_ScriptFuncParamType_Text
        .lngTotalListValues = 0
        ReDim .strListValues(0)
    End With

    '============================
    'script environment variables
    '============================

    m_lngMaxScriptEnvVars = 85
    ReDim ma_strScriptEnvVars(m_lngMaxScriptEnvVars)

    ma_strScriptEnvVars(0) = "%gamepathfile%"
    ma_strScriptEnvVars(1) = "%gamefile%"
    ma_strScriptEnvVars(2) = "%gamepath%"
    ma_strScriptEnvVars(3) = "%gamefilenoext%"
    ma_strScriptEnvVars(4) = "%imageindex%"
    ma_strScriptEnvVars(5) = "%imagename%"
    lngVar = 6
    For lngIdx = 1 To 20    'e.g. %gamepathfile(1)%
        ma_strScriptEnvVars(lngVar) = "%gamepathfile(" & lngIdx & ")%"
        ma_strScriptEnvVars(lngVar + 1) = "%gamefile(" & lngIdx & ")%"
        ma_strScriptEnvVars(lngVar + 2) = "%gamepath(" & lngIdx & ")%"
        ma_strScriptEnvVars(lngVar + 3) = "%gamefilenoext(" & lngIdx & ")%"
        lngVar = lngVar + 4
    Next lngIdx

End Sub
Private Sub GetScriptLineParts(strLine_in As String, strPre_out As String, strLine_out As String)

    '------------------------------------------------------------------
    'splits a script line into 2 parts
    '1. The spaces/tabs leading up to the first valid character
    '2. The characters in the line (excluding pre/post spaces and tabs)
    '
    'Inputs:    strLine_in:     The line to split
    '
    'Outputs:   strPre_out:     The string before the actual chars
    '           strLine_out:    The chars after the pre-spaces/tabs
    '
    'Remarks:   strLine_out will also have tabs replaced with spaces
    '------------------------------------------------------------------

    Dim strChar As String       'char we are testing
    Dim lngCharIdx As Long      'loop index

    'clear return strings
    strPre_out = ""
    strLine_out = ""

    'find first non space/tab character
    strPre_out = ""
    strLine_out = ""
    For lngCharIdx = 1 To Len(strLine_in)
        strChar = Mid$(strLine_in, lngCharIdx, 1)
        If (strChar = vbTab) Or (strChar = " ") Then
            strPre_out = strPre_out & strChar
        Else
            strLine_out = strChar
            Exit For
        End If
    Next lngCharIdx

    'get the actual script line
    If strLine_out <> "" Then
        strLine_out = Right$(strLine_in, Len(strLine_in) - Len(strPre_out))
        'trim the tabs and spaces from the end of the line
        For lngCharIdx = Len(strLine_out) To 1 Step -1
            strChar = Mid$(strLine_out, lngCharIdx, 1)
            If (strChar = vbTab) Or (strChar = " ") Then
                strLine_out = Left$(strLine_out, lngCharIdx - 1)
            Else
                Exit For
            End If
        Next lngCharIdx
        'replace tabs with spaces
        strLine_out = Replace(strLine_out, vbTab, " ")
    End If

End Sub
Private Function ValidateEmuFileType(strFT_inout As String) As Boolean

    If strFT_inout <> "" Then
        If Left$(strFT_inout, 1) = "." Then
            'strip dot
            If Len(strFT_inout) > 1 Then
                strFT_inout = Right$(strFT_inout, Len(strFT_inout) - 1)
                ValidateEmuFileType = True
                Exit Function
            End If
        Else
            'no dot
            ValidateEmuFileType = True
            Exit Function
        End If
    End If
    ValidateEmuFileType = False

End Function

Private Function ValidateFunctionLine(strLine_inout As String) As String

    '-----------------------------------------------------------------------
    'validates the function line passed in and formats it properly
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strUpper As String      'uppercase function line
    Dim strParams() As String   'the function parameters
    Dim lngNumParams As Long    'number of function parameters
    Dim strTemp As String       'for function calls
    Dim strTemp2 As String      'for function calls
    Dim strTemp3 As String      'for function calls
    Dim lngTemp As Long         'for function calls
    Dim blnTemp As Boolean      'for function calls
    Dim strFilename As String   'for function calls

    'validate the function that matches the line
    strUpper = UCase$(strLine_inout)
    If Left$(strUpper, Len(ma_udtScriptFunctions(0).strName)) = "ADD_CLP2" Then
        'ADD_CLP2()
        ValidateFunctionLine = Validate_ADD_CLP2(strLine_inout, strTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(1).strName)) = "ADD_CLP" Then
        'ADD_CLP()
        ValidateFunctionLine = Validate_ADD_CLP(strLine_inout, strTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(2).strName)) = "SET_INI_VALUE" Then
        'SET_INI_VALUE
        ValidateFunctionLine = Validate_SET_INI_VALUE(strLine_inout, lngTemp, strTemp, strTemp2, strTemp3, strFilename)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(3).strName)) = "SET_CFG_ITEM" Then
        'SET_CFG_ITEM
        ValidateFunctionLine = Validate_SET_CFG_ITEM(strLine_inout, lngTemp, strTemp, strTemp2, strTemp3, strFilename)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(4).strName)) = "SET_CFG_VALUE" Then
        'SET_CFG_VALUE
        ValidateFunctionLine = Validate_SET_CFG_VALUE(strLine_inout, lngTemp, strTemp, strTemp2, strFilename)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(5).strName)) = "RUN_PROGRAM" Then
        'RUN_PROGRAM()
        ValidateFunctionLine = Validate_RUN_PROGRAM(strLine_inout, lngTemp, strTemp, blnTemp, strFilename)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(11).strName)) = "RUN_EMULATOR_SEND_KEYS" Then
        'RUN_EMULATOR()
        ValidateFunctionLine = Validate_RUN_EMULATOR_SEND_KEYS(strLine_inout, strTemp, strTemp2, blnTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(6).strName)) = "RUN_EMULATOR" Then
        'RUN_EMULATOR()
        ValidateFunctionLine = Validate_RUN_EMULATOR(strLine_inout, blnTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(7).strName)) = "SHOW_MESSAGE" Then
        'SHOW_MESSAGE()
        ValidateFunctionLine = Validate_SHOW_MESSAGE(strLine_inout, strTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(8).strName)) = "EDIT_CLP" Then
        'EDIT_CLP()
        ValidateFunctionLine = Validate_EDIT_CLP(strLine_inout, strTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(9).strName)) = "RUN_GAMEFILE" Then
        'RUN_GAMEFILE()
        ValidateFunctionLine = Validate_RUN_GAMEFILE(strLine_inout, blnTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(10).strName)) = "CLEAR_CLP" Then
        'CLEAR_CLP()
        ValidateFunctionLine = Validate_CLEAR_CLP(strLine_inout)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(12).strName)) = "SHOW_KV" Then
        'SHOW_KV()
        ValidateFunctionLine = Validate_SHOW_KV(strLine_inout)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(13).strName)) = "WAIT" Then
        'WAIT()
        ValidateFunctionLine = Validate_WAIT(strLine_inout, strTemp)
    ElseIf Left$(strUpper, Len(ma_udtScriptFunctions(14).strName)) = "KILL_PROCESS" Then
        'WAIT()
        ValidateFunctionLine = Validate_KILL_PROCESS(strLine_inout, strTemp)
    Else
        'not a function
        ValidateFunctionLine = m_strLang_UnknownCommand
    End If

End Function

Private Function Validate_RUN_EMULATOR(strLine_inout As String, blnDOSClose_out As Boolean) As String

    '-----------------------------------------------------------------------
    'validates the RUN_EMULATOR function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Outputs:   blnDOSClose_out:    'Close DOS window afterwards' flag
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'function arguments array
    Dim strArgs2() As String    'function arguments array

    'we already know the first chars in the line are RUN_EMULATOR and the line is trimmed
    
    'check if there are no args
    lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs(), strError, 0)
    If strError <> "" Then
        'check if there is one arg
        lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs2(), strError, 1)
        If strError <> "" Then
            'error wrong no. args
            Validate_RUN_EMULATOR = strError
        Else
            'check arg
            If UCase$(strArgs2(0)) = "DOSCLOSE" Then
                'everything is fine!  return the proper string
                Validate_RUN_EMULATOR = ""
                blnDOSClose_out = True
                strLine_inout = "Run_Emulator(DOSCLOSE)"
            Else
                'invalid 2nd param
                Validate_RUN_EMULATOR = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": DOSCLOSE"
            End If
        End If
    Else
        'everything is fine!  return the proper string
        Validate_RUN_EMULATOR = ""
        blnDOSClose_out = False
        strLine_inout = "Run_Emulator()"
    End If

End Function
Private Function Validate_RUN_EMULATOR_SEND_KEYS(strLine_inout As String, strSendKeys_out As String, strKeyDownDuration_out As String, blnDOSClose_out As Boolean) As String

    '-----------------------------------------------------------------------
    'validates the RUN_EMULATOR function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Outputs:   lngSendKeysWait_out:    'num seconds to wait before sending keys
    '           strSendKeys_out:        'keys to send
    '           blnDOSClose_out:        'Close DOS window afterwards' flag
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'function arguments array
    Dim strArg1 As String
    Dim strArg2 As String
    Dim strArg3 As String

    'we already know the first chars in the line are RUN_EMULATOR_SEND_KEYS and the line is trimmed
    
    'check if there are 2 args (seconds to wait, keys to send)
    lngNumArgs = GetScriptFunctionArgs(22, strLine_inout, strArgs(), strError, 2)
    If strError <> "" Then
        'check if there are 3 args (keys to send, keydown duration, dosclose)
        lngNumArgs = GetScriptFunctionArgs(22, strLine_inout, strArgs(), strError, 3)
    End If
    If strError <> "" Then
            'error wrong no. args
            Validate_RUN_EMULATOR_SEND_KEYS = strError
    Else
        '2 or 3 args?
        strArg1 = strArgs(0)
        strArg2 = strArgs(1)
        strArg3 = ""
        If lngNumArgs = 3 Then strArg3 = strArgs(2)
        
        'grab/validate args
        strSendKeys_out = strArg1
        If strSendKeys_out = "" Then
            Validate_RUN_EMULATOR_SEND_KEYS = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": (KEYS TO SEND)"
            Exit Function
        End If

        strKeyDownDuration_out = Trim$(strArg2)
        If strKeyDownDuration_out = "" Then
            Validate_RUN_EMULATOR_SEND_KEYS = m_strLang_InvalidFunctionParam & " 2.  " & m_strLang_Expected & ": (KEYDOWN/UP DURATION MS)"
            Exit Function
        End If

        If UCase$(strArg3) = "DOSCLOSE" Then
            Validate_RUN_EMULATOR_SEND_KEYS = ""
            blnDOSClose_out = True
            strLine_inout = "Run_Emulator_Send_Keys(" & strArg1 & "||" & strKeyDownDuration_out & "||" & UCase$(strArg3) & ")"
        ElseIf strArg3 = "" Then
            Validate_RUN_EMULATOR_SEND_KEYS = ""
            blnDOSClose_out = False
            strLine_inout = "Run_Emulator_Send_Keys(" & strArg1 & "||" & strKeyDownDuration_out & ")"
        Else
            'invalid 3nd param
            Validate_RUN_EMULATOR_SEND_KEYS = m_strLang_InvalidFunctionParam & " 3.  " & m_strLang_Expected & ": DOSCLOSE"
        End If
        
    End If

End Function
Private Function Validate_RUN_GAMEFILE(strLine_inout As String, blnDOSClose_out As Boolean) As String

    '-----------------------------------------------------------------------
    'validates the RUN_GAMEFILE function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Outputs:   blnDOSClose_out:    'Close DOS window afterwards' flag
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'function arguments array
    Dim strArgs2() As String    'function arguments array

    'we already know the first chars in the line are RUN_EMULATOR and the line is trimmed
    
    'check if there are no args
    lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs(), strError, 0)
    If strError <> "" Then
        'check if there is one arg
        lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs2(), strError, 1)
        If strError <> "" Then
            'error wrong no. args
            Validate_RUN_GAMEFILE = strError
        Else
            'check arg
            If UCase$(strArgs2(0)) = "DOSCLOSE" Then
                'everything is fine!  return the proper string
                Validate_RUN_GAMEFILE = ""
                blnDOSClose_out = True
                strLine_inout = "Run_GameFile(DOSCLOSE)"
            Else
                'invalid 2nd param
                Validate_RUN_GAMEFILE = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": DOSCLOSE"
            End If
        End If
    Else
        'everything is fine!  return the proper string
        Validate_RUN_GAMEFILE = ""
        blnDOSClose_out = False
        strLine_inout = "Run_GameFile()"
    End If

End Function

Private Function Validate_CLEAR_CLP(strLine_inout As String) As String

    '-----------------------------------------------------------------------
    'validates the CLEAR_CLP function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)

    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are SHOW_MESSAGE and the line is trimmed

    'make sure there is 0 arg
    lngNumArgs = GetScriptFunctionArgs(9, strLine_inout, strArgs(), strError, 0)
    If strError <> "" Then
        Validate_CLEAR_CLP = strError
    Else
        'everything is fine!  return the proper string
        Validate_CLEAR_CLP = ""
        strLine_inout = "Clear_CLP()"
    End If

End Function

Private Function Validate_SHOW_MESSAGE(strLine_inout As String, strMessage_out As String) As String

    '-----------------------------------------------------------------------
    'validates the SHOW_MESSAGE function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    strMessage_out:     The message to display to the user!
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are SHOW_MESSAGE and the line is trimmed

    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_SHOW_MESSAGE = strError
    Else
        'everything is fine!  return the proper string
        Validate_SHOW_MESSAGE = ""
        strMessage_out = Replace(strArgs(0), "GAME_NOT_SUPPORTED", m_strLang_GameNotSupported & " (." & UCase$(m_strFldGameType) & ")")
        strLine_inout = "Show_Message(" & strArgs(0) & ")"
    End If

End Function

Private Function Validate_RUN_PROGRAM(strLine_inout As String, lngFileNum_out As Long, strCLP_out As String, blnWait_out As Boolean, strFilename_out As String) As String

    '------------------------------------------------------------------------------
    'validates the RUN_PROGRAM function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    lngFileNum_out:     Filenumber of program to run
    '           strCLP_out:         Command line parameters for program
    '           blnWait_out:        flag for whether to wait until program finishes
    '           strFilename_out:    empty string if using a file number, or the path\file to run
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '------------------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are RUN_PROGRAM and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(11, strLine_inout, strArgs(), strError, 3)
    If strError <> "" Then
        Validate_RUN_PROGRAM = strError
    Else
        'check params are ok
        strFilename_out = ""
        If ConvertStringToNumber(strArgs(0)) < 1 Then
            If strArgs(0) = "" Then
                'invalid file number
                Validate_RUN_PROGRAM = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": " & m_strLang_Number
                Exit Function
            Else
                strFilename_out = strArgs(0)
            End If
        End If
        If (UCase(strArgs(2)) = "WAIT") Or (UCase(strArgs(2)) = "NOWAIT") Then
            'everything is fine!  return the proper string
            Validate_RUN_PROGRAM = ""
            lngFileNum_out = Int(Val(strArgs(0)))
            strCLP_out = strArgs(1)
            blnWait_out = (UCase(strArgs(2)) = "WAIT")
            strLine_inout = "Run_Program(" & strArgs(0) & "||" & strArgs(1) & "||" & strArgs(2) & ")"
        Else
            'invalid 3rd wait parameter
            Validate_RUN_PROGRAM = m_strLang_InvalidFunctionParam & " 3.  " & m_strLang_Expected & ": [ WAIT | NOWAIT ]"
        End If
    End If

End Function
Private Function Validate_SET_CFG_VALUE(strLine_inout As String, lngFileNum_out As Long, strKey_out As String, strValue_out As String, strFilename_out As String) As String

    '------------------------------------------------------------------------------
    'validates the SET_CFG_VALUE function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    lngFileNum_out:     Filenumber of program to run
    '           strKey_out:         key to change
    '           strValue_out:       value to set
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '------------------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are SET_CFG_VALUE and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(13, strLine_inout, strArgs(), strError, 3)
    If strError <> "" Then
        Validate_SET_CFG_VALUE = strError
    Else
        'check params are ok
        
        strFilename_out = ""
        If ConvertStringToNumber(strArgs(0)) < 1 Then
            If strArgs(0) = "" Then
                'invalid file number
                Validate_SET_CFG_VALUE = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": " & m_strLang_Number
                Exit Function
            Else
                strFilename_out = strArgs(0)
            End If
        End If
        If strArgs(1) = "" Then
            'invalid 2nd KEY parameter
            Validate_SET_CFG_VALUE = m_strLang_InvalidFunctionParam & " 2.  " & m_strLang_Expected & ": KEY"
        Else
            'everything is fine!  return the proper string
            Validate_SET_CFG_VALUE = ""
            lngFileNum_out = Int(Val(strArgs(0)))
            strKey_out = strArgs(1)
            strValue_out = strArgs(2)
            strLine_inout = "Set_CFG_Value(" & strArgs(0) & "||" & strArgs(1) & "||" & strArgs(2) & ")"
        End If
    End If

End Function

Private Function Validate_SET_CFG_ITEM(strLine_inout As String, lngFileNum_out As Long, strKey_out As String, strDelim_out As String, strValue_out As String, strFilename_out As String) As String

    '------------------------------------------------------------------------------
    'validates the SET_CFG_VALUE function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    lngFileNum_out:     Filenumber of program to run
    '           strKey_out:         key to change
    '           strDelim_out:       the delimeter string
    '           strValue_out:       value to set
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '------------------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are SET_CFG_VALUE and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs(), strError, 4)
    If strError <> "" Then
        Validate_SET_CFG_ITEM = strError
    Else
        'check params are ok
        strFilename_out = ""
        If ConvertStringToNumber(strArgs(0)) < 1 Then
            If strArgs(0) = "" Then
                'invalid file number
                Validate_SET_CFG_ITEM = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": " & m_strLang_Number
                Exit Function
            Else
                strFilename_out = strArgs(0)
            End If
        End If
        If strArgs(1) = "" Then
            'invalid 2nd KEY parameter
            Validate_SET_CFG_ITEM = m_strLang_InvalidFunctionParam & " 2.  " & m_strLang_Expected & ": KEY"
        Else
            'everything is fine!  return the proper string
            Validate_SET_CFG_ITEM = ""
            lngFileNum_out = Int(Val(strArgs(0)))
            strKey_out = strArgs(1)
            strDelim_out = strArgs(2)
            strValue_out = strArgs(3)
            strLine_inout = "Set_CFG_Item(" & strArgs(0) & "||" & strArgs(1) & "||" & strArgs(2) & "||" & strArgs(3) & ")"
        End If
    End If

End Function

Private Function Validate_SET_INI_VALUE(strLine_inout As String, lngFileNum_out As Long, strSection_out As String, strKey_out As String, strValue_out As String, strFilename_out As String) As String

    '------------------------------------------------------------------------------
    'validates the SET_INI_VALUE function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    lngFileNum_out:     Filenumber of program to run
    '           strSection_out:     section in INI file
    '           strKey_out:         key to change
    '           strValue_out:       value to set
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '------------------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are SET_INI_VALUE and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(13, strLine_inout, strArgs(), strError, 4)
    If strError <> "" Then
        Validate_SET_INI_VALUE = strError
    Else
        'check params are ok
        strFilename_out = ""
        If ConvertStringToNumber(strArgs(0)) < 1 Then
            If strArgs(0) = "" Then
                'invalid file number
                Validate_SET_INI_VALUE = m_strLang_InvalidFunctionParam & " 1.  " & m_strLang_Expected & ": " & m_strLang_Number
                Exit Function
            Else
                strFilename_out = strArgs(0)
            End If
        End If
        If strArgs(1) = "" Then
            'invalid 2nd SECTION parameter
            Validate_SET_INI_VALUE = m_strLang_InvalidFunctionParam & " 2.  " & m_strLang_Expected & ": SECTION"
        ElseIf strArgs(2) = "" Then
            'invalid 3rd KEY parameter
            Validate_SET_INI_VALUE = m_strLang_InvalidFunctionParam & " 3.  " & m_strLang_Expected & ": KEY"
        Else
            'everything is fine!  return the proper string
            Validate_SET_INI_VALUE = ""
            lngFileNum_out = Int(Val(strArgs(0)))
            strSection_out = strArgs(1)
            strKey_out = strArgs(2)
            strValue_out = strArgs(3)
            strLine_inout = "Set_INI_Value(" & strArgs(0) & "||" & strArgs(1) & "||" & strArgs(2) & "||" & strArgs(3) & ")"
        End If
    End If

End Function
Private Function Validate_SHOW_KV(strLine_inout As String) As String

    '-----------------------------------------------------------------------
    'validates the SHOW_KV function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)

    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'make sure there is 0 arg
    lngNumArgs = GetScriptFunctionArgs(7, strLine_inout, strArgs(), strError, 0)
    If strError <> "" Then
        Validate_SHOW_KV = strError
    Else
        'everything is fine!  return the proper string
        Validate_SHOW_KV = ""
        strLine_inout = "Show_KV()"
    End If

End Function

Private Function Validate_WAIT(strLine_inout As String, strMillisecondsToWait_out As String) As String

    '-----------------------------------------------------------------------
    'validates the WAIT function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    strMillisecondsToWait_out:     The number of milliseconds to wait
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are WAIT and the line is trimmed

    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(4, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_WAIT = strError
    Else
        'everything is fine!  return the proper string
        Validate_WAIT = ""
        strMillisecondsToWait_out = strArgs(0)
        strLine_inout = "Wait(" & strArgs(0) & ")"
    End If

End Function

Private Function Validate_KILL_PROCESS(strLine_inout As String, strFilename_out As String) As String

    '-----------------------------------------------------------------------
    'validates the KILL_PROCESS function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    strFilename_out:     The filename to kill
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are KILL_PROCESS and the line is trimmed

    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(12, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_KILL_PROCESS = strError
    Else
        'everything is fine!  return the proper string
        Validate_KILL_PROCESS = ""
        strFilename_out = strArgs(0)
        strLine_inout = "Kill_Process(" & strArgs(0) & ")"
    End If

End Function

Private Function Validate_ADD_CLP(strLine_inout As String, strCLP_out As String) As String

    '-----------------------------------------------------------------------
    'validates the ADD_CLP function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    strCLP_out:         The clp string
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are ADD_CLP and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(7, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_ADD_CLP = strError
    Else
        'everything is fine!  return the proper string
        Validate_ADD_CLP = ""
        strCLP_out = strArgs(0)
        strLine_inout = "Add_CLP(" & strCLP_out & ")"
    End If

End Function

Private Function Validate_ADD_CLP2(strLine_inout As String, strCLP_out As String) As String

    '-----------------------------------------------------------------------
    'validates the ADD_CLP2 function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Output:    strCLP_out:         The clp string
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are ADD_CLP2 and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(8, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_ADD_CLP2 = strError
    Else
        'everything is fine!  return the proper string
        Validate_ADD_CLP2 = ""
        strCLP_out = strArgs(0)
        strLine_inout = "Add_CLP2(" & strCLP_out & ")"
    End If

End Function

Private Function Validate_EDIT_CLP(strLine_inout As String, strMessage_out As String) As String

    '-----------------------------------------------------------------------
    'validates the EDIT_CLP function line passed in
    '
    'Modify:    strLine_inout:      The function line (returns it formatted)
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '-----------------------------------------------------------------------

    Dim strError As String      'error string
    Dim lngNumArgs As Long      'number of function arguments
    Dim strArgs() As String     'array for function arguments

    'we already know the first chars in the line are EDIT_CLP and the line is trimmed
    
    'make sure there is 1 arg
    lngNumArgs = GetScriptFunctionArgs(8, strLine_inout, strArgs(), strError, 1)
    If strError <> "" Then
        Validate_EDIT_CLP = strError
    Else
        'everything is fine!  return the proper string
        Validate_EDIT_CLP = ""
        strMessage_out = strArgs(0)
        strLine_inout = "Edit_CLP(" & strArgs(0) & ")"
    End If

End Function
Friend Function ValidateScriptLine(strLine_inout As String) As String

    '---------------------------------------------------------------------
    'validates the script line passed in and formats it properly
    '
    'Modify:    strLine_inout:      The script line (returns it formatted)
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '---------------------------------------------------------------------

    Dim strPre As String            'the tabs/spaces before the start of the actual code
    Dim strLine As String           'the actual code in the script line
    Dim strError As String          'error string for function calls
    Dim strFunctionLine As String   'for function validation
    Dim strField As String
    Dim strOp As String
    Dim strData As String

    'assume line is ok
    ValidateScriptLine = ""

    If (strLine_inout <> "") And (strLine_inout <> "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*") Then
        Call GetScriptLineParts(strLine_inout, strPre, strLine)
        If strLine <> "" Then
            'don't bother testing a comment line
            If Left$(strLine, 1) <> ";" Then

                '-------------------
                'script line to test
                '-------------------

                If UCase(Left$(strLine, 3)) = "IF " Then
                    'validate comparison line
                    strError = ValidateComparisonLine(3, strLine, strField, strOp, strData)
                    If strError <> "" Then
                        ValidateScriptLine = strError
                    Else
                        strLine_inout = strLine
                        ValidateScriptLine = ""
                    End If
                ElseIf UCase(Left$(strLine, 8)) = "ELSE IF " Then
                    strError = ValidateComparisonLine(8, strLine, strField, strOp, strData)
                    If strError <> "" Then
                        ValidateScriptLine = strError
                    Else
                        strLine_inout = strLine
                        ValidateScriptLine = ""
                    End If
                ElseIf UCase(Left$(strLine, 7)) = "ELSEIF " Then
                    strError = ValidateComparisonLine(7, strLine, strField, strOp, strData)
                    If strError <> "" Then
                        ValidateScriptLine = strError
                    Else
                        strLine_inout = strLine
                        ValidateScriptLine = ""
                    End If
                ElseIf UCase(Left$(strLine, 4)) = "ELSE" Then
                    'ELSE word should be only thing on the line
                    If Len(strLine) > 4 Then
                        ValidateScriptLine = m_strLang_EndOfLineExpected
                    Else
                        strLine_inout = "Else"
                        ValidateScriptLine = ""
                    End If
                ElseIf (UCase(Left$(strLine, 5)) = "ENDIF") Then
                    'END IF words should be only thing on the line
                    If Len(strLine) > 5 Then
                        ValidateScriptLine = m_strLang_EndOfLineExpected
                    Else
                        strLine_inout = "End If"
                        ValidateScriptLine = ""
                    End If
                ElseIf (UCase(Left$(strLine, 6)) = "END IF") Then
                    'END IF words should be only thing on the line
                    If Len(strLine) > 6 Then
                        ValidateScriptLine = m_strLang_EndOfLineExpected
                    Else
                        strLine_inout = "End If"
                        ValidateScriptLine = ""
                    End If
                Else
                    'must be a function
                    strFunctionLine = strLine
                    strError = ValidateFunctionLine(strFunctionLine)
                    If strError <> "" Then
                        'command or statement not recognised
                        ValidateScriptLine = strError
                    Else
                        strLine_inout = strFunctionLine
                        ValidateScriptLine = ""
                    End If
                End If
            Else
                'comment
                strLine_inout = strLine
            End If
        Else
            'empty line not ""
            strLine_inout = strLine
        End If
    End If

End Function
Private Function ValidateComparisonLine(ByVal lngLenStatement_in As Long, strLine_inout As String, strField_out As String, strOp_out As String, strData_out As String) As String

    '---------------------------------------------------------------------
    'validates the script comparison line passed in and formats it properly
    '
    'Inputs:    lngLenStatement_in: The length of the if/elseif statement
    '
    'Modify:    strLine_inout:      The script line (returns it formatted)
    '
    'Returns:   Empty string if the line is ok
    '           Error message if the line has an error
    '---------------------------------------------------------------------

    Dim strExp As String    'expression string
    Dim lngFldIdx As Long   'loop index
    Dim lngOpIdx As Long    'loop index
    Dim lngListIdx As Long  'loop index
    Dim strUpper As String  'ucase field testing
    Dim strField As String  'storing the field
    Dim blnFound As Boolean 'testing flag
    Dim lngNum As Long      'number data comparison
    Dim strArgs() As String 'arguments for CONTAINS() function
    Dim lngArgIdx As Long   'argument index for CONTAINS() function
    Dim strOp As String     'temp var

    'we already know it starts with "IF", "ELSE IF", or "ELSEIF"

    'strip comparison statement and trim
    strExp = Trim$(Right$(strLine_inout, Len(strLine_inout) - lngLenStatement_in))

    If strExp = "" Then
        'no expression
        ValidateComparisonLine = m_strLang_ExpressionReqd
    Else
        'make sure there's a space between statement and expression
        If Mid$(strLine_inout, lngLenStatement_in, 1) = " " Then
            'test if it's a question
            If Left$(UCase$(strExp), 8) = "QUESTION" Then
                'return the field
                strField_out = "QUESTION"
                'grab the remainder of the expression
                strExp = Right$(strExp, Len(strExp) - 8)
                'check we have the end of the line ("yes" or "no") and store it
                If UCase$(Right$(strExp, 2)) = "NO" Then
                    strData_out = "NO"
                    'grab line upto "NO"
                    strExp = Trim$(Left$(strExp, Len(strExp) - 2))
                ElseIf UCase$(Right$(strExp, 3)) = "YES" Then
                    strData_out = "YES"
                    'grab line upto "YES"
                    strExp = Trim$(Left$(strExp, Len(strExp) - 3))
                Else
                    ValidateComparisonLine = m_strLang_Expected & ": ""= YES|NO"""
                    Exit Function
                End If
                'check there is an equals sign
                If Right$(strExp, 1) = "=" Then
                    strExp = Trim$(Left$(strExp, Len(strExp) - 1))
                Else
                    ValidateComparisonLine = m_strLang_Expected & ": ""= YES|NO"""
                    Exit Function
                End If
                'check we have a question to ask!
                If Left$(strExp, 1) <> "(" Then
                    'no opening bracket
                    ValidateComparisonLine = m_strLang_Expected & ": ""("""
                    Exit Function
                ElseIf Right$(strExp, 1) <> ")" Then
                    'no closing bracket
                    ValidateComparisonLine = m_strLang_Expected & ": "")"""
                    Exit Function
                ElseIf Len(strExp) < 3 Then
                    'no text
                    ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": ""QUESTION( text ) = YES|NO"""
                    Exit Function
                Else
                    'ok = parse and store value!
                    strOp_out = Mid$(strExp, 2, Len(strExp) - 2)
                End If
                '------------------------------------------------------------------------------
                'if here the line is ok and all data has been stored (FIELD, OPERATOR and DATA)
                'just format the line and return
                '------------------------------------------------------------------------------
                If UCase$(Left$(strLine_inout, 3)) = "IF " Then
                    strLine_inout = "If "
                ElseIf UCase$(Left$(strLine_inout, 8)) = "ELSE IF " Then
                    strLine_inout = "ElseIf "
                ElseIf UCase$(Left$(strLine_inout, 7)) = "ELSEIF " Then
                    strLine_inout = "ElseIf "
                End If
                strLine_inout = strLine_inout & "QUESTION(" & strOp_out & ") = " & strData_out
                ValidateComparisonLine = ""
                Exit Function
            ElseIf Left$(UCase$(strExp), 8) = "FILESIZE" Then
                'return the field
                strField_out = "QUESTION"
                'grab the remainder of the expression
                strExp = Right$(strExp, Len(strExp) - 8)
                'check we have the end of the line ("yes" or "no") and store it
                If UCase$(Right$(strExp, 2)) = "NO" Then
                    strData_out = "NO"
                    'grab line upto "NO"
                    strExp = Trim$(Left$(strExp, Len(strExp) - 2))
                ElseIf UCase$(Right$(strExp, 3)) = "YES" Then
                    strData_out = "YES"
                    'grab line upto "YES"
                    strExp = Trim$(Left$(strExp, Len(strExp) - 3))
                Else
                    ValidateComparisonLine = m_strLang_Expected & ": ""= YES|NO"""
                    Exit Function
                End If
                'check there is an equals sign
                If Right$(strExp, 1) = "=" Then
                    strExp = Trim$(Left$(strExp, Len(strExp) - 1))
                Else
                    ValidateComparisonLine = m_strLang_Expected & ": ""= YES|NO"""
                    Exit Function
                End If
                'check we have a question to ask!
                If Left$(strExp, 1) <> "(" Then
                    'no opening bracket
                    ValidateComparisonLine = m_strLang_Expected & ": ""("""
                    Exit Function
                ElseIf Right$(strExp, 1) <> ")" Then
                    'no closing bracket
                    ValidateComparisonLine = m_strLang_Expected & ": "")"""
                    Exit Function
                ElseIf Len(strExp) < 3 Then
                    'no text
                    ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": ""QUESTION( text ) = YES|NO"""
                    Exit Function
                Else
                    'ok = parse and store value!
                    strOp_out = Mid$(strExp, 2, Len(strExp) - 2)
                End If
                '------------------------------------------------------------------------------
                'if here the line is ok and all data has been stored (FIELD, OPERATOR and DATA)
                'just format the line and return
                '------------------------------------------------------------------------------
                If UCase$(Left$(strLine_inout, 3)) = "IF " Then
                    strLine_inout = "If "
                ElseIf UCase$(Left$(strLine_inout, 8)) = "ELSE IF " Then
                    strLine_inout = "ElseIf "
                ElseIf UCase$(Left$(strLine_inout, 7)) = "ELSEIF " Then
                    strLine_inout = "ElseIf "
                End If
                strLine_inout = strLine_inout & "QUESTION(" & strOp_out & ") = " & strData_out
                ValidateComparisonLine = ""
                Exit Function

            Else
                'get the field
                For lngFldIdx = 0 To m_lngMaxScriptFields
                    strUpper = UCase(Left$(strExp, Len(ma_udtScriptFields(lngFldIdx).strName)))
                    If strUpper = UCase(ma_udtScriptFields(lngFldIdx).strName) Then
    
                        'FIELD FOUND
                        With ma_udtScriptFields(lngFldIdx)
                            'return the field
                            strField_out = .strName
                            'grab the remainder of the expression
                            If .strName = "Key_" Then
                                If InStr(4, strExp, " ") > 0 Then
                                    strField_out = Left$(strExp, InStr(4, strExp, " ") - 1)
                                    strExp = Right$(strExp, Len(strExp) - InStr(4, strExp, " ") + 1)
                                    strField_out = "Key_" & Right$(strField_out, Len(strField_out) - 4)
                                End If
                            ElseIf .strName = "%" Then
                                If InStr(1, strExp, " ") > 0 Then
                                    strField_out = Left$(strExp, InStr(strExp, " ") - 1)
                                    strExp = Right$(strExp, Len(strExp) - InStr(strExp, " ") + 1)
                                End If
                            Else
                                strExp = Right$(strExp, Len(strExp) - Len(strUpper))
                            End If
                            'make sure there is a comparison operator
                            If .eValueType = e_ScriptValType_Text Then
                                'contains/excludes text operator
                                strOp = UCase(Left$(strExp, 8))
                                If strOp = "CONTAINS" Or strOp = "EXCLUDES" Then
                                    'no space before operator
                                    ValidateComparisonLine = m_strLang_Expected & ": "" """
                                    Exit Function
                                End If
                                strOp = UCase(Left$(strExp, 9))
                                If (strOp <> " CONTAINS") And (strOp <> " EXCLUDES") Then
                                    'wrong or missing operator
                                    ValidateComparisonLine = m_strLang_InvalidOp & vbCrLf & vbCrLf & m_strLang_Expected & ": ""CONTAINS()"" / ""EXCLUDES()"""
                                    Exit Function
                                End If
                                'ok - grab comparison value
                                strExp = Trim$(Right$(strExp, Len(strExp) - 9))
                                strOp_out = Right$(strOp, Len(strOp) - 1) '"CONTAINS" or "EXCLUDES"
                            Else
                                'other operator
                                strExp = LTrim(strExp)
                                blnFound = False
                                For lngOpIdx = 0 To (.lngTotalOperators - 1)
                                    If Left$(strExp, Len(.strOperators(lngOpIdx))) = .strOperators(lngOpIdx) Then
                                        'grab comparison value
                                        strOp_out = .strOperators(lngOpIdx)
                                        strExp = Trim$(Right$(strExp, Len(strExp) - Len(.strOperators(lngOpIdx))))
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngOpIdx
                                If blnFound = False Then
                                    'wrong or missing operator
                                    ValidateComparisonLine = m_strLang_InvalidOp & vbCrLf & vbCrLf & m_strLang_Expected & ": "
                                    For lngOpIdx = 0 To (.lngTotalOperators - 1)
                                        ValidateComparisonLine = ValidateComparisonLine & """" & .strOperators(lngOpIdx) & """"
                                        If lngOpIdx < (.lngTotalOperators - 1) Then
                                            ValidateComparisonLine = ValidateComparisonLine & ", "
                                        Else
                                            ValidateComparisonLine = ValidateComparisonLine & "."
                                        End If
                                    Next lngOpIdx
                                    Exit Function
                                End If
                            End If
                            'if here then all is ok so far... just comparison data to check!
                            If .eValueType = e_ScriptValType_Text Then
                                'CONTAINS(TEXT[||TEXT*])
                                If Left$(strExp, 1) <> "(" Then
                                    'no opening bracket
                                    ValidateComparisonLine = m_strLang_Expected & ": ""("""
                                    Exit Function
                                ElseIf Right$(strExp, 1) <> ")" Then
                                    'no closing bracket
                                    ValidateComparisonLine = m_strLang_Expected & ": "")"""
                                    Exit Function
                                ElseIf Len(strExp) < 3 Then
                                    'no text
                                    ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": ""CONTAINS( text [ || text || text... ] )"""
                                    Exit Function
                                Else
                                    'ok = parse and store value!
                                    strData_out = Mid$(strExp, 2, Len(strExp) - 2)
                                    strArgs() = Split(strData_out, "||")
                                    For lngArgIdx = 0 To UBound(strArgs)
                                        If strArgs(lngArgIdx) = "" Then
                                            ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": ""CONTAINS( text [ || text || text... ] )"""
                                            Exit Function
                                        End If
                                    Next lngArgIdx
                                End If
                            ElseIf .eValueType = e_ScriptValType_Number Then
                                'IS A NUMBER
                                If InStr(strExp, " ") <> 0 Then
                                    'end of line expected
                                    ValidateComparisonLine = m_strLang_EndOfLineExpected
                                    Exit Function
                                End If
                                lngNum = ConvertStringToNumber(strExp)
                                If lngNum < 0 Then
                                    'invalid number
                                    ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": " & m_strLang_Number & " > -1, " & m_strLang_Number & " < 10000."
                                    Exit Function
                                Else
                                    'ok = store value!
                                    strData_out = Trim$(lngNum)
                                End If
                            Else
                                If InStr(strExp, " ") <> 0 Then
                                    'end of line expected
                                    ValidateComparisonLine = m_strLang_EndOfLineExpected
                                    Exit Function
                                End If
                                'MATCHES ITEM IN LIST
                                blnFound = False
                                strUpper = UCase$(strExp)
                                For lngListIdx = 0 To (.lngTotalListValues - 1)
                                    If strUpper = UCase(.strListValues(lngListIdx)) Then
                                        strData_out = .strListValues(lngListIdx)
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngListIdx
                                If blnFound = False Then
                                    'invalid value
                                    ValidateComparisonLine = m_strLang_InvalidValue & vbCrLf & vbCrLf & m_strLang_Expected & ": "
                                    For lngListIdx = 0 To (.lngTotalListValues - 1)
                                        ValidateComparisonLine = ValidateComparisonLine & """" & .strListValues(lngListIdx) & """"
                                        If lngListIdx < (.lngTotalListValues - 1) Then
                                            ValidateComparisonLine = ValidateComparisonLine & ", "
                                        Else
                                            ValidateComparisonLine = ValidateComparisonLine & "."
                                        End If
                                    Next lngListIdx
                                    Exit Function
                                End If
                            End If
                            '------------------------------------------------------------------------------
                            'if here the line is ok and all data has been stored (FIELD, OPERATOR and DATA)
                            'just format the line and return
                            '------------------------------------------------------------------------------
                            If UCase$(Left$(strLine_inout, 3)) = "IF " Then
                                strLine_inout = "If "
                            ElseIf UCase$(Left$(strLine_inout, 8)) = "ELSE IF " Then
                                strLine_inout = "ElseIf "
                            ElseIf UCase$(Left$(strLine_inout, 7)) = "ELSEIF " Then
                                strLine_inout = "ElseIf "
                            End If
                            If .eValueType = e_ScriptValType_Text Then
                                strLine_inout = strLine_inout & strField_out & " " & strOp_out & "(" & strData_out & ")"
                            Else
                                strLine_inout = strLine_inout & strField_out & " " & strOp_out & " " & strData_out
                            End If
                            ValidateComparisonLine = ""
                            Exit Function
                        End With
                    End If
                Next lngFldIdx
            End If
            'if we've got here the field wasn't found
            ValidateComparisonLine = m_strLang_UnknownField & vbCrLf & vbCrLf & m_strLang_Expected & ": ""QUESTION( )"", "
            For lngFldIdx = 0 To m_lngMaxScriptFields
                ValidateComparisonLine = ValidateComparisonLine & """"
                If ma_udtScriptFields(lngFldIdx).strName = "%" Then
                    ValidateComparisonLine = ValidateComparisonLine & "%...%"""
                Else
                    ValidateComparisonLine = ValidateComparisonLine & ma_udtScriptFields(lngFldIdx).strName & """"
                End If
                If lngFldIdx < m_lngMaxScriptFields Then
                    ValidateComparisonLine = ValidateComparisonLine & ", "
                Else
                    ValidateComparisonLine = ValidateComparisonLine & "."
                End If
            Next lngFldIdx
        Else
            'no space before expression
            ValidateComparisonLine = m_strLang_Expected & ": "" """
        End If
    End If


End Function

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
Private Sub GetEmulatorInfo(strScript_in As String, ByVal blnUseShort_in As Boolean, strName_in As String, strPath_in As String, udtEmulator_inout As EmulatorType, strFiles_in() As String, ByVal lngNumFiles_in As Long, blnUseDefaultTypes_in As Boolean, strTypes_in As String, lngWait_in As Long)

    '---------------------------------------------------------
    'Fills the EmulatorType structure passed in.
    '
    'Inputs:    strScript_in:           script string.
    '           blnUseShort_in:         flag for short filename usage.
    '           strName_in:             Name of Sidplayer.
    '           strPath_in:             The path of the EXE to test.
    '           strFiles_in():          Array of emulator files.
    '           lngNumFiles_in:         Number of emulator files.
    '           blnUseDefaultTypes_in:  Flag to use default file types or not
    '           strTypes_in:            The supported file types if not using default
    '           lngWait_in:             How to wait for the emu to finish (if rezipping for example)
    '
    'Modify:    udtEmulator_inout:  The structure to fill.
    '---------------------------------------------------------

    Dim lngFileIdx As Long      'loop index
    Dim strValidateLocal As String
    Dim iLen As Integer

    'has to have at least 3 letters ("c:\")
    With udtEmulator_inout
    
        'save info not related to path
        .strName = strName_in
        .blnUseShortFilenames = blnUseShort_in
        .strScript = strScript_in
        .blnUseDefaultTypes = blnUseDefaultTypes_in
        .strTypes = strTypes_in
        .strFullPathAndFilename = strPath_in
        .lngWait = lngWait_in
        If .lngWait < 0 Or .lngWait > 1 Then
            .lngWait = 1
        End If
        
        Call m_xobjFunc.GetFileAndPath(.strFullPathAndFilename, .strFullPathOnly, .strFullFilenameOnly)
        iLen = Len(.strFullPathOnly)
        
        If iLen > 2 Then
            'convert root "\" to current drive root
            If Left$(.strFullPathOnly, 1) = "\" Then
                .strFullPathOnly = Left$(m_strAppPathWithBackSlash, 3) & Right(.strFullPathOnly, (iLen - 1))
                iLen = Len(.strFullPathOnly)
            End If
        
            'convert relative path to absolute
            If Mid$(.strFullPathOnly, 2, 1) <> ":" Then
                .strFullPathOnly = m_strAppPathWithBackSlash & .strFullPathOnly
                iLen = Len(.strFullPathOnly)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If m_xobjFunc.DirExists(.strFullPathOnly) = True Then
                    .strFullPathOnly = CurDir()
                End If
            End If
            
            'if specified path does not exist, try to validate on the current drive
            If m_xobjFunc.DirExists(.strFullPathOnly) = False Then
                strValidateLocal = Left$(m_strAppPathWithBackSlash, 3) & Right(.strFullPathOnly, (iLen - 3))
                If m_xobjFunc.DirExists(strValidateLocal) = True Then
                    .strFullPathOnly = strValidateLocal
                End If
            End If

            .strFullPathAndFilename = .strFullPathOnly & "\" & .strFullFilenameOnly

            'test drive type
            .ePathType = GetTypeOfDrive(Left$(.strFullPathAndFilename, 3))
            If .ePathType <> e_PathLocation_Invalid Then
                'check if path exists
                If m_xobjFunc.FileExists(.strFullPathAndFilename, False, "") = False Then
                    If .ePathType = e_PathLocation_FixedHD Then
                        'file doesn't exist on HD
                        .ePathType = e_PathLocation_Invalid
                        Exit Sub
                    End If
                Else
                    'file found - get info
                    .strShortPathAndFilename = m_xobjFunc.Long2Short(.strFullPathAndFilename)
                    Call m_xobjFunc.GetFileAndPath(.strFullPathAndFilename, .strFullPathOnly, .strFullFilenameOnly)
                    Call m_xobjFunc.GetFileAndPath(.strShortPathAndFilename, .strShortPathOnly, .strShortFilenameOnly)
                    .strFullPathOnlyWithBackSlash = .strFullPathOnly
                    .strShortPathOnlyWithBackSlash = .strShortPathOnly
                    If Right$(.strFullPathOnlyWithBackSlash, 1) <> "\" Then .strFullPathOnlyWithBackSlash = .strFullPathOnlyWithBackSlash & "\"
                    If Right$(.strShortPathOnlyWithBackSlash, 1) <> "\" Then .strShortPathOnlyWithBackSlash = .strShortPathOnlyWithBackSlash & "\"
                End If
                
                'load associated file info
                .lngNumFiles = lngNumFiles_in
                If .lngNumFiles > 0 Then
                    Dim strWorkingPath As String
                    Dim strWorkingName As String
                    ReDim .strFiles(.lngNumFiles)
                    
                    For lngFileIdx = 0 To (.lngNumFiles - 1)
                        Call m_xobjFunc.GetFileAndPath(strFiles_in(lngFileIdx), strWorkingPath, strWorkingName)
                        iLen = Len(strWorkingPath)
            
                        If iLen > 2 Then
                            'convert root "\" to current drive root
                            If Left$(strWorkingPath, 1) = "\" Then
                                strWorkingPath = Left$(m_strAppPathWithBackSlash, 3) & Right(strWorkingPath, (iLen - 1))
                                iLen = Len(strWorkingPath)
                            End If
        
                            'convert relative path to absolute
                            If Mid$(strWorkingPath, 2, 1) <> ":" Then
                                strWorkingPath = m_strAppPathWithBackSlash & strWorkingPath
                                iLen = Len(strWorkingPath)
                                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                                If m_xobjFunc.DirExists(strWorkingPath) = True Then
                                    strWorkingPath = CurDir()
                                End If
                            End If
            
                            'if specified path does not exist, try to validate on the current drive
                            If m_xobjFunc.DirExists(strWorkingPath) = False Then
                                strValidateLocal = Left$(m_strAppPathWithBackSlash, 3) & Right(strWorkingPath, (iLen - 3))
                                If m_xobjFunc.DirExists(strValidateLocal) = True Then
                                    strWorkingPath = strValidateLocal
                                End If
                            End If
                        End If
                    
                        .strFiles(lngFileIdx) = strWorkingPath & "\" & strWorkingName
                        '.strFiles(lngFileIdx) = strFiles_in(lngFileIdx)
                    Next lngFileIdx
                End If
            End If
        Else
            'invalid path\filename string
            .ePathType = e_PathLocation_Invalid
        End If
    End With

End Sub


Friend Function GetScriptLineText(strText_in As String, ByVal lngLineNumber_in As Long) As String

    '-------------------------------------------------------------------------
    'gets the line of text from the text string passed in
    'at the passed in line munber
    '
    'Inputs:    strText_in:         The full multiline text string
    '           lngLineNumber_in:   The line number of the line of text to get
    '
    'Returns:   The text from the specified line
    '-------------------------------------------------------------------------

    Dim lngLineIdx As Long      'loop index
    Dim lngStartChar As Long    'the starting char index of current line
    Dim lngLastChar As Long     'the ending character index of current line
    
    If InStr(strText_in, vbCrLf) = 0 Then
        'blank line - end of script
        GetScriptLineText = "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*"
    Else
        'must be some lines
        lngLineIdx = 1
        lngStartChar = 1
        Do
            lngLastChar = InStr(lngStartChar, strText_in, vbCrLf)
            If lngLastChar > 0 Then
                'line found
                If lngLineIdx = lngLineNumber_in Then
                    GetScriptLineText = Mid$(strText_in, lngStartChar, lngLastChar - lngStartChar)
                    Exit Function
                End If
                lngStartChar = lngLastChar + 2
                lngLineIdx = lngLineIdx + 1
            Else
                'maybe last line
                If lngLineIdx = lngLineNumber_in Then
                    GetScriptLineText = Right$(strText_in, Len(strText_in) - (lngStartChar - 1))
                Else
                    'no more lines
                    GetScriptLineText = "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*"
                End If
                Exit Function
            End If
        Loop
    End If

End Function

Public Function GetINIMessage(strSection_in As String, strKey_in As String, strDefault_in As String) As String

    '------------------------------------------------------------------------
    'Reads a text item from the GemusText.ini file
    '
    'Inputs:    strSection_in:      The INI file section
    '           strKey_in:          The INI file key
    '           strDefault_in:      The default string to return if not found
    '
    'Returns:   The found string in the INI file or the default string passed in!
    '------------------------------------------------------------------------

    Dim lngRes As Long                  'return Value
    Dim strReturned As String * 2096    'max length of ini file entry
    
    'get the required message
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, m_strLangINIPathFile)
    GetINIMessage = Left$(strReturned, lngRes)

End Function

Private Function GetNumLeftSpecificScriptLines(ByVal strText_in As String) As Long

    Dim lngLineIdx As Long
    Dim lngLen As Long
    Dim strFind As String

    GetNumLeftSpecificScriptLines = 0
    strFind = UCase$(strText_in)
    lngLen = Len(strFind)
    
    If m_lngTotalScriptLines > 0 Then
        For lngLineIdx = 0 To (m_lngTotalScriptLines - 1)
            If strFind = "ELSE" Then
                If UCase$(ma_strScriptLines(lngLineIdx)) = "ELSE" Then
                    GetNumLeftSpecificScriptLines = GetNumLeftSpecificScriptLines + 1
                End If
            Else
                If UCase$(Left$(ma_strScriptLines(lngLineIdx), lngLen)) = strFind Then
                    GetNumLeftSpecificScriptLines = GetNumLeftSpecificScriptLines + 1
                End If
            End If
        Next lngLineIdx
    End If

End Function
Private Function GetScriptLine(ByVal lngLineIdx_in As Long) As String

    If lngLineIdx_in > (m_lngTotalScriptLines - 1) Then
        GetScriptLine = "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*"
    Else
        GetScriptLine = ma_strScriptLines(lngLineIdx_in)
    End If

End Function
Private Function ValidateScriptBlock(ByVal lngStartLine_in As Long, lngNextLine_out As Long, lngErrLine_out As Long, ByVal lngLevel_in As Long) As String

    '------------------------------------------------------------------------
    'validates the structure of a script block (if/elseif/else/endif recursion)
    '
    'Inputs:    lngLevel_in:            Are we in the scripts root level?
    '
    'Outputs:   lngStartLine_in:        Starting line for this recursive pass
    '           lngNextLine_out:        Starting line of next recursive pass
    '           lngErrLine_out:         Error line if failed
    '
    'Returns:   Error message or "" if successful
    '------------------------------------------------------------------------

    Dim blnDoneIF As Boolean
    Dim blnDoneELSE As Boolean
    Dim lngLineIdx As Long
    Dim strLine As String
    Dim strUpper As String
    Dim strError As String
    Dim lngTabIdx As Long

    blnDoneIF = False
    blnDoneELSE = False
    
    lngLineIdx = lngStartLine_in
    strLine = GetScriptLine(lngLineIdx)
    
    Do Until strLine = "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*"

        strUpper = UCase$(strLine)
        If Left$(strUpper, 3) = "IF " Then
            'IF line found
            blnDoneIF = True
            'format the tabs before the script line text
            ma_strScriptLinesPre(lngLineIdx) = ""
            For lngTabIdx = 1 To lngLevel_in
                ma_strScriptLinesPre(lngLineIdx) = ma_strScriptLinesPre(lngLineIdx) & vbTab
            Next lngTabIdx
            'process sub block
            strError = ValidateScriptBlock(lngLineIdx + 1, lngNextLine_out, lngErrLine_out, lngLevel_in + 1)
            If strError = "" Then
                lngLineIdx = lngNextLine_out - 1
            Else
                ValidateScriptBlock = strError
                Exit Function
            End If

        ElseIf Left$(strUpper, 7) = "ELSEIF " Or Left$(strUpper, 8) = "ELSE IF " Then
            'ELSE IF line found
            If blnDoneIF = True Then
                If blnDoneELSE = True Then
                    ValidateScriptBlock = m_strLang_Expected & " 'END IF'"
                    lngErrLine_out = lngLineIdx
                    Exit Function
                Else
                    'format the tabs before the script line text
                    ma_strScriptLinesPre(lngLineIdx) = ""
                    For lngTabIdx = 1 To lngLevel_in
                        ma_strScriptLinesPre(lngLineIdx) = ma_strScriptLinesPre(lngLineIdx) & vbTab
                    Next lngTabIdx
                    'process sub block
                    strError = ValidateScriptBlock(lngLineIdx + 1, lngNextLine_out, lngErrLine_out, lngLevel_in + 1)
                    If strError = "" Then
                        lngLineIdx = lngNextLine_out - 1
                    Else
                        ValidateScriptBlock = strError
                        Exit Function
                    End If
                End If
            Else
                If lngLevel_in = 0 Then
                    ValidateScriptBlock = "'ELSE IF' " & m_strLang_Without & " 'IF'"
                    lngErrLine_out = lngLineIdx
                    Exit Function
                Else
                    'parent ELSE IF
                    lngNextLine_out = lngLineIdx
                    ValidateScriptBlock = ""
                    Exit Function
                End If
            End If

        ElseIf Left$(strUpper, 4) = "ELSE" Then
            'ELSE line found
            If blnDoneIF = True Then
                If blnDoneELSE = True Then
                    ValidateScriptBlock = m_strLang_Expected & " 'END IF'"
                    lngErrLine_out = lngLineIdx
                    Exit Function
                Else
                    blnDoneELSE = True
                    'format the tabs before the script line text
                    ma_strScriptLinesPre(lngLineIdx) = ""
                    For lngTabIdx = 1 To lngLevel_in
                        ma_strScriptLinesPre(lngLineIdx) = ma_strScriptLinesPre(lngLineIdx) & vbTab
                    Next lngTabIdx
                    'process sub block
                    strError = ValidateScriptBlock(lngLineIdx + 1, lngNextLine_out, lngErrLine_out, lngLevel_in + 1)
                    If strError = "" Then
                        lngLineIdx = lngNextLine_out - 1
                    Else
                        ValidateScriptBlock = strError
                        Exit Function
                    End If
                End If
            Else
                If lngLevel_in = 0 Then
                    ValidateScriptBlock = "'ELSE' " & m_strLang_Without & " 'IF'"
                    lngErrLine_out = lngLineIdx
                    Exit Function
                Else
                    'parent ELSE
                    lngNextLine_out = lngLineIdx
                    ValidateScriptBlock = ""
                    Exit Function
                End If
            End If

        ElseIf Left$(strUpper, 5) = "ENDIF" Or Left$(strUpper, 6) = "END IF" Then
            'END IF line found
            If blnDoneIF = True Then
                blnDoneIF = False
                blnDoneELSE = False
                'format the tabs before the script line text
                ma_strScriptLinesPre(lngLineIdx) = ""
                For lngTabIdx = 1 To lngLevel_in
                    ma_strScriptLinesPre(lngLineIdx) = ma_strScriptLinesPre(lngLineIdx) & vbTab
                Next lngTabIdx
            Else
                If lngLevel_in = 0 Then
                    ValidateScriptBlock = "'END IF' " & m_strLang_Without & " 'IF'"
                    lngErrLine_out = lngLineIdx
                    Exit Function
                Else
                    'parent END IF
                    lngNextLine_out = lngLineIdx
                    ValidateScriptBlock = ""
                    Exit Function
                End If
            End If
        
        Else
            'Function/blank/comment line
            'format the tabs before the script line text
            ma_strScriptLinesPre(lngLineIdx) = ""
            If strLine <> "" Then
                For lngTabIdx = 1 To lngLevel_in
                    ma_strScriptLinesPre(lngLineIdx) = ma_strScriptLinesPre(lngLineIdx) & vbTab
                Next lngTabIdx
            End If
        End If

        'move to next line
        lngLineIdx = lngLineIdx + 1
        strLine = GetScriptLine(lngLineIdx)

    Loop

    'if here then reached end of script

    If lngLevel_in = 0 Then
        If blnDoneIF = True Then
            lngErrLine_out = lngLineIdx
            ValidateScriptBlock = m_strLang_Expected & " 'END IF'"
        Else
            ValidateScriptBlock = ""
        End If
    Else
        lngErrLine_out = lngLineIdx
        ValidateScriptBlock = m_strLang_Expected & " 'END IF'"
    End If

End Function
Private Sub GetScriptFunctionsToProcess(ByVal lngStartLine_in As Long, lngNextLine_out As Long, ByVal lngLevel_in As Long, ByVal blnProcess_in As Boolean)

    '---------------------------------------------------------------------------
    'finds and returns all script lines that need processing by evaluating all
    'the expression/comparison lines (if/elseif/else/endif recursion)
    '
    'Inputs:    lngLevel_in:            Are we in the scripts root level?
    '           blnProcess_in:          Flag for when we want to store functions
    '
    'Outputs:   lngStartLine_in:        Starting line for this recursive pass
    '           lngNextLine_out:        Starting line of next recursive pass
    '---------------------------------------------------------------------------

    Dim blnDoneIF As Boolean
    Dim blnDoneELSE As Boolean
    Dim lngLineIdx As Long
    Dim strLine As String
    Dim strUpper As String
    Dim strError As String
    Dim lngTabIdx As Long
    Dim blnEvaluated As Boolean
    Dim blnBlockDone As Boolean

    blnDoneIF = False
    blnDoneELSE = False
    
    lngLineIdx = lngStartLine_in
    strLine = GetScriptLine(lngLineIdx)
    
    Do Until strLine = "|!%!__||||&:@~~~|||~:~:@@@}%{£%^£$%£+$%  +£$%£$+%ÑÍÄ*"

        strUpper = UCase$(strLine)

        If Left$(strUpper, 3) = "IF " Then
            'IF line found
            blnDoneIF = True
            'process sub block
            If blnProcess_in = True Then
                blnEvaluated = EvaluateLine(strLine)
                blnBlockDone = blnEvaluated
            Else
                blnEvaluated = False
            End If
            Call GetScriptFunctionsToProcess(lngLineIdx + 1, lngNextLine_out, lngLevel_in + 1, blnEvaluated)
            lngLineIdx = lngNextLine_out - 1

        ElseIf Left$(strUpper, 7) = "ELSEIF " Or Left$(strUpper, 8) = "ELSE IF " Then
            'ELSE IF line found
            If blnDoneIF = True Then
                'process sub block
                If blnProcess_in = True Then
                    If blnBlockDone = False Then
                        blnEvaluated = EvaluateLine(strLine)
                        blnBlockDone = blnEvaluated
                    Else
                        blnEvaluated = False
                    End If
                Else
                    blnEvaluated = False
                End If
                Call GetScriptFunctionsToProcess(lngLineIdx + 1, lngNextLine_out, lngLevel_in + 1, blnEvaluated)
                lngLineIdx = lngNextLine_out - 1
            Else
                'parent ELSE IF
                lngNextLine_out = lngLineIdx
                Exit Sub
            End If

        ElseIf Left$(strUpper, 4) = "ELSE" Then
            'ELSE line found
            If blnDoneIF = True Then
                blnDoneELSE = True
                'process sub block
                If blnProcess_in = True Then
                    If blnBlockDone = False Then
                        blnEvaluated = True
                        blnBlockDone = True
                    Else
                        blnEvaluated = False
                    End If
                Else
                    blnEvaluated = False
                End If
                Call GetScriptFunctionsToProcess(lngLineIdx + 1, lngNextLine_out, lngLevel_in + 1, blnEvaluated)
                lngLineIdx = lngNextLine_out - 1
            Else
                'parent ELSE
                lngNextLine_out = lngLineIdx
                Exit Sub
            End If

        ElseIf Left$(strUpper, 5) = "ENDIF" Or Left$(strUpper, 6) = "END IF" Then
            'END IF line found
            If blnDoneIF = True Then
                blnDoneIF = False
                blnDoneELSE = False
            Else
                'parent END IF
                lngNextLine_out = lngLineIdx
                Exit Sub
            End If
        
        Else
            'Function/blank/comment line
            If (Trim$(strLine) <> "") And (Left$(strLine, 1) <> ";") Then
                'function to process?
                If blnProcess_in = True Then
                    'yep!
                    ReDim Preserve ma_strScriptLinesToProcess(m_lngTotalScriptLinesToProcess)
                    ma_strScriptLinesToProcess(m_lngTotalScriptLinesToProcess) = strLine
                    m_lngTotalScriptLinesToProcess = m_lngTotalScriptLinesToProcess + 1
                End If
            End If
        End If

        'move to next line
        lngLineIdx = lngLineIdx + 1
        strLine = GetScriptLine(lngLineIdx)

    Loop

End Sub

Public Function ValidateScriptStructure(strScript_in As String, lngErrLine_out As Long) As String

    Dim lngNumIFs As Long
    Dim lngNumELSEIFs As Long
    Dim lngNumELSEs As Long
    Dim lngNumENDIFs As Long
    Dim lngLineIdx As Long
    Dim lngErrLine As Long
    Dim lngNextLine As Long
    Dim strText As String
    Dim strPre As String

    ValidateScriptStructure = ""
    lngErrLine_out = -1

    If strScript_in <> "" Then

        'match number of IFs, ENDIFs and ELSEs
        lngNumIFs = GetNumLeftSpecificScriptLines("if")
        lngNumENDIFs = GetNumLeftSpecificScriptLines("end if")
        lngNumENDIFs = lngNumENDIFs + GetNumLeftSpecificScriptLines("endif")
        If lngNumIFs < lngNumENDIFs Then
            ValidateScriptStructure = m_strLang_TooMany & " 'END IF' " & m_strLang_Statements & "!"
        ElseIf lngNumIFs > lngNumENDIFs Then
            ValidateScriptStructure = m_strLang_TooMany & " 'IF' " & m_strLang_Statements & "!"
        Else
            lngNumELSEIFs = GetNumLeftSpecificScriptLines("else if")
            lngNumELSEIFs = lngNumELSEIFs + GetNumLeftSpecificScriptLines("elseif")
            If (lngNumELSEIFs > lngNumIFs) And (lngNumIFs = 0) Then
                ValidateScriptStructure = m_strLang_TooMany & " 'ELSE IF' " & m_strLang_Statements & "!"
            Else
                lngNumELSEs = GetNumLeftSpecificScriptLines("else")
                If lngNumELSEs > lngNumIFs Then
                    ValidateScriptStructure = m_strLang_TooMany & " 'ELSE' " & m_strLang_Statements & "!"
                Else
                    'OK to validate the recursive script block structures!
                    ValidateScriptStructure = ValidateScriptBlock(0, lngNextLine, lngErrLine, 0)
                    lngErrLine_out = lngErrLine + 1
                End If
            End If
        End If
    End If

End Function

Private Function WriteCFGValue(strFile_in As String, strKey_in As String, strValue_in As String) As Boolean

    Dim intFileNum As Integer
    Dim strLines() As String
    Dim lngLineIdx As Long
    Dim blnFound As Boolean
    Dim strUpper As String
    Dim strLine As String
    Dim strChar As String

    On Error GoTo Error_Handler

    'write file if doesn't exist
    If m_xobjFunc.FileExists(strFile_in, False, "") = False Then
        intFileNum = FreeFile()
        Open strFile_in For Output As #intFileNum
        Print #intFileNum, ""
        Close #intFileNum
    End If

    intFileNum = FreeFile()
    lngLineIdx = 0
    ReDim strLines(0)

    Open strFile_in For Input As #intFileNum
    Do Until EOF(intFileNum)
        ReDim Preserve strLines(lngLineIdx)
        Line Input #intFileNum, strLines(lngLineIdx)
        lngLineIdx = lngLineIdx + 1
    Loop
    Close #intFileNum

    'find line with key
    blnFound = False
    strUpper = UCase$(strKey_in)
    If lngLineIdx > 0 Then
        For lngLineIdx = 0 To UBound(strLines)
            strLine = Trim$(strLines(lngLineIdx))
            If UCase$(Left$(strLine, Len(strUpper))) = strUpper Then
                If Len(strLine) > Len(strUpper) Then
                    If NextNonWhiteSpaceCharIs(strLine, Len(strUpper) + 1, "=") = True Then
                        'found line with key
                        strLines(lngLineIdx) = strKey_in & "=" & strValue_in
                        blnFound = True
                    End If
                ElseIf Len(strLine) = Len(strUpper) Then
                    'found line with key
                    strLines(lngLineIdx) = strKey_in & "=" & strValue_in
                    blnFound = True
                End If
            End If
        Next lngLineIdx
    End If
    If blnFound = False Then
        'add the line to end
        ReDim Preserve strLines(UBound(strLines) + 1)
        strLines(UBound(strLines)) = strKey_in & "=" & strValue_in
    End If
    'write back the file
    Kill strFile_in
    intFileNum = FreeFile()
    Open strFile_in For Output As #intFileNum
    For lngLineIdx = 0 To UBound(strLines)
        Print #intFileNum, strLines(lngLineIdx)
    Next lngLineIdx
    Close #intFileNum
    'ok!
    WriteCFGValue = True

    Exit Function
    
Error_Handler:
    WriteCFGValue = False

End Function
Private Function WriteCFGValue2(strFile_in As String, strKey_in As String, strDelim_in As String, strValue_in As String) As Boolean

    Dim intFileNum As Integer
    Dim strLines() As String
    Dim lngLineIdx As Long
    Dim blnFound As Boolean
    Dim strUpper As String
    Dim strLine As String
    Dim strDelim As String

    On Error GoTo Error_Handler

    'write file if doesn't exist
    If m_xobjFunc.FileExists(strFile_in, False, "") = False Then
        intFileNum = FreeFile()
        Open strFile_in For Output As #intFileNum
        Print #intFileNum, ""
        Close #intFileNum
    End If

    strDelim = strDelim_in
    If strDelim = "" Then
        strDelim = " "
    End If
    intFileNum = FreeFile()
    lngLineIdx = 0
    ReDim strLines(0)

    Open strFile_in For Input As #intFileNum
    Do Until EOF(intFileNum)
        ReDim Preserve strLines(lngLineIdx)
        Line Input #intFileNum, strLines(lngLineIdx)
        lngLineIdx = lngLineIdx + 1
    Loop
    Close #intFileNum

    'find line with key
    blnFound = False
    strUpper = UCase$(strKey_in)
    If lngLineIdx > 0 Then
        For lngLineIdx = 0 To UBound(strLines)
            strLine = Trim$(strLines(lngLineIdx))
            If UCase$(Left$(strLine, Len(strUpper))) = strUpper Then
                If Len(strLine) > Len(strUpper) Then
                    'found line with key
                    If Mid$(strLine, Len(strUpper) + 1, Len(strDelim)) = strDelim Then
                        If Len(strValue_in) > 0 Then
                            strLines(lngLineIdx) = strKey_in & strDelim & strValue_in
                        Else
                            strLines(lngLineIdx) = strKey_in
                        End If
                        blnFound = True
                    Else
                        'find first non whitespace char(s)
                        If NextNonWhiteSpaceCharIs(strLine, Len(strUpper) + 1, strDelim) = True Then
                            strLines(lngLineIdx) = strKey_in & strDelim & strValue_in
                            blnFound = True
                        End If
                    End If
                ElseIf Len(strLine) = Len(strUpper) Then
                    'found line with key
                    strLines(lngLineIdx) = strKey_in & strDelim & strValue_in
                    blnFound = True
                End If
            End If
        Next lngLineIdx
    End If
    If blnFound = False Then
        'add the line to end
        ReDim Preserve strLines(UBound(strLines) + 1)
        strLines(UBound(strLines)) = strKey_in & strDelim & strValue_in
    End If
    'write back the file
    Kill strFile_in
    intFileNum = FreeFile()
    Open strFile_in For Output As #intFileNum
    For lngLineIdx = 0 To UBound(strLines)
        Print #intFileNum, strLines(lngLineIdx)
    Next lngLineIdx
    Close #intFileNum
    'ok!
    WriteCFGValue2 = True

    Exit Function
    
Error_Handler:
    WriteCFGValue2 = False

End Function

Private Sub Class_Initialize()

    Dim strTextFile As String

    'create links to external objects
    Set m_xobjFunc = New GBFuncs.clsFunctions
    Set m_xobjRegistry = New GBFuncs.clsRegistry

    'for shelling dos progs
    m_strCommandCom = m_xobjFunc.Long2Short(Environ$("COMSPEC")) & " /c "

    'set variables
    m_strAppPathWithBackSlash = App.Path
    If g_blnInVBIDE Then
        m_strAppPathWithBackSlash = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    If Right$(m_strAppPathWithBackSlash, 1) <> "\" Then m_strAppPathWithBackSlash = m_strAppPathWithBackSlash & "\"
    m_strLangINIPathFile = m_strAppPathWithBackSlash & "GemusText.ini"
    m_strAppName = "GEMUS"
    m_blnWaitIfNotRezipping = False

    'create ini file if it doesn't exist
    If m_xobjFunc.FileExists(m_strAppPathWithBackSlash & "\GemusText.ini", False, "") = False Then
        strTextFile = GetResourceFileAsString(101)
        Call CreateFile(m_strAppPathWithBackSlash & "\GemusText.ini", strTextFile)
    End If

    m_blnRezipLegacy = False
    m_blnUserWantsToRepack = False
    m_blnAskUserBeforeRepack = False
    me_RepackAddtional = eRezipAdditionalAsk
    m_blnNotifyUserIfRepackWasOK = False
    m_blnAttemptRepack = False
    
    m_blnSilentMode = False
    m_blnSimulateRun = False
    m_strSimulateKV = ""

End Sub
Private Sub CreateFile(strFilename_in As String, strContents_in As String)

    Dim f As Integer

    f = FreeFile()
    Open strFilename_in For Output As #f
    Print #f, strContents_in
    Close #f

End Sub


Private Function GetResourceFileAsString(ByVal lngResID_in As Long) As String

    Dim a_bytFile() As Byte
    Dim lngIdx As Long
    Dim lngLastByte As Long

    GetResourceFileAsString = ""
    a_bytFile = LoadResData(lngResID_in, "CUSTOM")
    lngLastByte = UBound(a_bytFile)
    For lngIdx = 0 To lngLastByte
        GetResourceFileAsString = GetResourceFileAsString & Chr$(a_bytFile(lngIdx))
    Next lngIdx

End Function
Private Sub Class_Terminate()

    'remove links to external objects
    Set m_xobjFunc = Nothing
    Set m_xobjRegistry = Nothing

End Sub
Public Property Get TotalEmulators() As Long

    TotalEmulators = m_lngTotalEmulators

End Property
Public Function EmulatorPathExe(ByVal lngEmulatorIdx_in As Long) As String

    EmulatorPathExe = ma_udtEmulators(lngEmulatorIdx_in).strFullPathAndFilename

End Function

Public Function EmulatorTotalFiles(ByVal lngEmulatorIdx_in As Long) As Long

    EmulatorTotalFiles = ma_udtEmulators(lngEmulatorIdx_in).lngNumFiles

End Function
Public Function EmulatorUseShort(ByVal lngEmulatorIdx_in As Long) As Boolean

    EmulatorUseShort = ma_udtEmulators(lngEmulatorIdx_in).blnUseShortFilenames

End Function
Public Function EmulatorScript(ByVal lngEmulatorIdx_in As Long) As String

    EmulatorScript = ma_udtEmulators(lngEmulatorIdx_in).strScript

End Function
Public Function EmulatorTypes(ByVal lngEmulatorIdx_in As Long) As String

    EmulatorTypes = ma_udtEmulators(lngEmulatorIdx_in).strTypes

End Function
Public Function EmulatorWaits(ByVal lngEmulatorIdx_in As Long) As Long

    EmulatorWaits = ma_udtEmulators(lngEmulatorIdx_in).lngWait

End Function
Public Function EmulatorUseDefaultTypes(ByVal lngEmulatorIdx_in As Long) As Boolean

    EmulatorUseDefaultTypes = ma_udtEmulators(lngEmulatorIdx_in).blnUseDefaultTypes

End Function
Public Function EmulatorName(ByVal lngEmulatorIdx_in As Long) As String

    EmulatorName = ma_udtEmulators(lngEmulatorIdx_in).strName

End Function
Public Property Get TotalFileTypes() As Long

    TotalFileTypes = m_lngTotalEmulatorFileTypes

End Property
Public Function FileType(ByVal lngFileTypeIdx_in As Long) As String

    FileType = ma_strEmulatorFileTypes(lngFileTypeIdx_in)

End Function
Public Sub GetFileTypeAsArray(a_strFileTypes_out() As String)

    'returns the emulator filetypes as an array
    
    Dim lngFTIdx As Long

    ReDim a_strFileTypes_out(0)

    For lngFTIdx = 0 To (m_lngTotalEmulatorFileTypes - 1)
        ReDim Preserve a_strFileTypes_out(lngFTIdx)
        a_strFileTypes_out(lngFTIdx) = ma_strEmulatorFileTypes(lngFTIdx)
    Next lngFTIdx

End Sub

Private Function CheckOSVersion() As Long
    Dim tOS As OSVERSIONINFO
    tOS.dwOSVersionInfoSize = Len(tOS)
    Call GetVersionEx(tOS)
    CheckOSVersion = tOS.dwPlatformId
End Function
Private Function GetRunningProcessIDsByFilename(ByVal strFilename As String, a_lngProcessIDs() As Long) As Long

    'original source taken from here (thanks to Wokawidget):
    'http://www.vbforums.com/showpost.php?p=1878673&postcount=8

    Dim aPID() As Long
    Dim lProcesses As Long
    Dim lProcess As Long
    Dim lModule As Long
    Dim sName As String
    Dim iIndex As Integer
    Dim bCopied As Long
    Dim lSnapShot As Long
    Dim tPE As PROCESSENTRY32
    Dim blnIsFullPathFilename As Boolean

    GetRunningProcessIDsByFilename = 0
    ReDim a_lngProcessIDs(0)
    blnIsFullPathFilename = (InStr(strFilename, "\") > 0)

    On Error GoTo Err_Handler

    If CheckOSVersion() = VER_PLATFORM_WIN32_WINDOWS Then
        'Windows 9x unsupported - c'mon people get with the program
    Else
        'Windows NT
        'The EnumProcesses Function doesn't indicate how many Process there are,
        'so you need to pass a large array and trim off the empty elements
        'as cbNeeded will return the no. of Processes copied.
        ReDim aPID(255)
        Call EnumProcesses(aPID(0), 1024, lProcesses)
        lProcesses = lProcesses / 4
        ReDim Preserve aPID(lProcesses)

        For iIndex = 0 To lProcesses - 1
            'Get the Process Handle, by Opening the Process
            lProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aPID(iIndex))
            If lProcess Then
                'Just get the First Module, all we need is the Handle to get
                'the Filename..
                If EnumProcessModules(lProcess, lModule, 4, 0&) Then
                    sName = Space(260)
                    Call GetModuleFileNameExA(lProcess, lModule, sName, Len(sName))
                    If Not blnIsFullPathFilename Then
                        If InStr(sName, "\") > 0 Then
                            sName = Mid(sName, InStrRev(sName, "\") + 1)
                        End If
                    End If
                    If InStr(sName, Chr(0)) Then
                        sName = Left(sName, InStr(sName, Chr(0)) - 1)
                    End If
                    If StrComp(strFilename, sName, vbTextCompare) = 0 Then
                        ReDim Preserve a_lngProcessIDs(GetRunningProcessIDsByFilename)
                        a_lngProcessIDs(GetRunningProcessIDsByFilename) = aPID(iIndex)
                        GetRunningProcessIDsByFilename = GetRunningProcessIDsByFilename + 1
                    End If
                End If
                'Close the Process Handle
                CloseHandle lProcess
            End If
        Next
    End If

Err_Handler:
    
End Function
