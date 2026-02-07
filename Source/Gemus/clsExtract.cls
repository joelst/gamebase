VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtract"
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

'==== NOTE: THIS CLASS IS NOT USED FROM GAMEBASE v1.3 ONWARDS ====
'==================== USES "clsUnpack" INSTEAD ===================

Option Explicit

Private m_blnRezipLegacy As Boolean

Public Enum GB64ExtractResultEnum
    e_GB64ExtractResult_OK
    e_GB64ExtractResult_NoFile
    e_GB64ExtractResult_UnexpectedError
    e_GB64ExtractResult_NotCompatible
    e_GB64ExtractResult_Corrupt
    e_GB64ExtractResult_NoneSelected
End Enum
Private Enum GB64RecursiveStateEnum
    e_GB64RecursiveState_GotFile
    e_GB64RecursiveState_KeepGoing
    e_GB64RecursiveState_NotValid
End Enum
Private FileContents As FileListBox
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_eState As GB64RecursiveStateEnum
Private m_lngPass As Long
Private m_strFileList As String
Private m_strFilename As String
Private m_strError As String
Private m_strAppPath As String
Private m_strAppPathWBS As String
Private m_strRezipDir As String
Private m_strPassDir As String
Private m_strEmulatorPattern As String
Private m_strCommandCom As String

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

Public Function GetExtractorPath(strFilename_in As String) As String

    If m_xobjFunc.FileExists(m_strAppPathWBS & strFilename_in, False, "") = True Then
        GetExtractorPath = m_xobjFunc.Long2Short(m_strAppPathWBS & strFilename_in)
    ElseIf m_xobjFunc.FileExists(strFilename_in, False, "") = True Then
        GetExtractorPath = strFilename_in
    Else
        GetExtractorPath = ""
    End If

End Function

Public Function IsFileMatchingFileTypes(strFilename_in As String, a_strFileTypes_in() As String) As Boolean

    On Error GoTo Err_Handler

    Dim lngDotPos As Long
    Dim lngLen As Long
    Dim strExt As String
    Dim strFileTypes As String

    lngDotPos = InStrRev(strFilename_in, ".")
    lngLen = Len(strFilename_in)
    If lngLen > lngDotPos Then
        strExt = UCase$(Right$(strFilename_in, lngLen - lngDotPos))
        strFileTypes = ";" & UCase$(Join(a_strFileTypes_in, ";")) & ";"
        IsFileMatchingFileTypes = (InStr(strFileTypes, ";" & strExt & ";") > 0)
        Exit Function
    End If

Err_Handler:
    IsFileMatchingFileTypes = False

End Function

Public Sub SetReZipLegacy(ByVal blnUseLegacy_in As Boolean)

    m_blnRezipLegacy = blnUseLegacy_in

End Sub

Public Function SevenZipWalk(strArchivePathFile As String, strTempPathFile As String, _
                                lngTotalFiles_out As Long, a_strWalked_out() As String) As String

    On Error GoTo Err_Handler

    Dim str7zExe As String
    Dim strArchivePath As String
    Dim strArchiveFile As String
    Dim strTempPath As String
    Dim strTempFile As String
    Dim f As Integer
    Dim strLine As String
    Dim lngStartPos As Long
    Dim lngAttrPos As Long
    Dim blnGotStartPos As Boolean
    Dim lngNumLinesWithDashes As Long
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strEntry As String

    'get path to 7z.exe
    str7zExe = GetExtractorPath("7z.exe")
    If str7zExe = "" Then
        SevenZipWalk = GetINIMessage("Extracting", "Error_No7z", "Path to 7z.exe not found, unable to check contents of archive") & " " & strArchivePathFile
        Exit Function
    End If

    'create temp file path if it doesn't exist
    Call m_xobjFunc.GetFileAndPath(strTempPathFile, strTempPath, strTempFile)
    If m_xobjFunc.DirExists(strTempPath) = False Then Call MkDir(strTempPath)
    
    'delete temporary file if it does exist
    Call m_xobjFunc.DeletePathSilent(strTempPathFile)

    'list contents of archive into the temp file
    Call m_xobjFunc.GetFileAndPath(strArchivePathFile, strArchivePath, strArchiveFile)
    ChDrive strArchivePath
    ChDir strArchivePath
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    '  l    list contents of archive
    ' -y    answer yes on all queries
    ' (7-zip help file recommends using long filenames, as short [8.3] ones may not work 100% of the time)
    strCommandLine = m_strCommandCom & str7zExe & " l """ & strArchiveFile & """ -y > """ & strTempPathFile
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

    'check for existence of the file and read in its contents
    If m_xobjFunc.FileExists(strTempPathFile, False, "") = False Then
        SevenZipWalk = GetINIMessage("Extracting", "Error_TempFile", "Temporary list file could not be created. Unable to list contents of archive") & " " & strArchivePathFile
        Exit Function
    End If
    f = FreeFile()
    Open strTempPathFile For Input As #f
    blnGotStartPos = False
    lngNumLinesWithDashes = 0
    ReDim a_strWalked_out(0)
    lngTotalFiles_out = 0
    Do Until EOF(f)
        Line Input #f, strLine
        Call Trim$(strLine)
        If Not blnGotStartPos Then
            'look for line with "Date", "Attr", "Size" and "Name" in it
            lngStartPos = InStr(strLine, "Name")
            lngAttrPos = InStr(strLine, "Attr") - 1 'right justified, data starts one column before
            If InStr(strLine, "Date") > 0 And InStr(strLine, "Size") > 0 And lngStartPos > 0 And lngAttrPos > 0 Then
                blnGotStartPos = True
            End If
        ElseIf Left$(strLine, 1) = "-" Then
            'line with dashes (borders of file list)
            lngNumLinesWithDashes = lngNumLinesWithDashes + 1
        ElseIf lngNumLinesWithDashes = 1 Then
            'should be line containing a file or folder
            If Mid$(strLine, lngAttrPos, 1) <> "D" Then
                'not a directory
                strEntry = Trim(Right$(strLine, Len(strLine) - lngStartPos + 1))
                If strEntry <> "" Then
                    ReDim Preserve a_strWalked_out(lngTotalFiles_out)
                    a_strWalked_out(lngTotalFiles_out) = strEntry
                    lngTotalFiles_out = lngTotalFiles_out + 1
                End If
            End If
        End If
    Loop
    Close #f

    'no errors
    SevenZipWalk = ""
    Call m_xobjFunc.DeletePathSilent(strTempPathFile)
    Exit Function

Err_Handler:
    SevenZipWalk = GetINIMessage("Extracting", "Error_Unexpected", "Unexpected Error. Unable to list contents of archive") & " " & strArchivePathFile _
        & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")"

End Function

Private Sub UnLha(strExtractorPath_in As String, strPath_in As String, strFile_in As String, strDestPath_in As String)
    
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strShortRawFilename As String
    Dim strShortFullFile As String
    Dim strShortPath As String
    'UnLHA to new Dir
    ChDrive strPath_in
    ChDir strPath_in
    
    'Run LHA.EXE.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    With m_xobjFunc
        strShortFullFile = .Long2Short(strPath_in & "\" & strFile_in)
        Call m_xobjFunc.GetFileAndPath(strShortFullFile, strShortPath, strShortRawFilename)
    End With
    strCommandLine = m_strCommandCom & strExtractorPath_in & " e /m " & strShortRawFilename & " " & strDestPath_in & "\"
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

End Sub

Public Sub SevenZip(strExtractorPath_in As String, strPath_in As String, strFile_in As String, strDestPath_in As String)
    
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strShortRawFilename As String
    Dim strShortFullFile As String
    Dim strShortPath As String

    'extract to new dir
    ChDrive strPath_in
    ChDir strPath_in
    
    'Run 7z.exe.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    '  x    extract with full paths
    ' -y    answer yes on all queries
    ' -r    recurse directories
    ' -o    output folder
    ' (7-zip help file recommends using long filenames, as short [8.3] ones may not work 100% of the time)
    strCommandLine = m_strCommandCom & strExtractorPath_in & " x """ & strFile_in & """ -yro""" & strDestPath_in & "\"""
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

End Sub
Public Function Extract(strFileToExtract_in As String, strDefaultFile_inout As String, a_strFileTypes_in() As String, ByVal lngTotalFileTypes_in As Long, strDestFolder_in As String, strReturnFilename_out As String) As String

    On Error GoTo Err_Handler
    
    Dim strJustFile As String
    Dim strJustPath As String
    Dim strDestFolderWBS As String
    Dim eResult As GB64ExtractResultEnum
    Dim lngIdx As Long
    Dim strStatusArchiveContents As String
    Dim lngDotPos As Long
    Dim strExt As String
    Dim blnRecognised As Boolean
    Dim strUpper As String

    Extract = ""

    '-----------------------------
    'do we have a valid game file?
    '-----------------------------

    lngDotPos = InStrRev(strFileToExtract_in, ".")
    If lngDotPos > 0 Then
        If Right$(strFileToExtract_in, 1) <> "." Then
            blnRecognised = False
            strExt = UCase$(Right$(strFileToExtract_in, Len(strFileToExtract_in) - lngDotPos))
            strUpper = UCase(strExt)
            For lngIdx = 0 To (lngTotalFileTypes_in - 1)
                If strUpper = a_strFileTypes_in(lngIdx) Then
                    'recognised
                    blnRecognised = True
                    Exit For
                End If
            Next lngIdx
            If blnRecognised = True Then

                '---------------
                'valid game file
                '---------------

                strReturnFilename_out = strFileToExtract_in

            ElseIf (strExt = "LNX") Or (strExt = "LHA") Or (strExt = "LZH") _
                Or (strExt = "ZIP") Or (strExt = "GZ") Or (strExt = "GZIP") Or (strExt = "7Z") Or (strExt = "TGZ") _
                Or (strExt = "BZ2") Or (strExt = "BZIP2") Or (strExt = "TBZ2") Or (strExt = "TBZ") Or (strExt = "TAR") _
                Or (strExt = "LZMA") Or (strExt = "RAR") Or (strExt = "CAB") Or (strExt = "ARJ") Or (strExt = "Z") _
                Or (strExt = "TAZ") Or (strExt = "ISO") _
            Then

                '--------------------------
                'game file needs extracting
                '--------------------------

                'initial test
                If lngTotalFileTypes_in = 0 Then
                    Extract = GetINIMessage("Running Games", "Status_NotCompatible", "This archive contains no compatible emulator files.")
                    Exit Function
                End If

                'set initial variables for the recursive routine
                m_lngPass = 0
                m_strFileList = ""
                m_strFilename = ""
                m_strError = ""
                m_eState = e_GB64RecursiveState_KeepGoing
                m_strAppPath = App.Path
                If g_blnInVBIDE Then
                    m_strAppPath = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
                End If
                m_strAppPathWBS = m_strAppPath
                If Right$(m_strAppPathWBS, 1) <> "\" Then m_strAppPathWBS = m_strAppPathWBS & "\"

                g_strDefaultFile = strDefaultFile_inout
            
                m_strEmulatorPattern = ""
                For lngIdx = 0 To (lngTotalFileTypes_in - 1)
                    m_strEmulatorPattern = m_strEmulatorPattern & "*." & a_strFileTypes_in(lngIdx) & ";"
                Next lngIdx
                m_strEmulatorPattern = Left$(m_strEmulatorPattern, Len(m_strEmulatorPattern) - 1)

                g_strArchiveCaption = GetINIMessage("Running Games", "ArchiveCaption", "Select file in archive")
                g_strArchivePrompt = GetINIMessage("Running Games", "ArchivePrompt", "This archive file has more than one compatible game file.  Select the file to load from the list below")
                g_strOKButton = GetINIMessage("General Buttons", "OK", "&OK")
                g_strCancelButton = GetINIMessage("General Buttons", "Cancel", "&Cancel")
                Load frmGB64Main
                Set FileContents = frmGB64Main.FileBox

                'set up gbgame folder and copy file into there for working on
                Call m_xobjFunc.GetFileAndPath(strFileToExtract_in, strJustPath, strJustFile)
                If m_xobjFunc.DirExists(strDestFolder_in & "\P1") = True Then
                    Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\P1")
                End If
                Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\ReZip")
                If m_xobjFunc.DirExists(strDestFolder_in) = False Then
                    Call MkDir(strDestFolder_in)
                End If
                strDestFolderWBS = strDestFolder_in
                If Right$(strDestFolderWBS, 1) <> "\" Then strDestFolderWBS = strDestFolderWBS & "\"
                m_strRezipDir = strDestFolderWBS & "ReZip\"
                FileCopy strFileToExtract_in, strDestFolderWBS & strJustFile

                'extract the game and return the info
                g_blnRezippable = False
                eResult = ExtractAgain(strDestFolder_in)
                strReturnFilename_out = m_strFilename
                strDefaultFile_inout = g_strDefaultFile

                If eResult <> e_GB64ExtractResult_OK Then
                    strStatusArchiveContents = GetINIMessage("Running Games", "Status_ArchiveContents", "Archive contents") & ":" & vbCrLf
                    If eResult = e_GB64ExtractResult_NotCompatible Then
                        Extract = GetINIMessage("Running Games", "Status_NotCompatible", "This archive contains no compatible emulator files.") & IIf((m_strFileList = ""), "", vbCrLf & vbCrLf & strStatusArchiveContents & m_strFileList)
                    ElseIf eResult = e_GB64ExtractResult_Corrupt Then
                        Extract = GetINIMessage("Running Games", "Status_NonConform", "There was an error reading from the file or the file does not conform to the standard.") & IIf((m_strFileList = ""), "", vbCrLf & vbCrLf & strStatusArchiveContents & m_strFileList)
                    ElseIf eResult = e_GB64ExtractResult_NoFile Then
                        Extract = GetINIMessage("Running Games", "Status_NoFile", "To extract this archive you need to put the following file in your GB64 folder, or somewhere on your %path%") & ":" & vbCrLf & vbCrLf & m_strError & IIf((m_strFileList = ""), "", vbCrLf & vbCrLf & strStatusArchiveContents & m_strFileList)
                    ElseIf eResult = e_GB64ExtractResult_NoneSelected Then
                        Extract = GetINIMessage("Running Games", "Status_NoneSelected", "You did not select a compatible emulator file from this archive.") & IIf((m_strFileList = ""), "", vbCrLf & vbCrLf & strStatusArchiveContents & m_strFileList)
                    Else 'unexpected
                        Extract = GetINIMessage("Running Games", "Status_UnExpected", "There was an unexpected error whilst reading from this archive.") & vbCrLf & vbCrLf & m_strError
                    End If
                End If

            Else
                'incompatible extension
                Extract = GetINIMessage("Running Games", "Status_InvalidType", "The file you selected is not a recognised emulator file.")
            End If
        Else
            'dot at end of filename - no extension
            Extract = GetINIMessage("Running Games", "Status_InvalidType", "The file you selected is not a recognised emulator file.")
        End If
    Else
        'no dot in filename
        Extract = GetINIMessage("Running Games", "Status_InvalidType", "The file you selected is not a recognised emulator file.")
    End If

    Exit Function
    
Err_Handler:
    Extract = GetINIMessage("Running Games", "Status_UnExpected", "There was an unexpected error whilst reading from this archive.") & vbCrLf & vbCrLf & Err.Description & " (Err " & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")"

End Function

Private Function GetINIMessage(strSection_in As String, strKey_in As String, strDefault_in As String) As String

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
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, m_strAppPathWBS & "GemusText.ini")
    GetINIMessage = Left$(strReturned, lngRes)

End Function
Private Function ExtractAgain(strDestFolder_in As String) As GB64ExtractResultEnum

    'recursive routine to get the valid filename for shelling into the emulator
    
    Dim lngFileIdx As Long              'loop index
    Dim strExtractorPath As String      'extractor path
    Dim strTempFile As String
    Dim strTempPath As String
    Dim strRawFile As String
    Dim strUpperFile As String
    Dim strDirFile As String
    Dim strExt As String
    Dim strFileToExtract As String

    On Error GoTo Err_Handler
    
    If m_eState = e_GB64RecursiveState_KeepGoing Then 'valid/invalid file not found yet
    
        'increase Directory Level
        m_lngPass = m_lngPass + 1
        m_strPassDir = "P" & Trim$(m_lngPass)
              
        With FileContents
        
            '**********************************************
            '* FIRST TEST IF FILES IN DIRECTORY ARE VALID *
            '**********************************************
            
            .Path = strDestFolder_in
            .Pattern = m_strEmulatorPattern
            .Refresh
            
            If .ListCount = 1 Then
                'got the file
                m_strFilename = m_strRezipDir & .List(0)
                MoveToRezipFolder strDestFolder_in
                m_eState = e_GB64RecursiveState_GotFile
                ExtractAgain = e_GB64ExtractResult_OK
                If (m_lngPass <> 2) Then
                    g_blnRezippable = False
                End If
                If g_blnRezippable = True And m_blnRezipLegacy = True Then
                    'must be no subfolders in the zip
                    strDirFile = Dir(m_strRezipDir & "*.*")
                    Do While strDirFile <> ""
                        If strDirFile <> "." And strDirFile <> ".." Then
                            If GetAttr(m_strRezipDir & strDirFile) And vbDirectory Then
                                g_blnRezippable = False
                                Exit Do
                            End If
                        End If
                        strDirFile = Dir()
                    Loop
                End If
                GoTo NextPass
                
            ElseIf .ListCount > 1 Then 'more than one valid file
        
                'is the specified file there?
                If g_strDefaultFile <> "" Then
                    strUpperFile = UCase$(g_strDefaultFile)
                    g_strDefaultFile = ""
                    For lngFileIdx = 0 To (.ListCount - 1)
                        If UCase$(.List(lngFileIdx)) = strUpperFile Then
                            g_strDefaultFile = .List(lngFileIdx)
                            Exit For
                        End If
                    Next lngFileIdx
                End If
                'show form to pick default file?
                If g_strDefaultFile = "" Then
                    With frmGB64Main.lstFiles
                        For lngFileIdx = 0 To FileContents.ListCount - 1
                            .AddItem FileContents.List(lngFileIdx)
                        Next lngFileIdx
                        .ListIndex = 0
                    End With
                    Call frmGB64Main.Show(vbModal)
                End If
                'do we have a default file?
                If g_strDefaultFile = "" Then
                    ExtractAgain = e_GB64ExtractResult_NoneSelected
                    m_strFileList = FilesInTempDIR(strDestFolder_in)
                    m_eState = e_GB64RecursiveState_NotValid
                    GoTo NextPass
                Else
                    'got the file
                    m_strFilename = m_strRezipDir & g_strDefaultFile
                    MoveToRezipFolder strDestFolder_in
                    m_eState = e_GB64RecursiveState_GotFile
                    ExtractAgain = e_GB64ExtractResult_OK
                    If (m_lngPass <> 2) Then
                        g_blnRezippable = False
                    End If
                    If g_blnRezippable = True And m_blnRezipLegacy = True Then
                        'must be no subfolders in the zip
                        strDirFile = Dir(m_strRezipDir & "*.*")
                        Do While strDirFile <> ""
                            If strDirFile <> "." And strDirFile <> ".." Then
                                If GetAttr(m_strRezipDir & strDirFile) And vbDirectory Then
                                    g_blnRezippable = False
                                    Exit Do
                                End If
                            End If
                            strDirFile = Dir()
                        Loop
                    End If
                    GoTo NextPass
                End If

            End If


            '********************************************************************
            '* NO VALID FILES IN DIRECTORY - TEST WHAT KIND OF FILES ARE THERE! *
            '********************************************************************
            
            '=============
            'LHA/LZH FILES
            '=============

            .Pattern = "*.lha;*.lzh" 'VALID if 1 or Pick First lha/lzh file
            .Refresh
            If .ListCount > 0 Then
                strExtractorPath = GetExtractorPath("lha.exe")
                If strExtractorPath <> "" Then
                    Call MkDir(strDestFolder_in & "\" & m_strPassDir)
                    Call UnLha(strExtractorPath, strDestFolder_in, .List(0), m_strPassDir)
                    'delete the zipped file
                    Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & .List(0))
                    'go again
                    ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                    GoTo NextPass
                Else
                    m_strError = "LHA.EXE"
                    m_eState = e_GB64RecursiveState_NotValid
                    ExtractAgain = e_GB64ExtractResult_NoFile
                    GoTo NextPass
                End If

            End If

            '==========
            'LYNX FILES
            '==========

            .Pattern = "*.lnx"
            .Refresh
            If .ListCount > 0 Then
                strExtractorPath = GetExtractorPath("starlynx.exe")
                If strExtractorPath <> "" Then
                    Call MkDir(strDestFolder_in & "\" & m_strPassDir)
                    For lngFileIdx = 0 To .ListCount - 1
                        strTempFile = m_xobjFunc.Long2Short(strDestFolder_in & "\" & .List(lngFileIdx))
                        Call m_xobjFunc.GetFileAndPath(strTempFile, strTempPath, strRawFile)
                        Call UnLynx(strExtractorPath, m_strPassDir, strDestFolder_in, strRawFile)
                        'delete the starlynx'd file
                        Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & .List(lngFileIdx))
                    Next lngFileIdx
                    ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                    GoTo NextPass
                Else
                    m_strError = "STARLYNX.EXE"
                    m_eState = e_GB64RecursiveState_NotValid
                    ExtractAgain = e_GB64ExtractResult_NoFile
                    GoTo NextPass
                End If
            End If

            '=========================
            'DISKPACKED ZIPCODED FILES
            '=========================

            .Pattern = "1!*.*"
            .Refresh
            If .ListCount > 0 Then
                strExtractorPath = GetExtractorPath("starzip.exe")
                If strExtractorPath <> "" Then
                    Call MkDir(strDestFolder_in & "\" & m_strPassDir)
                    For lngFileIdx = 0 To .ListCount - 1
                        strTempFile = m_xobjFunc.Long2Short(strDestFolder_in & "\" & .List(lngFileIdx))
                        Call m_xobjFunc.GetFileAndPath(strTempFile, strTempPath, strRawFile)
                        Call UnZipcode(strExtractorPath, m_strPassDir, strDestFolder_in, strRawFile)
                        'delete the zipcoded file
                        Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & .List(lngFileIdx))
                    Next lngFileIdx
                    ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                    GoTo NextPass
                Else
                    m_strError = "STARLYNX.EXE"
                    m_eState = e_GB64RecursiveState_NotValid
                    ExtractAgain = e_GB64ExtractResult_NoFile
                    GoTo NextPass
                End If
            End If

            '=========================
            'FILEPACKED ZIPCODED FILES
            '=========================

            .Pattern = "A!*.*"
            .Refresh
            If .ListCount > 0 Then
                strExtractorPath = GetExtractorPath("starzip.exe")
                If strExtractorPath <> "" Then
                    Call MkDir(strDestFolder_in & "\" & m_strPassDir)
                    For lngFileIdx = 0 To .ListCount - 1
                        strTempFile = m_xobjFunc.Long2Short(strDestFolder_in & "\" & .List(lngFileIdx))
                        Call m_xobjFunc.GetFileAndPath(strTempFile, strTempPath, strRawFile)
                        Call UnZipcode(strExtractorPath, m_strPassDir, strDestFolder_in, strRawFile)
                        'delete the filepacked file
                        Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & .List(lngFileIdx))
                    Next lngFileIdx
                    ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                    GoTo NextPass
                Else
                    m_strError = "STARLYNX.EXE"
                    m_eState = e_GB64RecursiveState_NotValid
                    ExtractAgain = e_GB64ExtractResult_NoFile
                    GoTo NextPass
                End If
            End If

            '==================
            'ZIP FILES (LEGACY)
            '==================

            If m_blnRezipLegacy = True Then
                .Pattern = "*.zip"
                .Refresh
                If .ListCount > 0 Then 'VALID if 1 or Pick First Zip file
                    MkDir strDestFolder_in & "\" & m_strPassDir
                    g_blnZipError = False
                    Call UnzipToTempDIR(strDestFolder_in & "\" & m_strPassDir, strDestFolder_in & "\" & .List(0))
                    If g_blnZipError = True Then
                        'error unzipping
                        ExtractAgain = e_GB64ExtractResult_Corrupt
                        m_eState = e_GB64RecursiveState_NotValid
                        GoTo NextPass
                    Else
                        'delete the zipped file
                        Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & .List(0))
                        'see what was unzipped!
                        If (m_lngPass = 1) Then
                            g_blnRezippable = True
                        Else
                            g_blnRezippable = False
                        End If
                        ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                        GoTo NextPass
                    End If
                End If
            End If

            '========================
            '7-ZIP SUPPORTED ARCHIVES
            '========================

             'VALID if 1 or Pick First file
            .Pattern = "*.7z;*.gz;*.gzip;*.tgz;*.bz2;*.bzip2;*.tbz2;*.tbz;*.tar;*.lzma;*.rar;*.cab;*.arj;*.z;*.taz;*.iso"
            If m_blnRezipLegacy = False Then .Pattern = .Pattern & ";*.zip"
            .Refresh
            If .ListCount > 0 Then
                strExtractorPath = GetExtractorPath("7z.exe")
                If strExtractorPath <> "" Then
                    Call MkDir(strDestFolder_in & "\" & m_strPassDir)
                    strFileToExtract = .List(0)
                    Call SevenZip(strExtractorPath, strDestFolder_in, strFileToExtract, m_strPassDir)
                    'delete the zipped file
                    Call m_xobjFunc.DeletePathSilent(strDestFolder_in & "\" & strFileToExtract)
                    If (m_lngPass = 1) Then
                        strExt = UCase$(Right$(strFileToExtract, Len(strFileToExtract) - InStrRev(strFileToExtract, ".")))
                        'supported archives: .7z .zip .gz .gzip .tgz .bz2 .bzip2 .tbz2 .tbz .tar
                        If (strExt = "7Z") Or (strExt = "ZIP") Or (strExt = "GZ") Or (strExt = "GZIP") Or (strExt = "TGZ") Or _
                            (strExt = "BZ2") Or (strExt = "BZIP2") Or (strExt = "TBZ2") Or (strExt = "TBZ") Or (strExt = "TAR") _
                        Then
                            g_blnRezippable = True
                        Else
                            g_blnRezippable = False
                        End If
                    Else
                        g_blnRezippable = False
                    End If
                    'go again
                    ExtractAgain = ExtractAgain(strDestFolder_in & "\" & m_strPassDir)
                    GoTo NextPass
                Else
                    m_strError = "7Z.EXE"
                    m_eState = e_GB64RecursiveState_NotValid
                    ExtractAgain = e_GB64ExtractResult_NoFile
                    GoTo NextPass
                End If

            End If

            '=========
            'NOT VALID
            '=========

            'if we got here then the current folder contains no valid game files for the emulators
            ExtractAgain = e_GB64ExtractResult_NotCompatible
            If .ListCount > 0 Then
                m_strFileList = FilesInTempDIR(strDestFolder_in)
            End If
            m_eState = e_GB64RecursiveState_NotValid
            GoTo NextPass

        End With

    End If

Err_Handler:

    m_strError = Err.Description & " (Err " & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")"
    ExtractAgain = e_GB64ExtractResult_UnexpectedError
    m_eState = e_GB64RecursiveState_NotValid

NextPass:

End Function
Private Sub UnLynx(strExtractorPath_in As String, strDestPath_in As String, strLynxPath_in As String, strLynxFile_in As String)
    
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strShortRawFilename As String
    Dim strDestFile As String
    Dim strTempPath As String
    Dim strSourceFile As String

    strShortRawFilename = m_xobjFunc.Long2Short(strLynxPath_in & "\" & strLynxFile_in)
    Call m_xobjFunc.GetFileAndPath(strShortRawFilename, strTempPath, strDestFile)
    strSourceFile = strDestFile
    strDestFile = Left$(strDestFile, Len(strDestFile) - 4)

    ChDrive strLynxPath_in
    ChDir strLynxPath_in

    'Run STARLYNX.EXE.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    strCommandLine = m_strCommandCom & strExtractorPath_in & " -x " & strSourceFile & " /y " & strDestPath_in & "\" & strDestFile
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

End Sub


Private Sub UnZipcode(strExtractorPath_in As String, strDestPath_in As String, strZipPath_in As String, strZipFile_in As String)
    
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strShortRawFilename As String
    Dim strDestFile As String
    Dim strTempPath As String
    Dim strSourceFile As String
    Dim strZipName As String

    strShortRawFilename = m_xobjFunc.Long2Short(strZipPath_in & "\" & strZipFile_in)
    Call m_xobjFunc.GetFileAndPath(strShortRawFilename, strTempPath, strDestFile)
    strSourceFile = strDestFile
    If InStr(strSourceFile, "!") = Len(strSourceFile) Then
        Exit Sub
    Else
        strZipName = Right$(strSourceFile, Len(strSourceFile) - InStr(strSourceFile, "!"))
    End If
    strDestFile = Left$(strDestFile, Len(strDestFile) - 4)

    ChDrive strZipPath_in
    ChDir strZipPath_in

    'Run STARZIP.EXE.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    strCommandLine = m_strCommandCom & strExtractorPath_in & " -x " & strZipName & " /y " & strDestPath_in & "\"
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

End Sub
Private Sub MoveToRezipFolder(strFromDir_in As String)

    On Error GoTo Whoops

'    Dim intFileIdx As Integer
'    Dim strFromDir As String
'
'    strFromDir = strFromDir_in
'    With FileContents
'        .Path = strFromDir
'        .Pattern = "*.*"
'        .Refresh
'        If m_xobjFunc.DirExists(m_strRezipDir) = False Then
'            MkDir m_strRezipDir
'        End If
'        For intFileIdx = 0 To .ListCount - 1
'            If m_xobjFunc.FileExists(m_strRezipDir & .List(intFileIdx), False, "") = False Then
'                Name (strFromDir_in & "\" & .List(intFileIdx)) As (m_strRezipDir & .List(intFileIdx))
'            End If
'        Next intFileIdx
'        Call m_xobjFunc.DeletePathSilent(strFromDir)
'    End With

    If m_xobjFunc.DirExists(m_strRezipDir) Then
        Call m_xobjFunc.DeletePathSilent(m_strRezipDir)
    End If
    Name (strFromDir_in) As (m_strRezipDir)

    Exit Sub

Whoops:

End Sub

Private Function FilesInTempDIR(ByVal strDirectory_in As String) As String

    Dim strTempFile As String

    FilesInTempDIR = ""
    strTempFile = Dir(strDirectory_in & "\*.*")
    Do While strTempFile <> ""
        FilesInTempDIR = FilesInTempDIR & strTempFile & vbCrLf
        strTempFile = Dir
    Loop

End Function

Private Sub Class_Initialize()

    Set m_xobjFunc = New GBFuncs.clsFunctions
    m_strCommandCom = m_xobjFunc.Long2Short(Environ$("COMSPEC")) & " /c "
    m_blnRezipLegacy = False

    m_strAppPath = App.Path
    If g_blnInVBIDE Then
        m_strAppPath = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    m_strAppPathWBS = m_strAppPath
    If Right$(m_strAppPathWBS, 1) <> "\" Then m_strAppPathWBS = m_strAppPathWBS & "\"
    
End Sub

Private Sub Class_Terminate()

    Set m_xobjFunc = Nothing
    Unload frmGB64Main
    Set FileContents = Nothing
    
End Sub
