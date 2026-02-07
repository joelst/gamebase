VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnpack"
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

Private m_xobjFunc As GBFuncs.clsFunctions
Private WithEvents m_xobjFuncPre As GBFuncs.clsFunctions
Attribute m_xobjFuncPre.VB_VarHelpID = -1
Private WithEvents m_xobjFuncPost As GBFuncs.clsFunctions
Attribute m_xobjFuncPost.VB_VarHelpID = -1
Private WithEvents m_xobjFuncRecursive As GBFuncs.clsFunctions
Attribute m_xobjFuncRecursive.VB_VarHelpID = -1
Private m_objCRC As clsCRC

Private m_strAppPath As String
Private m_strAppPathWBS As String
Private m_strCommandCom As String

Private m_lngFinalFolderIdx As Long
Private m_strFinalFolder As String
Private m_strRunnableRelativePathFile As String

Private ma_strUnpackableFileTypes() As String
Private m_strUnpackableFileTypes As String

Private ma_strRepackableFileTypes() As String
Private m_strRepackableFileTypes As String

Private Type UnpackedFileSnapshot
    strCRC As String
    strRelativePathFile As String
    strUCaseRelativePathFile As String
End Type
Private ma_udtPreUnpackedFilesSnapshot() As UnpackedFileSnapshot
Private m_lngTotalPreUnpackedFilesSnapshot As Long
Private ma_udtPostUnpackedFilesSnapshot() As UnpackedFileSnapshot
Private m_lngTotalPostUnpackedFilesSnapshot As Long

Private ma_strRecursiveUnpackedFilesSnapshot() As String
Private m_lngTotalRecursiveUnpackedFilesSnapshot As Long

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

Private Function DidSnapshotsChangeDuringPlay() As Boolean

    'assumes both pre and post snapshots exist

    Dim lngPreIdx As Long
    Dim lngPostIdx As Long
    Dim lngFoundPostIdx As Long

    'assume they changed, unless everything is the same
    DidSnapshotsChangeDuringPlay = True

    If m_lngTotalPreUnpackedFilesSnapshot <> m_lngTotalPostUnpackedFilesSnapshot Then Exit Function

    For lngPreIdx = 0 To m_lngTotalPreUnpackedFilesSnapshot - 1
        With ma_udtPreUnpackedFilesSnapshot(lngPreIdx)
            lngFoundPostIdx = -1
            For lngPostIdx = 0 To m_lngTotalPostUnpackedFilesSnapshot - 1
                If .strUCaseRelativePathFile = ma_udtPostUnpackedFilesSnapshot(lngPostIdx).strUCaseRelativePathFile Then
                    lngFoundPostIdx = lngPostIdx
                    Exit For
                End If
            Next lngPostIdx
            If lngFoundPostIdx = -1 Then Exit Function
            If .strCRC <> ma_udtPostUnpackedFilesSnapshot(lngFoundPostIdx).strCRC Then Exit Function
        End With
    Next lngPreIdx

    'no changes
    DidSnapshotsChangeDuringPlay = False

End Function

Public Function GetSnapshotChangedState(eWithAdditionalFiles As EnumRezipAdditionalFiles, _
                        blnDidSnapshotChange_out As Boolean, strError_out As String) As Boolean

    GetSnapshotChangedState = False
    strError_out = ""

    'must have both snapshots
    If m_lngTotalPreUnpackedFilesSnapshot = 0 Or m_lngTotalPostUnpackedFilesSnapshot = 0 Then
        strError_out = GetINIMessage("Running Games", "NoSnapshots", "Snapshots of unpacked folder before/after playing do not exist.")
        Exit Function
    End If

    'success
    GetSnapshotChangedState = True

End Function

Private Function GenerateRepackListFileFromPostSnapshot(strRepackFolderWBS_in As String, strListFile_out As String, _
                                                        strError_out As String) As Boolean

    'uses the post snapshot to generate a 7-zip compatible list file
    'of files to repack into it's parent archive (which may be the main game archive)
    
    'assumes the post snapshot exists

    On Error GoTo Err_Handler

    Dim intFreeFileNum As Integer
    Dim lngIdx As Long
    Dim strListPathFile As String

    GenerateRepackListFileFromPostSnapshot = False
    strError_out = ""
    strListFile_out = "__gemus_repack.lst"
    strListPathFile = strRepackFolderWBS_in & strListFile_out

    'create the file
    Call m_xobjFunc.DeletePathSilent(strListPathFile)
    intFreeFileNum = FreeFile()
    Open strListPathFile For Output As intFreeFileNum
    For lngIdx = 0 To m_lngTotalPostUnpackedFilesSnapshot - 1
        Print #intFreeFileNum, ma_udtPostUnpackedFilesSnapshot(lngIdx).strRelativePathFile
    Next lngIdx
    Close #intFreeFileNum

    'success
    GenerateRepackListFileFromPostSnapshot = True
    Exit Function

Err_Handler:
    strError_out = Err.Description & " (" & Err.Number & "," & Trim$(Err.LastDllError) & ")"
    GenerateRepackListFileFromPostSnapshot = False

End Function

Public Function GetFinalFolder() As String

    GetFinalFolder = m_strFinalFolder

End Function


Public Function GetFinalFolderIdx() As Long

    GetFinalFolderIdx = m_lngFinalFolderIdx

End Function


Function GetNumFilesMatchingFileTypes(a_strFiles_in() As String, a_strFileTypes_in() As String, _
                                        a_strRunnableFiles_out() As String) As Long
    Dim lngIdx As Long
    Dim lngNumFiles As Long
    Dim lngNumOut As Long
    Dim strFile As String

    GetNumFilesMatchingFileTypes = 0
    ReDim a_strRunnableFiles_out(0)
    lngNumFiles = UBound(a_strFiles_in) + 1
    lngNumOut = 0

    For lngIdx = 0 To lngNumFiles - 1
        strFile = a_strFiles_in(lngIdx)
        If IsFileMatchingFileTypes(strFile, a_strFileTypes_in) Then
            GetNumFilesMatchingFileTypes = GetNumFilesMatchingFileTypes + 1
            ReDim Preserve a_strRunnableFiles_out(lngNumOut)
            a_strRunnableFiles_out(lngNumOut) = strFile
            lngNumOut = lngNumOut + 1
        End If
    Next lngIdx

End Function

Private Function GetRepackParentArchive(strPathFileWBS_in As String, ByVal lngFolderIdx_in As Long, _
                                        strParentPathFile_out As String, strError_out As String) As Boolean

    Dim intFreeFileNum As Integer
    Dim strFile As String
    
    'not successful yet!
    GetRepackParentArchive = False

    strFile = strPathFileWBS_in & "__gemus"
    If m_xobjFunc.FileExists(strFile, False, "GEMUS") = False Then
        strError_out = GetINIMessage("Running Games", "GemusFileNotFound", "Unable to repack. Could not find required __gemus file in unpacked folder") & " " & lngFolderIdx_in
        Exit Function
    End If
    
    'read in the parent archive relative path/filename
    intFreeFileNum = FreeFile()
    Open strFile For Input As #intFreeFileNum
    strParentPathFile_out = ""
    Do Until EOF(intFreeFileNum)
        Line Input #intFreeFileNum, strParentPathFile_out
    Loop
    Close #intFreeFileNum

    'it has gotta exist
    If strParentPathFile_out = "" Then
        strError_out = GetINIMessage("Running Games", "ParentArchiveReadError", "Unable to repack. Could not read parent archive from __gemus file in unpacked folder") & " " & lngFolderIdx_in
                            
        Exit Function
    ElseIf Not m_xobjFunc.FileExists(strPathFileWBS_in & strParentPathFile_out, False, "GEMUS") Then
        strError_out = GetINIMessage("Running Games", "ParentArchiveReadError2", "Unable to repack. Could not find parent archive read from __gemus file in unpacked folder") & " " & lngFolderIdx_in
        Exit Function
    End If

    'success
    strParentPathFile_out = strPathFileWBS_in & strParentPathFile_out
    GetRepackParentArchive = True
    
End Function

Public Function GetRunnableFileInArchiveFromUser(strArchivePathFile_in As String, strTempPathWBS_in As String, _
                                    a_strRunnableFileTypes_in() As String, strRunnable_out As String) As Boolean
    With g_udtArchiveBrowse
        .strArchivePathFile = strArchivePathFile_in
        .strTempPathWBS = strTempPathWBS_in
        .strRunnableFile = ""
        .a_strRunnableFileTypes = a_strRunnableFileTypes_in
        .blnOKPressed = False
    End With

    frmArchive.Show vbModal

    With g_udtArchiveBrowse
        If .blnOKPressed Then
            strRunnable_out = .strRunnableFile
            GetRunnableFileInArchiveFromUser = True
        Else
            GetRunnableFileInArchiveFromUser = False
        End If
    End With

End Function

Public Function GetUnpackableFileTypeString() As String

    GetUnpackableFileTypeString = m_strUnpackableFileTypes

End Function

Public Function GetRepackableFileTypeString() As String

    GetRepackableFileTypeString = m_strRepackableFileTypes

End Function
Public Function Repack(ByVal blnAskBefore_in As Boolean, ByVal blnNotifyAfter_in As Boolean, strGBGamePathWBS_in As String, _
                    strGamePathFile_in As String, eWithAdditionalFiles As EnumRezipAdditionalFiles, _
                    blnWasRepacked_out As Boolean, strError_out As String) As Boolean

    On Error GoTo Err_Handler

    Dim lngFolderIdx As Long

    Dim strDestArchivePathFile As String

    Dim strFolderToRepack As String
    Dim strFolderToRepackWBS As String
    Dim strParentFolderWBS As String
    Dim strExt As String
    Dim strGamePath As String
    Dim strGameFile As String
    Dim str7ZipListFile As String
    Dim str7zExe As String

    Repack = False
    strError_out = ""
    blnWasRepacked_out = False

    Call PushLogFileTrace("clsUnpack.Repack")

    'need 7-Zip
    str7zExe = GetExtractorPath("7z.exe")
    If str7zExe = "" Then
        strError_out = GetINIMessage("Running Games", "SevenZipNotFound", "Unable to locate 7z.exe for repacking archives")
        Call PopLogFileTrace
        Exit Function
    End If
    Call LogStatus("7-Zip Path: [" & str7zExe & "]")

    'final folder must exist
    If m_lngFinalFolderIdx < 0 Then
        strError_out = GetINIMessage("Running Games", "FinalFolderNotKnown", "Unable to repack. Final unpacked archive folder not known. Perhaps nothing was unpacked?")
        Call PopLogFileTrace
        Exit Function
    End If
    lngFolderIdx = m_lngFinalFolderIdx

    'get the starting folder to repack
    strFolderToRepack = strGBGamePathWBS_in
    strFolderToRepack = strFolderToRepack & m_lngFinalFolderIdx
    strFolderToRepackWBS = strFolderToRepack
    If Right$(strFolderToRepackWBS, 1) <> "\" Then strFolderToRepackWBS = strFolderToRepackWBS & "\"
    Call LogStatus("Folder to repack: [" & strFolderToRepack & "]")

    'get the main archive path and file
    Call m_xobjFunc.GetFileAndPath(strGamePathFile_in, strGamePath, strGameFile)

    'runnable file must exist
    If (m_strRunnableRelativePathFile = "") Or (Not m_xobjFunc.FileExists(strFolderToRepackWBS & m_strRunnableRelativePathFile, False, "GEMUS")) Then
        strError_out = GetINIMessage("Running Games", "RunnableFileNotFound", "Unable to repack. The runnable file was not found in the unpacked folder.")
        Call PopLogFileTrace
        Exit Function
    End If

    'must have pre-play snapshot
    If m_lngTotalPreUnpackedFilesSnapshot = 0 Then
        strError_out = GetINIMessage("Running Games", "NoPreSnapshot", "Unable to repack. Snapshot of unpacked folder before playing does not exist.")
        Call PopLogFileTrace
        Exit Function
    End If

    'get post-play snapshot
    Call StorePostUnpackedFilesSnapshot(strFolderToRepack)
    If m_lngTotalPostUnpackedFilesSnapshot = 0 Then
        strError_out = GetINIMessage("Running Games", "NoFilesInUnpackedFolder", "Unable to repack. No files exist in the unpacked folder after playing.")
        Call PopLogFileTrace
        Exit Function
    End If

    'get the modified post snapshot for repacking
    'NOTE: this asks the user which additional files to repack, if necessary
    Call GenerateModifiedPostSnapshot(strFolderToRepackWBS, strGameFile, eWithAdditionalFiles)
    Call LogStatus("Stored current files snapshot from [" & strFolderToRepackWBS & "]")

    'only need to repack if the snapshots have changed
    If Not DidSnapshotsChangeDuringPlay() Then
        Repack = True 'not an error, just not repacked (not needed)
        Call LogStatus("No files changed. Repack aborted.")
        Call PopLogFileTrace
        Exit Function
    End If

    If blnAskBefore_in = True Then
        If MsgBox(GetINIMessage("Running Games", "RepackConfirm", "There have been changes made to the game file(s).  Do you want to repack them") & "?", vbQuestion Or vbYesNo, "GEMUS") = vbNo Then
            Repack = True 'not an error, just not repacked (not wanted)
            Call LogStatus("User decided not to repack when asked. Repack aborted.")
            Call PopLogFileTrace
            Exit Function
        End If
    End If

    'generate the 7-zip list file to use in the initial repack from the post-snapshot
    If Not GenerateRepackListFileFromPostSnapshot(strFolderToRepackWBS, str7ZipListFile, strError_out) Then
        strError_out = GetINIMessage("Running Games", "RepackListFileError", "Unable to repack. Could not create repack list file for 7-Zip.") & " " & strError_out
        Repack = False
        Call PopLogFileTrace
        Exit Function
    End If
    Call LogStatus("Generated 7-Zip list file containing list of files/folders to repack [" & str7ZipListFile & "]")

    'repack the final folder (from the generated list file) into the first parent archive
    If lngFolderIdx > 0 Then
        'sub archive
        strParentFolderWBS = strGBGamePathWBS_in & Trim$(lngFolderIdx - 1) & "\"
        If Not GetRepackParentArchive(strParentFolderWBS, lngFolderIdx - 1, strDestArchivePathFile, strError_out) Then
            Call PopLogFileTrace
            Exit Function
        End If
    Else
        'main archive
        strExt = Right$(strGamePathFile_in, Len(strGamePathFile_in) - InStrRev(strGamePathFile_in, "."))
        strDestArchivePathFile = strGBGamePathWBS_in & "repacked." & strExt
    End If
    Call LogStatus("Repacking [" & strFolderToRepack & "] to [" & strDestArchivePathFile & "] using list file [" & str7ZipListFile & "]")
    If Not SevenZipRepackWithListFile(str7zExe, strFolderToRepack, str7ZipListFile, strDestArchivePathFile, strError_out) Then
        Call PopLogFileTrace
        Exit Function
    End If

    'repack each remaining folder up to the main archive
    lngFolderIdx = lngFolderIdx - 1
    Do While lngFolderIdx > -1

        'get the parent archive to repack into
        If lngFolderIdx > 0 Then
            'sub archive
            strParentFolderWBS = strGBGamePathWBS_in & Trim$(lngFolderIdx - 1) & "\"
            If Not GetRepackParentArchive(strParentFolderWBS, lngFolderIdx - 1, strDestArchivePathFile, strError_out) Then
                Call PopLogFileTrace
                Exit Function
            End If
        Else
            'main archive
            strExt = Right$(strGamePathFile_in, Len(strGamePathFile_in) - InStrRev(strGamePathFile_in, "."))
            strDestArchivePathFile = strGBGamePathWBS_in & "repacked." & strExt
        End If

        'before repacking the current folder, remove the __gemus file, and make
        'sure its deleted before continuing as we don't want to repack it
        Call m_xobjFunc.DeletePathSilent(strGBGamePathWBS_in & lngFolderIdx & "\__gemus")
        If m_xobjFunc.FileExists(strGBGamePathWBS_in & lngFolderIdx & "\__gemus", False, "GEMUS") Then
            strError_out = GetINIMessage("Running Games", "RemoveTempGemusFileError", "Unable to repack. Unable to remove temporary __gemus file") _
                & vbCrLf & vbCrLf & _
                strGBGamePathWBS_in & lngFolderIdx & "\__gemus"
            Call PopLogFileTrace
            Exit Function
        End If

        'repack the current folder to the parent folder archive
        Call LogStatus("Repacking [" & strGBGamePathWBS_in & Trim$(lngFolderIdx) & "] to [" & strDestArchivePathFile & "]")
        If Not SevenZipRepackFolder(str7zExe, strGBGamePathWBS_in, lngFolderIdx, strDestArchivePathFile, strError_out) Then
            Call PopLogFileTrace
            Exit Function
        End If

        'next level up
        lngFolderIdx = lngFolderIdx - 1
    Loop

    'now the file "repacked.[archive-file-extension]" should exist in
    'the temp GBGame folder so place the original game file with it

    'rename existing gamefile to .old
    Call LogStatus("Renaming [" & strGamePathFile_in & "] to [" & strGamePathFile_in & ".old]")
    Call m_xobjFunc.DeletePathSilent(strGamePathFile_in & ".old")
    Name (strGamePathFile_in) As (strGamePathFile_in & ".old")
    'rename the temp repacked archived as the game file
    Call LogStatus("Renaming [" & strDestArchivePathFile & "] to [" & strGamePathFile_in & "]")
    Name (strDestArchivePathFile) As (strGamePathFile_in)
    'delete the .old file
    Call LogStatus("Deleting [" & strGamePathFile_in & ".old]")
    Call m_xobjFunc.DeletePathSilent(strGamePathFile_in & ".old")

    'success
    If blnNotifyAfter_in = True Then
        Call MsgBox(GetINIMessage("Running Games", "RepackSuccess", "The game was repacked successfully!"), vbOKOnly Or vbInformation, "GEMUS")
    End If
    blnWasRepacked_out = True
    strError_out = ""
    Repack = True
    Call PopLogFileTrace
    Exit Function

Err_Handler:
    strError_out = GetINIMessage("Running Games", "UnableToRepack", "Unable to repack.") & " " & Err.Description & " (" & Err.Number & "," & Trim$(Err.LastDllError) & ")"
    Repack = False
    Call PopLogFileTrace

End Function
Private Sub GenerateModifiedPostSnapshot(strUnpackedFolderWBS_in As String, strArchiveName_in As String, _
                                            eWithAdditionalFiles As EnumRezipAdditionalFiles)

    'assumes the pre and post snapshots exist

    Dim blnFound As Boolean

    Dim lngPreIdx As Long
    Dim lngPostIdx As Long
    Dim lngAdditionalIdx As Long
    Dim lngTotalAdditional As Long
    Dim lngTotalExcludingAdditional As Long

    Dim udtItem As UnpackedFileSnapshot
    Dim a_strAdditional() As String
    Dim a_udtAdditional() As UnpackedFileSnapshot
    Dim a_udtExcludingAdditional() As UnpackedFileSnapshot

    'just use the post snapshot if we're to repack any additional files
    If eWithAdditionalFiles = eRezipAdditionalZip Then Exit Sub

    'look for any items in the post-snapshot that aren't in the pre-snapshot
    lngTotalAdditional = 0
    lngTotalExcludingAdditional = 0
    ReDim a_strAdditional(0)
    ReDim a_udtAdditional(0)
    ReDim a_udtExcludingAdditional(0)
    For lngPostIdx = 0 To m_lngTotalPostUnpackedFilesSnapshot - 1
        blnFound = False
        udtItem = ma_udtPostUnpackedFilesSnapshot(lngPostIdx)
        For lngPreIdx = 0 To m_lngTotalPreUnpackedFilesSnapshot - 1
            If ma_udtPreUnpackedFilesSnapshot(lngPreIdx).strUCaseRelativePathFile = ma_udtPostUnpackedFilesSnapshot(lngPostIdx).strUCaseRelativePathFile Then
                blnFound = True
                Exit For
            End If
        Next lngPreIdx
        If Not blnFound Then
            'additional file
            ReDim Preserve a_strAdditional(lngTotalAdditional)
            a_strAdditional(lngTotalAdditional) = udtItem.strRelativePathFile
            ReDim Preserve a_udtAdditional(lngTotalAdditional)
            a_udtAdditional(lngTotalAdditional) = udtItem
            lngTotalAdditional = lngTotalAdditional + 1
        Else
            'it was found in the pre-snapshot
            ReDim Preserve a_udtExcludingAdditional(lngTotalExcludingAdditional)
            a_udtExcludingAdditional(lngTotalExcludingAdditional) = udtItem
            lngTotalExcludingAdditional = lngTotalExcludingAdditional + 1
        End If
    Next lngPostIdx

    'if no additional files, just use the unmodified post snapshot
    If lngTotalAdditional = 0 Then Exit Sub

    'post snapshot is now by default without any new/additional files
    ma_udtPostUnpackedFilesSnapshot() = a_udtExcludingAdditional()
    m_lngTotalPostUnpackedFilesSnapshot = lngTotalExcludingAdditional

    'use the array that excludes the additional files if the user wants to ignore additional files
    If eWithAdditionalFiles = eRezipAdditionalIgnore Then Exit Sub

    'ask the user which of the additional files he/she wants to repack
    Call PushLogFileTrace("clsUnpack.GenerateModifiedPostSnapshot")
    Call LogStatus("Additional files found. Asking user which ones to include in repack.")
    Load frmRepackAdditional
    Call frmRepackAdditional.SetArchiveName(strArchiveName_in)
    Call frmRepackAdditional.SetAdditionalFiles(a_strAdditional)
    frmRepackAdditional.Show vbModal
    Call frmRepackAdditional.GetAdditionalFiles(a_strAdditional, lngTotalAdditional)
    Unload frmRepackAdditional
    Call LogStatus(lngTotalAdditional & " files selected for inclusion by user.")
    Call PopLogFileTrace

    'add any selected additional files to the post snapshot
    For lngAdditionalIdx = 0 To lngTotalAdditional - 1
        ReDim Preserve ma_udtPostUnpackedFilesSnapshot(m_lngTotalPostUnpackedFilesSnapshot)
        With ma_udtPostUnpackedFilesSnapshot(m_lngTotalPostUnpackedFilesSnapshot)
            .strCRC = m_objCRC.CalculateFile(strUnpackedFolderWBS_in & a_strAdditional(lngAdditionalIdx))
            .strRelativePathFile = a_strAdditional(lngAdditionalIdx)
            .strUCaseRelativePathFile = UCase$(.strRelativePathFile)
        End With
        m_lngTotalPostUnpackedFilesSnapshot = m_lngTotalPostUnpackedFilesSnapshot + 1
    Next lngAdditionalIdx
    
End Sub

Private Function SevenZipRepackFolder(strExtractorPath_in As String, strGBGamePathWBS_in As String, ByVal lngFolderIdx_in As Long, _
                                        strDestArchivePathFile_in As String, strError_out As String) As Boolean
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strDefaultError As String
    Dim lngNumFiles As Long
    Dim a_strFiles() As String
    Dim strRepackFolder As String
    Dim strRepackFolderWBS As String
    Dim lngRecursiveIdx As Long
    Dim lngAddedIdx As Long
    Dim strItem As String
    Dim blnFound As Boolean

    Call PushLogFileTrace("clsUnpack.SevenZipRepackFolder")

    strError_out = ""
    SevenZipRepackFolder = False

    'delete the existing destination archive
    Call m_xobjFunc.DeletePathSilent(strDestArchivePathFile_in)

    'get the repack folder
    strRepackFolder = strGBGamePathWBS_in & lngFolderIdx_in
    strRepackFolderWBS = strRepackFolder & "\"

    'change to drive/directory that wants repacking
    ChDrive strRepackFolder
    ChDir strRepackFolder
    Call LogStatus("Changed directory to [" & strRepackFolder & "]")

    'Run 7z.exe.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    '  a    add to archive
    ' -r    recursive
    ' -y    answer yes on all queries
    ' (7-zip help file recommends using long filenames, as short [8.3] ones may not work 100% of the time)
    strCommandLine = m_strCommandCom & strExtractorPath_in & " a """ & strDestArchivePathFile_in & """ * -ry"
    Call LogStatus("Running 7-Zip [" & strCommandLine & "]")
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

    strDefaultError = " [" & strDestArchivePathFile_in & "] " & _
                        GetINIMessage("Running Games", "SevenZipRepackCommandFrom", "7-Zip repack command from") & _
                        " [" & strRepackFolder & "]:" & vbCrLf & vbCrLf & strCommandLine

    'if the dest folder is there, it might have worked
    If Not m_xobjFunc.FileExists(strDestArchivePathFile_in, False, "GEMUS") Then
        strError_out = GetINIMessage("", "UnableToCreateArchive", "Unable to create archive") & strDefaultError
        Call PopLogFileTrace
        Exit Function
    End If

    'get a list file so the contents of the created archive can be compared to those that should have been added
    strError_out = SevenZipList(strDestArchivePathFile_in, strRepackFolderWBS & "__gemus_repack_check.txt", lngNumFiles, a_strFiles)
    If strError_out <> "" Then
        strError_out = GetINIMessage("Running Games", "UnableToGetListOfArchiveContents", "Unable to get list of contents for repacked archive") & strDefaultError & vbCrLf & vbCrLf & strError_out
        Call PopLogFileTrace
        Exit Function
    End If

    'make sure all added files are in the archive
    Call StoreRecursiveUnpackedFilesSnapshot(strRepackFolder)
    For lngRecursiveIdx = 0 To m_lngTotalRecursiveUnpackedFilesSnapshot - 1
        strItem = UCase$(ma_strRecursiveUnpackedFilesSnapshot(lngRecursiveIdx))
        blnFound = False
        For lngAddedIdx = 0 To lngNumFiles - 1
            If UCase$(a_strFiles(lngAddedIdx)) = strItem Then
                blnFound = True
                Exit For
            End If
        Next lngAddedIdx
        If Not blnFound Then
            strError_out = GetINIMessage("Running Games", "RepackAbortedMissingFiles", "Repack aborted. Did not find all required files in the destination archive") & " " & strDefaultError
            Call PopLogFileTrace
            Exit Function
        End If
    Next lngRecursiveIdx

    'successfully repacked
    SevenZipRepackFolder = True
    Call PopLogFileTrace

End Function

Public Sub StorePreUnpackedFilesSnapshot(strFolder_in As String)

    Dim lngIdx As Long
    Dim lngLen As Long

    ReDim ma_udtPreUnpackedFilesSnapshot(0)
    m_lngTotalPreUnpackedFilesSnapshot = 0

    If strFolder_in = "" Then Exit Sub

    Set m_xobjFuncPre = New GBFuncs.clsFunctions
    Call m_xobjFuncPre.FindFiles(strFolder_in, "*.*") 'calls callback function
    Set m_xobjFuncPre = Nothing

    lngLen = Len(strFolder_in)
    For lngIdx = 0 To m_lngTotalPreUnpackedFilesSnapshot - 1
        With ma_udtPreUnpackedFilesSnapshot(lngIdx)
            .strRelativePathFile = Right$(.strRelativePathFile, Len(.strRelativePathFile) - lngLen)
            .strUCaseRelativePathFile = UCase$(.strRelativePathFile)
        End With
    Next lngIdx

End Sub

Public Sub StorePostUnpackedFilesSnapshot(strFolder_in As String)

    Dim lngIdx As Long
    Dim lngLen As Long

    ReDim ma_udtPostUnpackedFilesSnapshot(0)
    m_lngTotalPostUnpackedFilesSnapshot = 0
    
    If strFolder_in = "" Then Exit Sub
    
    Set m_xobjFuncPost = New GBFuncs.clsFunctions
    Call m_xobjFuncPost.FindFiles(strFolder_in, "*.*") 'calls callback function
    Set m_xobjFuncPost = Nothing

    lngLen = Len(strFolder_in)
    For lngIdx = 0 To m_lngTotalPostUnpackedFilesSnapshot - 1
        With ma_udtPostUnpackedFilesSnapshot(lngIdx)
            .strRelativePathFile = Right$(.strRelativePathFile, Len(.strRelativePathFile) - lngLen)
            .strUCaseRelativePathFile = UCase$(.strRelativePathFile)
        End With
    Next lngIdx

End Sub
Public Sub StoreRecursiveUnpackedFilesSnapshot(strFolder_in As String)

    Dim lngIdx As Long
    Dim lngLen As Long
    Dim strPathFile As String

    ReDim ma_strRecursiveUnpackedFilesSnapshot(0)
    m_lngTotalRecursiveUnpackedFilesSnapshot = 0
    
    If strFolder_in = "" Then Exit Sub
    
    Set m_xobjFuncRecursive = New GBFuncs.clsFunctions
    Call m_xobjFuncRecursive.FindFiles(strFolder_in, "*.*") 'calls callback function
    Set m_xobjFuncRecursive = Nothing

    lngLen = Len(strFolder_in)
    For lngIdx = 0 To m_lngTotalRecursiveUnpackedFilesSnapshot - 1
        strPathFile = ma_strRecursiveUnpackedFilesSnapshot(lngIdx)
        ma_strRecursiveUnpackedFilesSnapshot(lngIdx) = Right$(strPathFile, Len(strPathFile) - lngLen)
    Next lngIdx

End Sub
Public Function IsFileASupportedArchiveFileType(strFile As String) As Boolean

    IsFileASupportedArchiveFileType = IsFileMatchingFileTypes(strFile, ma_strUnpackableFileTypes)

End Function
Public Function IsFileARepackableArchiveFileType(strFile As String) As Boolean

    IsFileARepackableArchiveFileType = IsFileMatchingFileTypes(strFile, ma_strRepackableFileTypes)

End Function


Public Function SevenZipUnpack(strExtractorPath_in As String, strPath_in As String, strFile_in As String, _
                                strDestPath_in As String) As Boolean

    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String

    Call PushLogFileTrace("clsUnpack.SevenListUnpack")

    'change to drive/directory containing archive
    ChDrive strPath_in
    ChDir strPath_in
    
    Call LogStatus("Changed Directory to [" & strPath_in & "]")

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
    Call LogStatus("Running 7-Zip [" & strCommandLine & "]")
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

    'if the dest folder is there, assume it worked
    SevenZipUnpack = m_xobjFunc.DirExists(strDestPath_in)
    
    Call PopLogFileTrace

End Function

Public Function SevenZipRepackWithListFile(strExtractorPath_in As String, strRepackFolder_in As String, strListFile_in As String, _
                                strDestArchivePathFile_in As String, strError_out As String) As Boolean

    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strDefaultError As String
    Dim lngNumFiles As Long
    Dim a_strFiles() As String
    Dim strRepackFolderWBS As String
    Dim lngPostIdx As Long
    Dim lngAddedIdx As Long
    Dim strItem As String
    Dim blnFound As Boolean

    Call PushLogFileTrace("clsUnpack.SevenZipRepackWithListFile")

    strError_out = ""
    SevenZipRepackWithListFile = False
    strRepackFolderWBS = strRepackFolder_in
    If Right$(strRepackFolderWBS, 1) <> "\" Then strRepackFolderWBS = strRepackFolderWBS & "\"

    'delete the existing destination archive
    Call m_xobjFunc.DeletePathSilent(strDestArchivePathFile_in)

    'change to drive/directory containing archive
    ChDrive strRepackFolder_in
    ChDir strRepackFolder_in
    Call LogStatus("Changed directory to [" & strRepackFolder_in & "]")

    'Run 7z.exe.
    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With
    '  a    add to archive
    ' -y    answer yes on all queries
    ' (7-zip help file recommends using long filenames, as short [8.3] ones may not work 100% of the time)
    strCommandLine = m_strCommandCom & strExtractorPath_in & " a """ & strDestArchivePathFile_in & """ -y @" & strListFile_in
    Call LogStatus("Running 7-Zip [" & strCommandLine & "]")
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

    strDefaultError = " [" & strDestArchivePathFile_in & "] " & _
                        GetINIMessage("Running Games", "SevenZipRepackCommandFrom", "7-Zip repack command from") & _
                        " [" & strRepackFolder_in & "]:" & vbCrLf & vbCrLf & strCommandLine

    'if the dest file is there, it might have worked
    If Not m_xobjFunc.FileExists(strDestArchivePathFile_in, False, "GEMUS") Then
        strError_out = GetINIMessage("Running Games", "UnableToCreateArchive", "Unable to create archive") & strDefaultError
        Call PopLogFileTrace
        Exit Function
    End If

    'get a list file so the contents of the created archive can be compared to those that should have been added
    strError_out = SevenZipList(strDestArchivePathFile_in, strRepackFolderWBS & "__gemus_repack_check.txt", lngNumFiles, a_strFiles)
    If strError_out <> "" Then
        strError_out = GetINIMessage("Running Games", "UnableToGetListOfArchiveContents", "Unable to get list of contents for repacked archive") & strDefaultError & vbCrLf & vbCrLf & strError_out
        Call PopLogFileTrace
        Exit Function
    End If

    'make sure all added files are in the archive
    For lngPostIdx = 0 To m_lngTotalPostUnpackedFilesSnapshot - 1
        strItem = ma_udtPostUnpackedFilesSnapshot(lngPostIdx).strUCaseRelativePathFile
        blnFound = False
        For lngAddedIdx = 0 To lngNumFiles - 1
            If UCase$(a_strFiles(lngAddedIdx)) = strItem Then
                blnFound = True
                Exit For
            End If
        Next lngAddedIdx
        If Not blnFound Then
            strError_out = GetINIMessage("Running Games", "RepackAbortedMissingFiles", "Repack aborted. Did not find all required files in the destination archive") & " " & strDefaultError
            Call PopLogFileTrace
            Exit Function
        End If
    Next lngPostIdx

    'successfully repacked
    SevenZipRepackWithListFile = True
    Call PopLogFileTrace

End Function
Public Function SevenZipList(strArchivePathFile As String, strTempPathFile As String, _
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

    Call PushLogFileTrace("clsUnpack.SevenZipList")

    'get path to 7z.exe
    str7zExe = GetExtractorPath("7z.exe")
    If str7zExe = "" Then
        SevenZipList = GetINIMessage("Extracting", "Error_No7z", "Path to 7z.exe not found, unable to check contents of archive") & " " & strArchivePathFile
        Call PopLogFileTrace
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
    Call LogStatus("Changed Directory to [" & strArchivePath & "]")

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
    Call LogStatus("Running 7-Zip [" & strCommandLine & "]")
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

    'check for existence of the file and read in its contents
    If m_xobjFunc.FileExists(strTempPathFile, False, "") = False Then
        SevenZipList = GetINIMessage("Extracting", "Error_TempFile", "Temporary list file could not be created. Unable to list contents of archive") & " " & strArchivePathFile
        Call PopLogFileTrace
        Exit Function
    End If
    f = FreeFile()
    Open strTempPathFile For Input As #f
    Call LogStatus("Opened file for reading [" & strTempPathFile & "]")
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
    Call LogStatus("Closed file [" & strTempPathFile & "] after reading " & lngTotalFiles_out & " valid file lines")

    'no errors
    SevenZipList = ""
    Call m_xobjFunc.DeletePathSilent(strTempPathFile)
    Call PopLogFileTrace
    Exit Function

Err_Handler:
    SevenZipList = GetINIMessage("Extracting", "Error_Unexpected", "Unexpected Error. Unable to list contents of archive") & " " & strArchivePathFile _
        & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")"
    Call PopLogFileTrace

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
Public Function GetExtractorPath(strFilename_in As String) As String

    If m_xobjFunc.FileExists(m_strAppPathWBS & strFilename_in, False, "") = True Then
        GetExtractorPath = m_xobjFunc.Long2Short(m_strAppPathWBS & strFilename_in)
    ElseIf m_xobjFunc.FileExists(strFilename_in, False, "") = True Then
        GetExtractorPath = strFilename_in
    Else
        GetExtractorPath = ""
    End If

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
Public Function Unpack(strPathFileToUnpack_in As String, strRunnableFile_inout As String, _
                        a_strFileTypes_in() As String, ByVal lngTotalFileTypes_in As Long, _
                        strDestFolderWBS_in As String, blnIsRepackable_out As Boolean, _
                        strPathFile_out As String, blnNeedsUserToPickRunnableFile_out As Boolean, _
                        strSelectError_out As String, strRunError_out As String) As Boolean

    Dim blnHasRunnable As Boolean
    Dim lngNumRunnable As Long
    Dim lngNumFiles As Long
    Dim lngNumRunnableFiles As Long
    Dim lngNumUnpackableFiles As Long
    Dim str7zExe As String
    Dim strArchivePathFile As String
    Dim strArchivePath As String
    Dim strArchiveFile As String
    Dim strRunnable As String
    Dim strRunnablePathFile As String
    Dim a_strRunnable() As String
    Dim a_strFiles() As String
    Dim a_strRunnableFiles() As String
    Dim a_strUnpackableFiles() As String
    Dim a_strGeneratedRunnable() As String

    lngNumRunnable = 0
    m_lngFinalFolderIdx = -1
    m_strFinalFolder = ""
    m_strRunnableRelativePathFile = ""
    blnIsRepackable_out = False

    Call PushLogFileTrace("clsUnpack.Unpack")

    'must have filetypes
    If lngTotalFileTypes_in < 1 Then
        strRunError_out = GetINIMessage("Running Games", "NoConfiguredFileTypes", "No configured filetypes")
        strSelectError_out = strRunError_out
        Unpack = False
        Call PopLogFileTrace
        Exit Function
    End If

    'must be either a compatible file type, or a supported archive file type
    If IsFileMatchingFileTypes(strPathFileToUnpack_in, a_strFileTypes_in) = False _
     And IsFileMatchingFileTypes(strPathFileToUnpack_in, ma_strUnpackableFileTypes) = False Then
        strRunError_out = GetINIMessage("Running Games", "SpecifiedFileTypeNotSupported", "Specified file type is not supported")
        strSelectError_out = strRunError_out
        Unpack = False
        Call PopLogFileTrace
        Exit Function
    End If

    'split the runnable file string
    blnIsRepackable_out = False
    blnHasRunnable = False
    If Trim$(strRunnableFile_inout) <> "" Then
        'make sure the last runnable file is a valid filetype
        If Not IsFileMatchingFileTypes(strRunnableFile_inout, a_strFileTypes_in) Then
            'specified runnable file does not match the allowed file types
            strRunError_out = GetINIMessage("Running Games", "RunnableFileTypeMismatch", "Specified runnable file does not match the configured filetypes")
            strSelectError_out = strRunError_out
            Unpack = False
            Call PopLogFileTrace
            Exit Function
        End If
        a_strRunnable = Split(Trim$(strRunnableFile_inout), ":")
        lngNumRunnable = UBound(a_strRunnable) + 1
        blnHasRunnable = True
        Call LogStatus("Split runnable file into " & lngNumRunnable & " part(s): [ " & Join(a_strRunnable, " : ") & " ]")
    End If

    'prepare initial variables
    blnNeedsUserToPickRunnableFile_out = False
    strArchivePathFile = strPathFileToUnpack_in

    'no need to unpack the archive if it's a valid file type
    'and there is no runnable file set within the archive
    If Not blnHasRunnable Then
        If IsFileMatchingFileTypes(strArchivePathFile, a_strFileTypes_in) Then
            'archive file is the "game" file to send to GEMUS
            Call LogStatus("Skipped unpacking [" & strArchivePathFile & "]. File matches configured filetypes.")
            strPathFile_out = strArchivePathFile
            Unpack = True
            Call PopLogFileTrace
            Exit Function
        End If
    End If

    'only continue to the unpacking loop if the file is one 7-zip can handle
    If Not IsFileMatchingFileTypes(strArchivePathFile, ma_strUnpackableFileTypes) Then
            strRunError_out = GetINIMessage("Running Games", "SpecifiedArchiveTypeNotSupported", "Specified archive type is not supported")
            strSelectError_out = strRunError_out
            Unpack = False
            Call PopLogFileTrace
            Exit Function
    End If

    'is repack possible?
    blnIsRepackable_out = True
    If lngNumRunnable > 1 Then
        For m_lngFinalFolderIdx = 0 To lngNumRunnable - 2 'ignore last one as it's the actual runnable file
            strRunnable = a_strRunnable(m_lngFinalFolderIdx)
            If Not IsFileMatchingFileTypes(strRunnable, ma_strRepackableFileTypes) Then
                blnIsRepackable_out = False
                Call LogStatus("Repacking not possible due to the filetype of this file: [" & strRunnable & "]")
                Exit For
            End If
        Next m_lngFinalFolderIdx
    End If
    If blnIsRepackable_out = True Then
        'main archive file must be repackable also
        If Not IsFileMatchingFileTypes(strArchivePathFile, ma_strRepackableFileTypes) Then
            blnIsRepackable_out = False
            Call LogStatus("Repacking not possible due to the filetype of this file: [" & strArchivePathFile & "]")
        Else
            Call LogStatus("Repacking is possible")
        End If
    End If

    'need 7-Zip
    str7zExe = GetExtractorPath("7z.exe")
    If str7zExe = "" Then
        strRunError_out = GetINIMessage("Running Games", "SevenZipNotFound", "Unable to locate 7z.exe for unpacking archives")
        strSelectError_out = strRunError_out
        Unpack = False
        Call PopLogFileTrace
        Exit Function
    End If
    Call LogStatus("7-Zip Path: [" & str7zExe & "]")

    'loop until the runnable file is found or an error occurs
    m_lngFinalFolderIdx = 0
    ReDim a_strGeneratedRunnable(0)
    Do
        'extract the archive
        Call m_xobjFunc.GetFileAndPath(strArchivePathFile, strArchivePath, strArchiveFile)
        m_strFinalFolder = strDestFolderWBS_in & m_lngFinalFolderIdx
        Call m_xobjFunc.DeletePathSilent(m_strFinalFolder)
        Call LogStatus("Unpacking [" & strArchivePathFile & "] to folder [" & m_strFinalFolder & "]")
        If Not SevenZipUnpack(str7zExe, strArchivePath, strArchiveFile, m_strFinalFolder) Then
            strRunError_out = GetINIMessage("Running Games", "SevenZipFailedUnpackArchive", "7-Zip failed to unpack archive") & " " & strArchivePathFile & " " & GetINIMessage("Running Games", "SevenZipFailedTo", "to") & " " & m_strFinalFolder
            strSelectError_out = GetINIMessage("Running Games", "SevenZipFailedUnpackArchive", "7-Zip failed to unpack archive")
            Unpack = False
            Call PopLogFileTrace
            Exit Function
        End If

        If blnHasRunnable Then
        
            '*************************** THERE IS A RUNNABLE FILE SPECIFIED ***************************
        
            strRunnable = a_strRunnable(m_lngFinalFolderIdx)
            strRunnablePathFile = m_strFinalFolder & "\" & strRunnable
            Call LogStatus("Trying to find runnable file or archive [" & strRunnablePathFile & "]")
            If Not m_xobjFunc.FileExists(strRunnablePathFile, False, "") Then
                strRunError_out = GetINIMessage("Running Games", "ExpectedFile", "Expected file") & " [" & strRunnable & "] " & GetINIMessage("Running Games", "NotFoundInArchive", "not found in archive") & " " & strArchivePathFile
                strSelectError_out = GetINIMessage("Running Games", "ExpectedFileNotFoundInArchive", "Expected file not found in archive")
                Unpack = False
                Call PopLogFileTrace
                Exit Function
            End If
            If UBound(a_strRunnable) = m_lngFinalFolderIdx Then
                'final runnable file
                strPathFile_out = strRunnablePathFile
                m_strRunnableRelativePathFile = strRunnable
                Unpack = True
                Call LogStatus("Found runnable file [" & strRunnablePathFile & "]")
                Call PopLogFileTrace
                Exit Function
            Else
                'more to come, this must be another archive to unpack
                Call LogStatus("Found runnable file (archive) [" & strRunnablePathFile & "]")
                Call LogStatus("Creating archive log file [" & m_strFinalFolder & "\__gemus" & "] with contents [" & strRunnable & "]")
                If Not m_xobjFunc.CreateFileWithContents(m_strFinalFolder & "\__gemus", strRunnable, strRunError_out) Then
                    strRunError_out = GetINIMessage("Running Games", "UnableToWriteArchiveLogFile", "Unable to write archive log file") & " " & m_strFinalFolder & "\__gemus" & vbCrLf & vbCrLf & strRunError_out
                    strSelectError_out = GetINIMessage("Running Games", "UnableToWriteArchiveLogFile", "Unable to write archive log file")
                    Unpack = False
                    Call PopLogFileTrace
                    Exit Function
                End If
                'next unpacking iteration...
                m_lngFinalFolderIdx = m_lngFinalFolderIdx + 1
                strArchivePathFile = strRunnablePathFile
            End If

        Else

            '*************************** TRY TO FIND *ANY* RUNNABLE FILE ***************************

            'get a list of this archive's contents
            Call LogStatus("Listing contents of archive [" & strArchivePathFile & "] into temp file [" & strDestFolderWBS_in & "list.txt" & "]")
            strRunError_out = SevenZipList(strArchivePathFile, strDestFolderWBS_in & "list.txt", lngNumFiles, a_strFiles)
            If strRunError_out <> "" Then
                strSelectError_out = "Unable to get list of archive contents"
                Unpack = False
                Call PopLogFileTrace
                Exit Function
            End If
            If lngNumFiles = 0 Then
                strRunError_out = GetINIMessage("Running Games", "NoRunnableMatchingFilesInArchive", "No runnable files found in archive matching configured file types.")
                strSelectError_out = strRunError_out
                Unpack = False
                Call PopLogFileTrace
                Exit Function
            End If
            'count the number of runnable files and archive files
            lngNumRunnableFiles = GetNumFilesMatchingFileTypes(a_strFiles, a_strFileTypes_in, a_strRunnableFiles)
            lngNumUnpackableFiles = GetNumFilesMatchingFileTypes(a_strFiles, ma_strUnpackableFileTypes, a_strUnpackableFiles)
            Call LogStatus("Found " & lngNumRunnableFiles & " runnable file(s), and " & lngNumUnpackableFiles & " unpackable file(s)")
            If lngNumRunnableFiles = 1 And lngNumUnpackableFiles = 0 Then
                'found single runnable file
                strPathFile_out = m_strFinalFolder & "\" & a_strRunnableFiles(0)
                ReDim Preserve a_strGeneratedRunnable(m_lngFinalFolderIdx)
                a_strGeneratedRunnable(m_lngFinalFolderIdx) = a_strRunnableFiles(0)
                strRunnableFile_inout = Join(a_strGeneratedRunnable, ":")
                m_strRunnableRelativePathFile = a_strRunnableFiles(0)
                Unpack = True
                Call LogStatus("Found runnable file [" & strRunnableFile_inout & "]")
                Call PopLogFileTrace
                Exit Function
            ElseIf lngNumRunnableFiles = 0 And lngNumUnpackableFiles = 1 Then
                'found single unpackable archive
                strArchivePathFile = m_strFinalFolder & "\" & a_strUnpackableFiles(0)
                ReDim Preserve a_strGeneratedRunnable(m_lngFinalFolderIdx)
                a_strGeneratedRunnable(m_lngFinalFolderIdx) = a_strUnpackableFiles(0)
                m_lngFinalFolderIdx = m_lngFinalFolderIdx + 1
                Call LogStatus("Found single unpackable archive [" & strArchivePathFile & "]")
            ElseIf (lngNumRunnableFiles > 1 Or lngNumUnpackableFiles > 1) _
            Or (lngNumRunnableFiles = 1 And lngNumUnpackableFiles = 1) Then
                'user needs to select which file to run
                blnNeedsUserToPickRunnableFile_out = True
                Unpack = True
                strPathFile_out = ""
                Call LogStatus("Found more than one runnable file/archive. User needs to pick one.")
                Call PopLogFileTrace
                Exit Function
            Else
                'no compatible files
                strRunError_out = GetINIMessage("Running Games", "NoRunnableMatchingFilesInArchive", "No runnable files found in archive matching configured file types.")
                strSelectError_out = strRunError_out
                Unpack = False
                Call PopLogFileTrace
                Exit Function
            End If

        End If
        
    Loop

    Unpack = True
    Exit Function

Err_Handler:
    strRunError_out = Err.Description & " (" & Err.Number & "," & Trim$(Err.LastDllError) & ")"
    Unpack = False
    Call PopLogFileTrace

End Function

Private Sub Class_Initialize()

    Set m_objCRC = New clsCRC

    Set m_xobjFunc = New GBFuncs.clsFunctions
    m_strCommandCom = m_xobjFunc.Long2Short(Environ$("COMSPEC")) & " /c "

    m_strAppPath = App.Path
    If g_blnInVBIDE Then
        m_strAppPath = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    m_strAppPathWBS = m_strAppPath
    If Right$(m_strAppPathWBS, 1) <> "\" Then m_strAppPathWBS = m_strAppPathWBS & "\"

    m_strUnpackableFileTypes = "zip;7z;gz;gzip;tgz;bz2;bzip2;tbz2;tbz;tar;lzma;rar;cab;arj;z;taz;iso;lzh;lha"
    ma_strUnpackableFileTypes = Split(m_strUnpackableFileTypes, ";")

    m_strRepackableFileTypes = "7z;zip;gz;gzip;tgz;bz2;bzip2;tbz2;tbz;tar"
    ma_strRepackableFileTypes = Split(m_strRepackableFileTypes, ";")
    
    m_lngFinalFolderIdx = -1
    m_strFinalFolder = ""
    m_strRunnableRelativePathFile = ""

End Sub




Private Sub Class_Terminate()

    Set m_objCRC = Nothing
    Set m_xobjFunc = Nothing

End Sub


Private Sub m_xobjFuncPost_FoundFile(ByVal strPathAndFile As String)

    On Error GoTo Err_Handler

    If strPathAndFile = "__gemus" Then Exit Sub
    If (GetAttr(strPathAndFile) And vbDirectory) Then Exit Sub

    ReDim Preserve ma_udtPostUnpackedFilesSnapshot(m_lngTotalPostUnpackedFilesSnapshot)
    With ma_udtPostUnpackedFilesSnapshot(m_lngTotalPostUnpackedFilesSnapshot)
        .strCRC = m_objCRC.CalculateFile(strPathAndFile)
        .strRelativePathFile = strPathAndFile
        .strUCaseRelativePathFile = ""
    End With
    m_lngTotalPostUnpackedFilesSnapshot = m_lngTotalPostUnpackedFilesSnapshot + 1

Err_Handler:

End Sub

Private Sub m_xobjFuncPre_FoundFile(ByVal strPathAndFile As String)

    On Error GoTo Err_Handler

    If strPathAndFile = "__gemus" Then Exit Sub
    If (GetAttr(strPathAndFile) And vbDirectory) Then Exit Sub

    ReDim Preserve ma_udtPreUnpackedFilesSnapshot(m_lngTotalPreUnpackedFilesSnapshot)
    With ma_udtPreUnpackedFilesSnapshot(m_lngTotalPreUnpackedFilesSnapshot)
        .strCRC = m_objCRC.CalculateFile(strPathAndFile)
        .strRelativePathFile = strPathAndFile
        .strUCaseRelativePathFile = ""
    End With
    m_lngTotalPreUnpackedFilesSnapshot = m_lngTotalPreUnpackedFilesSnapshot + 1

Err_Handler:

End Sub


Private Sub m_xobjFuncRecursive_FoundFile(ByVal strPathAndFile As String)

    On Error GoTo Err_Handler

    If strPathAndFile = "__gemus" Then Exit Sub
    If (GetAttr(strPathAndFile) And vbDirectory) Then Exit Sub
    If LCase$(Right$(strPathAndFile, 7)) = "__gemus" Then Exit Sub

    ReDim Preserve ma_strRecursiveUnpackedFilesSnapshot(m_lngTotalRecursiveUnpackedFilesSnapshot)
    ma_strRecursiveUnpackedFilesSnapshot(m_lngTotalRecursiveUnpackedFilesSnapshot) = strPathAndFile
    m_lngTotalRecursiveUnpackedFilesSnapshot = m_lngTotalRecursiveUnpackedFilesSnapshot + 1

Err_Handler:

End Sub


