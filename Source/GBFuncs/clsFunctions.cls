VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2010 James Burrows
' Email: gamebase.frontend@leesti.com
' Homepage: http://www.bu22.com/
'----------------------------------------------
'   GBFuncs:  General Functions used by GB
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


'Member Variables
Private m_strAppPathWithBackslash As String
Private m_strMessagesFile As String

'Public Events
Public Event FoundFile(ByVal strPathAndFile As String)

'Header information for Sidplay tunes
Private Const mc_lngPSID_ID = 1145656144
Private Const mc_lngRSID_ID = 1145656146
Private Type SidHeadType
    PSID As Long 'file identifier
    Version As Integer ' 0001 = v1, 0002 = v2
    BinaryStart As Integer ' 0076 = v1, 007C = v2
    LoadAddr As Integer ' C64 Memory Location where to put Binary tune data
    InitAddr As Integer ' Start Address of Subroutine then inits the song
    PlayerAddr As Integer ' Address of the Player
    MaxSongsHI As Byte ' Total Number of tunes in the g_udtSid file (1 - 256)
    MaxSongsLO As Byte
    StartSongHI As Byte ' Default Sidplay Starting song
    StartSongLO As Byte
    Speed As Long ' Speed of tune
    SidName As String * 32
    Author As String * 32
    Copyright As String * 32
End Type


'------------------------------------
'API CONSTANTS/TYPES AND DECLARATIONS
'------------------------------------


'for INI file communication
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'for checking if a file exists
Private Const OFS_MAXPATHNAME = 128
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


'for getting short paths from long paths
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'for setting always on top windows
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'for deleting temporary zip files\dirs
Private Const FO_DELETE = &H3
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4
Private Const FO_COPY = &H2
Private Const FOF_RENAMEONCOLLISION = &H8
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'Drive/Volume APIs
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'for Finding Files
Private Const MAX_PATH As Long = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'for shelling files with their default program
Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
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
Private Const INFINITE = &HFFFF
Private Const SW_SHOW = 5
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long



Public Function GetBigFileSize(ByVal File As String) As Currency

' Currency handles file sizes over 2 gig
' Thanks to Richard Newcombe from codeguru.com

    Const GENERIC_READ = &H80000000
    Const FILE_SHARE_READ = &H1
    Const FILE_SHARE_WRITE = &H2
    Const OPEN_EXISTING = 3
    Dim lngHandle As Long
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim curFileSize As Currency
    
    ' Open the file
    lngHandle = CreateFile(File, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
    ' Get the file size
    lngLow = GetFileSize(lngHandle, lngHigh)
    CloseHandle lngHandle
    ' Combine the Low and High values into one currency
    ' Must use the '@' currency declaration or IDE will balk
    curFileSize = 4294967295@ * lngHigh
    If lngLow < 0 Then
        curFileSize = curFileSize + (4294967295@ + (lngLow + 1))
    Else
        curFileSize = curFileSize + lngLow
    End If
    GetBigFileSize = curFileSize

End Function
Public Function CreateFileWithContents(strPathFile_in As String, strContents_in As String, strError_out As String) As Boolean

    On Error GoTo Err_Handler
    
    Dim intFreeFileNumber As Integer
    
    strError_out = ""
    
    intFreeFileNumber = FreeFile()
    Open strPathFile_in For Output As #intFreeFileNumber
    Print #intFreeFileNumber, strContents_in
    Close #intFreeFileNumber
    
    CreateFileWithContents = True
    Exit Function

Err_Handler:
    CreateFileWithContents = False
    strError_out = Err.Description & " (" & Err.Number & "," & Trim$(Err.LastDllError) & ")"

End Function

Public Sub INI_DeleteKey(strSection_in As String, strKey_in As String, strFile_in As String)

    Call WritePrivateProfileString(strSection_in, strKey_in, vbNullString, strFile_in)

End Sub
Public Function INI_DeleteNumberedValues(strSection_in As String, ByVal lngFirstNumberToDelete_in As Long, strFile_in As String) As Boolean

    Dim lngIdx As Long

    For lngIdx = lngFirstNumberToDelete_in To 50
        Call INI_DeleteKey(strSection_in, Trim$(lngIdx), strFile_in)
    Next lngIdx
    INI_DeleteNumberedValues = INI_WriteSetting(strSection_in, Trim$(lngFirstNumberToDelete_in), "", strFile_in)

End Function

Public Sub INI_DeleteSection(strSection_in As String, strFile_in As String)

    Call WritePrivateProfileString(strSection_in, vbNullString, vbNullString, strFile_in)

End Sub

Public Function GetINIMessage(strSection_in As String, strKey_in As String, strDefault_in As String) As String

    '------------------------------------------------------------------------
    'Reads a text item from the GB64Text.ini file
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
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, m_strMessagesFile)
    GetINIMessage = Left$(strReturned, lngRes)

End Function
Public Function INI_GetSetting(strSection_in As String, strKey_in As String, strDefault_in As String, strFile_in As String) As String

    Dim lngRes As Long                  'return Value
    Dim strReturned As String * 2096    'max length of ini file entry
    
    'get the required message
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, strFile_in)
    INI_GetSetting = Left$(strReturned, lngRes)

End Function

Public Property Let MessagesFile(strFile As String)

    m_strMessagesFile = strFile

End Property
Public Property Get MessagesFile() As String

    MessagesFile = m_strMessagesFile

End Property

Public Function RTrimWhitespace(strToTrim_in As String) As String

    'trims whitespace characters from the end of a string
    'returning the trimmed version

    Dim lngIdx As Long
    Dim lngLen As Long
    Dim lngTrimmedLen As Long
    Dim strChar As String

    lngLen = Len(strToTrim_in)
    lngTrimmedLen = lngLen

    For lngIdx = lngLen To 1 Step -1
        strChar = Mid$(strToTrim_in, lngIdx, 1)
        If strChar = " " Or strChar = vbTab Or strChar = vbCr Or strChar = vbLf Then
            lngTrimmedLen = lngTrimmedLen - 1
        Else
            Exit For
        End If
    Next lngIdx

    RTrimWhitespace = Left$(strToTrim_in, lngTrimmedLen)

End Function

Public Sub SetGBTextINIFile(strDBFolderFilename_in As String)

    If FileExists(strDBFolderFilename_in, False, "") = True Then
        m_strMessagesFile = strDBFolderFilename_in
    End If

End Sub

Public Function ShellExecuteFile(strFile_in As String, strCLP_in As String, blnWait_in As Boolean) As Boolean

    '--------------------------------------------------------------
    'shells a program and optionally waits until it's finished
    '
    'Inputs:    strFile_in:     Filename of file to run
    '           strCLP_in:      Any command line parameters
    '           blnWait_in:     Wait until program is finished flag
    '--------------------------------------------------------------

    Dim SEI As SHELLEXECUTEINFO     'api structure

    'Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI
        .hwnd = GetDesktopWindow()
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
        ShellExecuteFile = False
    Else
        ShellExecuteFile = True
        'wait till finished?
        If blnWait_in = True Then
            If SEI.hProcess <> 0 Then
                Call WaitForSingleObject(SEI.hProcess, INFINITE)
            End If
        End If
    End If

End Function

Public Function INI_WriteSetting(strSection_in As String, strKey_in As String, strValue_in As String, strFile_in As String) As Boolean

    Dim lngRes As Long 'return Value
    
    'write the required value
    lngRes = WritePrivateProfileString(strSection_in, strKey_in, strValue_in, strFile_in)
    INI_WriteSetting = (lngRes <> 0)

End Function
Public Sub WriteToDebugLog(strMessage_in As String)

    '-------------------------------------------------------------
    'Appends the incoming message to the end of the debug log file
    '
    'Inputs:    strMessage_in       The message to write!
    '-------------------------------------------------------------
    
    Dim intFileNum As Integer       'file number
    
    'set error handler just incase
    On Error GoTo Err_Handler
    
    'open the log file
    intFileNum = FreeFile()
    Open m_strAppPathWithBackslash & "debug.log" For Append As #intFileNum
    
    'add the message
    Print #intFileNum, strMessage_in
    
    'close the file
    Close #intFileNum
    
    'release the error handler and exit
    On Error GoTo 0
    Exit Sub
    
Err_Handler:

End Sub


Public Function FindFile(strFilename_in As String, ByVal strPath_in As String) As String

    '---------------------------------------------------------------
    'Finds a File in a given path
    '
    'Inputs:    strFilename_in:       The filename of the file to find
    '           strPath_in:           The path to search
    '
    'Returns:   Path\Filename of the found file or "" if not found
    '
    'Remarks:   This function is recursive
    '---------------------------------------------------------------

    Dim hFile As Long               'File handle
    Dim strTestDir As String        'test string if file found is a dir
    Dim udtWFD As WIN32_FIND_DATA   'API structure
    Dim lngRetVal As Long           'API return value
    Dim strAttempt As String        'Recursive return value for this function
    Dim strFFPath As String         'for FindFirstFile() API
    
    Dim strSPPath As String         ' for SearchPath API
    Dim strSPFilename As String     '          "
    Dim lngSPBufferLen As Long      '          "
    Dim strSPBuffer As String       '          "
    Dim lpFilePart As String        '          "
    
    'set up the search variables
    strFFPath = GetFolderWithBackslash(strPath_in) & "*.*" & Chr$(0)
    strSPPath = strPath_in & Chr$(0)
    strSPFilename = strFilename_in & Chr$(0)
    strSPBuffer = String$(MAX_PATH, 0)
    lngSPBufferLen = Len(strSPBuffer)
    
    'can windows find it simply in the root specified folder?
    lngRetVal = SearchPath(strSPPath, strSPFilename, vbNullString, lngSPBufferLen, strSPBuffer, lpFilePart)
    If lngRetVal Then
        'found it so exit!
        FindFile = StripNull(strSPBuffer)
        Exit Function
    End If

    '''''''''''''''''''''''''''''''
    'Need to do a recursive search!
    '''''''''''''''''''''''''''''''
    
    'start the search off...
    hFile = FindFirstFile(strFFPath, udtWFD)
    
    Do
        'check if found item is a directory
        If udtWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            strTestDir = StripNull(udtWFD.cFileName)
            If Not (strTestDir = "." Or strTestDir = "..") Then
                'Don't look for hidden or system directories
                If Not (udtWFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
                    'Search directory recursively
                    strAttempt = FindFile(strFilename_in, GetFolderWithBackslash(strPath_in) & strTestDir)
                    If strAttempt <> "" Then
                        FindFile = strAttempt
                        Exit Do
                    End If
                End If
            End If
        End If
        'next file
        udtWFD.cFileName = ""
        lngRetVal = FindNextFile(hFile, udtWFD)
    Loop Until lngRetVal = 0

    'tidy up
    Call FindClose(hFile)
    
End Function

Private Function StripNull(ByVal strSource_in As String) As String

    '----------------------------------------------------------------------------
    'Used by FindFile() to strip the last NULL character off the passed in string
    '
    'Inputs:    strSource_in:   The source string to strip the NULL from
    '
    'Returns:   copy of the source string but with no NULL char at the end!
    '----------------------------------------------------------------------------

    Dim lngPos As Integer   'position of NULL char
    
    'get null position in the string
    lngPos = InStr(strSource_in, Chr$(0))

    'amend the string if needed
    If lngPos > 0 Then
        StripNull = Left$(strSource_in, lngPos - 1)
    Else
        StripNull = strSource_in
    End If

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

Public Function FindFiles(strPath_in As String, strSearchPattern_in As String)

    '----------------------------------------------------------------
    'Finds files matching a specified search pattern in a given path
    '
    'Inputs:    strPath_in:             The path to search
    '           strSearchPattern_in:    The pattern to search for
    '
    'Returns:   The accumulated file-size of all found files matching
    '           the specified search criteria
    '
    'Remarks:   This function is recursive and sends an Event to the
    '           calling client with each file found.
    '           The client must declare an instance of this class
    '           WithEvents if they want to use this function.
    '----------------------------------------------------------------
    
    Dim strFilename As String           'Working filename variable
    Dim strDirName As String            'Working SubDirectory variable
    Dim a_strDirNames() As String       'Buffer for directory name entries
    Dim intSubDirCount As Integer       'Number of directories in current path
    Dim intSubDirIdx As Integer         'loop index
    Dim hSearch As Long                 'Search Handle
    Dim udtWFD As WIN32_FIND_DATA       'structure for API
    Dim intContinue As Integer          'continue searching flag
    
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
        Loop
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(strPath_in & strSearchPattern_in, udtWFD)
    intContinue = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While intContinue
            strFilename = StripNulls(udtWFD.cFileName)
            If (strFilename <> ".") And (strFilename <> "..") Then
                'add to accumulated filesize
                FindFiles = FindFiles + (udtWFD.nFileSizeHigh * MAXDWORD) + udtWFD.nFileSizeLow
                RaiseEvent FoundFile(strPath_in & strFilename)
            End If
            'Get next file
            intContinue = FindNextFile(hSearch, udtWFD)
        Wend
        'tidy up
        intContinue = FindClose(hSearch)
    End If
    
    'Any sub-directories?
    If intSubDirCount > 0 Then
        'Recursively process them
        For intSubDirIdx = 0 To intSubDirCount - 1
            'add to accumulated filesize
            FindFiles = FindFiles + FindFiles(strPath_in & a_strDirNames(intSubDirIdx) & "\", strSearchPattern_in)
        Next intSubDirIdx
    End If
    
End Function
Private Function GetFolderWithBackslash(strPath_in As String) As String

    '---------------------------------------------------
    'Adds a backslash on the end of a path, if required.
    '
    'Inputs:    strPath_in:     Path to process
    '
    'Returns:   the path with a backslash on the end
    '---------------------------------------------------

    'invalid path
    If strPath_in = "" Then
        Exit Function
    End If
    
    'add backslash if needed
    If Right$(strPath_in, 1) = "\" Then
        GetFolderWithBackslash = strPath_in
    Else
        GetFolderWithBackslash = strPath_in & "\"
    End If

End Function

Public Sub GetFileAndPath(strPathAndFile_in As String, strPath_out As String, strFilename_out As String)
    
    '-----------------------------------------------------------------
    'Splits a path\filename up into it's separate parts
    '
    'Inputs:    strPathAndFile_in:      The path and filename to split
    '
    'Outputs:   strPath_out:        The path segment
    '           strFilename_out:    The filename segment
    '
    'Returns:   None
    '
    'Remarks:   If there is no path (i.e. no "\" in the strPathAndFile
    '           parameter) then this function returns just a filename.
    '-----------------------------------------------------------------
    
    Dim lngCharIdx As Long      'loop index
    
    If InStr(strPathAndFile_in, "\") = 0 Then
        'no path - assume just a file
        strPath_out = ""
        strFilename_out = strPathAndFile_in
    Else
        'get position of "\" splitter character
        strPath_out = Left$(strPathAndFile_in, InStrRev(strPathAndFile_in, "\"))
        strFilename_out = Right$(strPathAndFile_in, Len(strPathAndFile_in) - Len(strPath_out))
        If Len(strPath_out) > 3 Then
            strPath_out = Left$(strPath_out, Len(strPath_out) - 1)
        End If
    End If

End Sub
Public Function GetSidInfo(strSidFile_in As String, strName_out As String, strAuthor_out As String, strCopyright_out As String, lngTotalSongs_out As Long, lngDefaultSong_out As Long) As Boolean
    
    '---------------------------------------------------------------------
    'returns info from a passed SidFile
    '
    'Inputs:    strSidFile_in:          The Sidtune path\filename to get
    '                                   the info for
    '
    'Outputs:   strName_out:            Internal Sidtune Name
    '           strAuthor_out:          Internal Sidtune Author
    '           strCopyright_out:       Internal Sidtune Copyright Info
    '           lngTotalSongs_out:      Number of songs in the Sidtune
    '           lngDefaultSong_out:     Index of the default starting song
    '
    'Returns:   False if the Sid File is not found or Sid is not pSID
    '           True otherwise
    '---------------------------------------------------------------------
    
    Dim udtSid As SidHeadType       'Sid-Tune file header structure
    Dim intFileNum As Integer       'file number
    Dim lngPSID_ID As Long
    
    'error handler just incase
    On Error GoTo Err_Handler
    
    'assume failure or imminent error
    GetSidInfo = False

    If FileExists(strSidFile_in, False, "") = False Then
        Exit Function
    End If

    'open the file
    intFileNum = FreeFile()
    Open strSidFile_in For Binary As #intFileNum
    
    'attempt to get info
    Get #intFileNum, , lngPSID_ID
    If lngPSID_ID = mc_lngPSID_ID Or lngPSID_ID = mc_lngRSID_ID Then
        'file is a PSID or RSID!
        Get #intFileNum, 1, udtSid
        strName_out = udtSid.SidName
        strAuthor_out = udtSid.Author
        strCopyright_out = udtSid.Copyright
        lngTotalSongs_out = 256 * udtSid.MaxSongsHI + udtSid.MaxSongsLO
        lngDefaultSong_out = 256 * udtSid.StartSongHI + udtSid.StartSongLO
        GetSidInfo = True
    End If
    
    'close file, release error handler and exit
    Close #intFileNum
    On Error GoTo 0
    Exit Function

Err_Handler:
    'release error handler and return failure
    On Error GoTo 0
    GetSidInfo = False
    
End Function

Public Function GetSidSongInfo(ByVal strSidFile As String, intTotalSongs As Integer, intDefaultSong As Integer) As Boolean
    
    'returns info from a passed SidFile
    'returns False if the g_udtSid File is not found or g_udtSid is not pSID
    
    Dim g_udtSid As SidHeadType
    Dim f As Integer
    Dim intID As Long
    
    On Error GoTo Err_Handler
    
    GetSidSongInfo = False
    f = FreeFile()
    Open strSidFile For Binary As #f
    Get #f, , intID
    
    If intID = mc_lngPSID_ID Or intID = mc_lngRSID_ID Then
        Get #f, 1, g_udtSid
        intTotalSongs = 256 * g_udtSid.MaxSongsHI + g_udtSid.MaxSongsLO
        intDefaultSong = 256 * g_udtSid.StartSongHI + g_udtSid.StartSongLO
        GetSidSongInfo = True
    End If
    
    Close #f
    Exit Function

Err_Handler:
    GetSidSongInfo = False
    
End Function
Public Function SaveDefaultSidSong(ByVal strSidFile As String, intDefaultSong As Integer) As Boolean
    
    'saves a new default g_udtSid song for the passed in g_udtSid file
    
    Dim g_udtSid As SidHeadType
    Dim f As Integer
    Dim intID As Long
    Dim Att As VbFileAttribute
    
    On Error GoTo Err_Handler
    
    SaveDefaultSidSong = False
    
    If DriveTypeIsCDRom(Left$(strSidFile, 2)) = True Then
        GoTo Err_Handler
    End If
    Att = GetAttr(strSidFile)
    If (Att And vbReadOnly) Then
        GoTo Err_Handler
    End If
    
    f = FreeFile()
    Open strSidFile For Binary As #f
    Get #f, , intID
    
    If intID = mc_lngPSID_ID Then
        Get #f, 1, g_udtSid
        If intDefaultSong = 256 Then
            g_udtSid.StartSongHI = 1
            g_udtSid.StartSongLO = 0
        Else
            g_udtSid.StartSongHI = 0
            g_udtSid.StartSongLO = intDefaultSong
        End If
        Put #f, 1, g_udtSid
        SaveDefaultSidSong = True
    End If
    
    Close #f
    Exit Function

Err_Handler:
    SaveDefaultSidSong = False
    
End Function
Public Function GetSidAuthor(ByVal strSidFile As String) As String
    
    'returns Author from a passed SidFile
    'returns empty String if an error
    
    Dim Author As String * 32
    Dim f As Integer
    
    On Error GoTo Err_Handler

    f = FreeFile()
    Open strSidFile For Binary As #f
    Get #f, 55, Author
    GetSidAuthor = Author
    Close #f
    Exit Function

Err_Handler:
    GetSidAuthor = ""
    
End Function
Public Function JustFileName(ByVal strPathAndFile As String) As String
    
    'returns just the filename of a given path\file
    
    Dim Letter As Integer
    Dim Lenny As Integer
    
    Lenny = Len(strPathAndFile)
    JustFileName = strPathAndFile
    If InStr(strPathAndFile, "\") = 0 Then Exit Function
    For Letter = Lenny To 1 Step -1
        If Mid$(strPathAndFile, Letter, 1) = "\" Then
            Exit For
        End If
    Next Letter
    JustFileName = Right$(strPathAndFile, Lenny - Letter)
    
End Function
Public Sub SetAlwaysOnTop(ByVal SetOn As Boolean, ByVal WindowHwnd As Long)
    
    If SetOn = False Then
        SetWindowPos WindowHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    Else
        SetWindowPos WindowHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    End If
    
End Sub


Public Function WaitForCDRom(ByVal RootDrive As String, ByVal CDRomName As String, strAppName As String) As Boolean
    
    Dim VolumeName As String
    
    Do
        DoEvents
        VolumeName = GetVolumeName(RootDrive)
        If VolumeName <> CDRomName Then
            If MsgBox("Insert CD: '" & CDRomName & "' into Drive " & UCase(RootDrive), vbOKCancel, strAppName) = vbCancel Then
                WaitForCDRom = False
                Exit Function
            End If
        Else
            'correct CD is inserted --> continue
            WaitForCDRom = True
            Exit Do
        End If
    Loop
    
End Function
Public Function WaitForCDRomPath(ByVal strPath As String, strAppName As String) As Boolean
    
    Do
        DoEvents
        If Me.DirExists(strPath) = False Then
            If MsgBox("Insert the CD for Path: '" & strPath & "'", vbOKCancel, strAppName) = vbCancel Then
                WaitForCDRomPath = False
                Exit Function
            End If
        Else
            'correct CD is inserted --> continue
            WaitForCDRomPath = True
            Exit Do
        End If
    Loop
    
End Function
Public Function WaitForCDRomFile(ByVal strPath As String, strAppName As String) As Boolean
    
    Do
        DoEvents
        If Me.FileExists(strPath, False, strAppName) = False Then
            If MsgBox("Insert the CD for File: '" & strPath & "'", vbOKCancel, strAppName) = vbCancel Then
                WaitForCDRomFile = False
                Exit Function
            End If
        Else
            'correct CD is inserted --> continue
            WaitForCDRomFile = True
            Exit Do
        End If
    Loop
    
End Function
Public Function GetVolumeName(ByVal RootPathName As String) As String

    'This function returns the name of a CDRom whose path is specified
    
    Dim Counter As Integer
    Dim X As Long
    Dim DriveName As String
    Dim DriveNameCompare As String
    Dim FileSysName As String
    Dim lngMaxCompLen As Long
    Dim lngFlags As Long
    Dim RootPathNameSend As String
    
    DriveName = Space(128)
    FileSysName = Space(128)
    RootPathNameSend = RootPathName & "\"
    DriveNameCompare = DriveName
    X = GetVolumeInformation(RootPathNameSend, DriveName, 128, vbNull, lngMaxCompLen, lngFlags, FileSysName, 128)
    If DriveName = DriveNameCompare Then
        GetVolumeName = ""
    Else
        Counter = 1
        GetVolumeName = ""
        Do While Counter < Len(DriveName)
            GetVolumeName = GetVolumeName + Mid(DriveName, Counter, 1)
            Counter = Counter + 1
            If Mid(DriveName, Counter, 1) = Chr(0) Then
                Exit Do
            End If
        Loop
    End If
    
End Function

Public Function Long2Short(ByVal S As String) As String
    
    If S = "" Then
        Long2Short = ""
        Exit Function
    End If
    Dim Buff As String
    Dim r As Integer
    Buff = Space(256)
    r = GetShortPathName(S, Buff, 256)
    Long2Short = Left(Buff, r)
    
End Function
Public Sub DeletePathSilent(ByVal strDestPath As String)
    
    Dim ShOp As SHFILEOPSTRUCT
    Dim RetVal As Long
    
    ShOp.fFlags = FOF_SILENT Or FOF_NOCONFIRMATION
    ShOp.pFrom = strDestPath
    ShOp.wFunc = FO_DELETE
    RetVal = SHFileOperation(ShOp)
    DoEvents
    
End Sub
Public Function DriveTypeIsCDRom(ByVal RootPathName As String) As Boolean
    
    Dim a As Long
    
    a = GetDriveType(RootPathName & "\")
    If a = 5 Or a = 1 Then
        DriveTypeIsCDRom = True
    Else
        DriveTypeIsCDRom = False
    End If
    
End Function

Public Function FileExists(ByVal Filename As String, ByVal DisplayMsgBox As Boolean, ByVal MBCaption As String) As Boolean
    
    On Error GoTo File_Does_Not_Exist
    
    Dim RetCode As Long
    Dim OpenFileStructure As OFSTRUCT
    Const OF_EXIST = &H4000
    Const FILE_NOT_FOUND = 2
    Const PATH_NOT_FOUND = 3
    
    If Filename = "" Then
        FileExists = False
        If DisplayMsgBox = True Then
            MsgBox "FILE NOT FOUND: NO FILENAME PASSED!", vbExclamation, MBCaption
        End If
        Exit Function
    End If
    
    RetCode = OpenFile(Filename, OpenFileStructure, OF_EXIST)
    
    If OpenFileStructure.nErrCode = FILE_NOT_FOUND Then
        FileExists = False
        If DisplayMsgBox = True Then
            MsgBox "FILE NOT FOUND: " & Filename, vbExclamation, MBCaption
        End If
    ElseIf OpenFileStructure.nErrCode = PATH_NOT_FOUND Then
        FileExists = False
        If DisplayMsgBox = True Then
            MsgBox "PATH NOT FOUND: " & Filename, vbExclamation, MBCaption
        End If
    ElseIf OpenFileStructure.nErrCode = 5 Then
        FileExists = False
        If DisplayMsgBox = True Then
            MsgBox "Device Not Ready for " & Filename, vbExclamation, MBCaption
        End If
    Else
        If RetCode = 1 Then
            FileExists = True
        Else
            'check file exists another way, as path might exceed 128 chars (max allowed by OpenFile)
            RetCode = GetAttr(Filename) 'goes to error handler if can't get the attribute
            FileExists = True
        End If
    End If
    
    Exit Function

File_Does_Not_Exist:
    FileExists = False
    
End Function
Public Function DirExists(ByVal strDir As String) As Boolean
    
    On Error GoTo No_Dir
    ChDir strDir
    DirExists = True
    Exit Function
    
No_Dir:
    DirExists = False
    
End Function

Private Sub Class_Initialize()

    'set the folder
    m_strAppPathWithBackslash = App.Path
    If Right$(m_strAppPathWithBackslash, 1) <> "\" Then
        m_strAppPathWithBackslash = m_strAppPathWithBackslash & "\"
    End If
    
    'set the messages file
    m_strMessagesFile = m_strAppPathWithBackslash & "GBText.ini"

End Sub


