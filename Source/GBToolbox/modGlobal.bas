Attribute VB_Name = "modGlobal"
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

Public Const g_blnInVBIDE As Boolean = False
Public Const gc_strAPP_PATH_WBS_WHEN_IN_VBIDE = "C:\GameBase\Build\"

Public g_strAppPathWithBackSlash As String

Public Const gc_lngMAXMEMOTEXTLENGTH As Long = 65526

'inputbox window
Public Type InputBoxWindowType
    blnAllowEmptyReturn As Boolean
    blnOKClicked As Boolean
    lngMaxLength As Long
    strCancelButtonText As String
    strCaption As String
    strInvalidChars As String
    strOKButtonText As String
    strPrompt As String
    strReturnedText As String
    strStartingText As String
End Type
Public g_udtInputBox As InputBoxWindowType
Public g_strDatabase As String

Public Const gc_sngREQ_DB_VERSION As Single = 2.5
Public g_lngDBMajorVersion As Long
Public g_lngDBMinorVersion As Long

Public g_strRC2DataFile As String
Public g_blnRC2Imported As Boolean
Public g_blnCancelPressed As Boolean
Public g_strRC2GBFolderWBS As String
Public g_strRC2Author As String
Public g_strRC2Version As String
Public g_strRC2Comment As String
Public g_blnRC2ExistingDB As Boolean
Public g_strRC2DBFile As String
Public g_strRC2DBName As String

Public Type RC2ImportOptionsType
    lngNameColIdx As Long
    lngFileColIdx As Long
    blnIgnoreDups As Boolean
    blnTosec As Boolean
    blnAddZip As Boolean
    blnScan As Boolean
    strRomPath As String
    strRomPathWBS As String
    blnMatch As Boolean
    strGameBaseName As String
    strGameBaseFolder As String
    strGameBaseFolderWBS As String
End Type
Public g_udtRC2ImportOptions As RC2ImportOptionsType

Public g_strBrowseForFolder As String

'for delete records from table
Public g_strRemoveDB As String
Public g_strDelSQL As String
Public g_lngMoveToID As Long
Public g_strMoveToName As String

'ini file functions
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'for Finding Files
Public Const MAX_PATH As Long = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
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
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Function GetDatabaseVersion(strDBFile_in As String) As Single

    Dim dbGameBase As DAO.Database
    Dim rsGameBase As Recordset
    Dim lngDBMajorVersion As Long
    Dim lngDBMinorVersion As Long
    Dim lngTestVar As Long

    lngTestVar = 1
    On Error GoTo Err_Handler

    Set dbGameBase = OpenDatabase(strDBFile_in)
    lngTestVar = 2
    Set rsGameBase = dbGameBase.OpenRecordset("Config", dbOpenSnapshot)
    With rsGameBase
        If Not .EOF Then
            .MoveFirst
            lngDBMajorVersion = !MajorVersion
            lngDBMinorVersion = !MinorVersion
            .Close
        Else
            lngDBMajorVersion = 1
            lngDBMinorVersion = 3
            .Close
        End If
    End With
    Set rsGameBase = Nothing
    GetDatabaseVersion = lngDBMajorVersion + CSng(lngDBMinorVersion / 10)
    dbGameBase.Close
    Set dbGameBase = Nothing
    Exit Function

Err_Handler:
    GetDatabaseVersion = 1.3
    If lngTestVar = 2 Then
        dbGameBase.Close
    End If
    Set dbGameBase = Nothing

End Function

Public Sub CreateFile(strFilename_in As String, strContents_in As String)

    Dim f As Integer

    f = FreeFile()
    Open strFilename_in For Output As #f
    Print #f, strContents_in
    Close #f

End Sub
Public Function GetResourceFileAsString(ByVal lngResID_in As Long) As String

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
Public Function GetINIMessage(strFile_in As String, strSection_in As String, strKey_in As String, strDefault_in As String) As String

    Dim lngRes As Long                  'return Value
    Dim strReturned As String * 4096    'max length of ini file entry
    
    'get the required message
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 4096, strFile_in)
    GetINIMessage = Left$(strReturned, lngRes)

End Function
