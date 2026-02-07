Attribute VB_Name = "modGlobals"
Option Explicit

Public Const gc_strAPPNAME As String = "GameBase Launcher"
Public Const gc_strAUTHOR As String = "James Burrows"
Public Const gc_strRELEASEDATE As String = "14 November 2009"
Public Const gc_sngREQ_DB_VERSION As Single = 2.5

Public g_strAPPVERSION As String

Public g_sngDBVersion As Single
Public g_lngDBMajorVersion As Long
Public g_lngDBMinorVersion As Long
Public g_strErrorUpdatingDatabase As String
Public g_lngOfficialDBUpdate As Long

Public g_dbGamebase As Database
Public g_rsGames As Recordset

Public g_strAppPathWithBackslash As String
Public g_strGameBasePath As String
Public g_strGameBasePathWBS As String
Public g_strGameBaseName As String
Public g_strGameBaseDBPath As String
Public g_strEmuFile As String
Public g_strPathFile As String
Public g_strSettingsFile As String

Public g_xobjFunc As GBFuncs.clsFunctions

Public g_blnReZip As Boolean
Public g_blnRezipAsk As Boolean
Public ge_RezipAdditional As Gemus.EnumRezipAdditionalFiles
Public g_blnZipNotify As Boolean
Public g_blnPromptForGamePathDevice As Boolean

Public Enum PathLocationEnum
    e_PathLocation_Removable
    e_PathLocation_FixedHD
    e_PathLocation_CD
    e_PathLocation_Invalid
End Enum
Public Type PathsType
    strWithBackSlash As String
    strRealPath As String
    ePathType As PathLocationEnum
    blnIsSet As Boolean
End Type
Public ga_udtRegPathGame() As PathsType
Public g_lngTotalGamePaths As Long
Public g_blnValidGamePathExists As Boolean

'Game data
Public Type GameDataType
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
Public g_udtGame As GameDataType

Public Sub GetPathInfo(strPath_in As String, udtPath_inout As PathsType)

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
                .strRealPath = Left$(g_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 1))
                iLen = Len(.strRealPath)
            End If
            
            'convert relative path to absolute
            If Mid$(.strRealPath, 2, 1) <> ":" Then
                .strRealPath = g_strAppPathWithBackslash & .strRealPath
                iLen = Len(.strRealPath)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If g_xobjFunc.DirExists(.strRealPath) = True Then
                    .strRealPath = CurDir()
                End If
            End If
                    
            'if specified path does not exist, try to validate on the current drive
            If g_xobjFunc.DirExists(.strRealPath) = False Then
                strValidateLocal = Left$(g_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 3))
                If g_xobjFunc.DirExists(strValidateLocal) = True Then
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
                If g_xobjFunc.DirExists(.strRealPath) = False Then
                    .ePathType = e_PathLocation_Invalid
                End If
            End If
        Else
            .ePathType = e_PathLocation_Invalid
        End If
    End With
    
End Sub
