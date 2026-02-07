VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GBRegRun"
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

Private Const m_blnInVBIDE As Boolean = False
Private Const mc_strAPP_PATH_WBS_WHEN_IN_VBIDE = "C:\GameBase\Build\"
Private Const mc_strINIFILENAME = "GBRegRun.ini"
Private Const mc_strAPPNAME As String = "GBRegRun"
Private Const mc_strAUTHOR As String = "James Burrows"

Private m_strAppPathWithBackslash As String
Private m_strAPPVERSION As String
Private m_strCommandCom As String

Private Enum GBRegRunJobTypeEnum
    e_UnRegister
    e_Register
    e_Run_no_Wait
    e_Run_and_Wait
End Enum
Private Type GBRegRunJobType
    eJobType As GBRegRunJobTypeEnum
    strRelativePathFile As String
End Type
Private ma_udtJobs() As GBRegRunJobType
Private m_lngTotalJobs As Long
Private m_blnSilentMode As Boolean

'winapi
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
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
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
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_SILENT = &H4
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const INFINITE = &HFFFF
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Function GetConfig() As Boolean

    On Error GoTo Err_Handler

    Dim strIniPathFile As String
    Dim f As Integer
    Dim strLine As String
    Dim strSection As String
    Dim strKey As String
    Dim strVal As String
    Dim strAppFile As String
    Dim strAppPath As String
    Dim lngIdx As Long
    Dim lngPos As Long

    GetConfig = False

    m_lngTotalJobs = 0
    ReDim ma_udtJobs(0)

    strIniPathFile = m_strAppPathWithBackslash & mc_strINIFILENAME

    If Not FileExists(strIniPathFile, False, "") Then Exit Function

    strSection = ""

    f = FreeFile()
    Open strIniPathFile For Input As #f
    Do Until EOF(f)
        Line Input #f, strLine
        strLine = Trim$(strLine)
        If Left$(strLine, 1) = "[" And Right$(strLine, 1) = "]" Then
            'new section
            If Len(strLine) > 2 Then
                strSection = Mid$(strLine, 2, Len(strLine) - 2)
            End If
        ElseIf strSection = "Files" Or strSection = "Settings" Then
            'in correct section: grab key=value pair and add to array
            lngPos = InStr(strLine, "=")
            If lngPos > 1 And Left$(strLine, 1) <> ";" Then
                strKey = Trim$(Left$(strLine, lngPos - 1))
                strVal = Right$(strLine, Len(strLine) - lngPos)
                strVal = Trim$(strVal)
                If strVal <> "" Then
                    If strSection = "Files" Then
                        If strKey = "reg" Or strKey = "unreg" Or strKey = "run" Or strKey = "wait" Then
                            ReDim Preserve ma_udtJobs(m_lngTotalJobs)
                            With ma_udtJobs(m_lngTotalJobs)
                                .strRelativePathFile = strVal
                                If strKey = "reg" Then
                                    .eJobType = e_Register
                                ElseIf strKey = "run" Then
                                    .eJobType = e_Run_no_Wait
                                ElseIf strKey = "wait" Then
                                    .eJobType = e_Run_and_Wait
                                ElseIf strKey = "unreg" Then
                                    .eJobType = e_UnRegister
                                End If
                                m_lngTotalJobs = m_lngTotalJobs + 1
                            End With
                        End If
                    ElseIf strSection = "Settings" Then
                        If strKey = "mode" Then
                            If strVal = "silent" Then
                                m_blnSilentMode = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    Close #f

    If m_lngTotalJobs = 0 Then
        MsgBox "No jobs to run found in " & mc_strINIFILENAME, vbInformation, mc_strAPPNAME
        Exit Function
    End If

    'check they exist
    ChDrive Left$(m_strAppPathWithBackslash, 3)
    ChDir m_strAppPathWithBackslash
    For lngIdx = 0 To m_lngTotalJobs - 1
        If Not FileExists(ma_udtJobs(lngIdx).strRelativePathFile, (m_blnSilentMode = False), mc_strAPPNAME) Then
            Exit Function
        End If
    Next lngIdx

    GetConfig = True
    Exit Function

Err_Handler:
    GetConfig = False

End Function

Private Sub Form_Load()

    Dim lngIdx As Long
    Dim lngHandle As Long
    Dim lngExitCode As Long
    Dim dblProcessID As Double

    'app path and version
    m_strAppPathWithBackslash = App.Path
    If Right$(m_strAppPathWithBackslash, 1) <> "\" Then m_strAppPathWithBackslash = m_strAppPathWithBackslash & "\"

    If m_blnInVBIDE = True Then
        m_strAppPathWithBackslash = mc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If

    m_strAPPVERSION = Format$(Val(App.Major), "0") & "." & Format$(Val(App.Minor), "0")

    m_strCommandCom = Long2Short(Environ$("COMSPEC")) & " /c "
    
    m_blnSilentMode = False

    If Not GetConfig() Then
        MsgBox "Unable to get config from " & mc_strINIFILENAME, vbInformation, mc_strAPPNAME
        End
    End If

    For lngIdx = 0 To m_lngTotalJobs - 1
        With ma_udtJobs(lngIdx)
            If .eJobType = e_Register Then
            
                regSvr32 .strRelativePathFile, True
                
            ElseIf .eJobType = e_UnRegister Then
            
                regSvr32 .strRelativePathFile, False
                
            ElseIf .eJobType = e_Run_no_Wait Then
            
                dblProcessID = Shell(.strRelativePathFile, vbNormalFocus)
                
            ElseIf .eJobType = e_Run_and_Wait Then

                dblProcessID = Shell(.strRelativePathFile, vbNormalFocus)
                lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dblProcessID)
                Do
                    Call GetExitCodeProcess(lngHandle, lngExitCode)
                    DoEvents
                    If lngExitCode <> STATUS_PENDING Then
                        Exit Do
                    End If
                Loop
                Call CloseHandle(lngHandle)

            End If
        End With
        DoEvents
    Next lngIdx

    End

End Sub
Private Sub regSvr32(strRelativePathFile_in As String, blnRegister_in As Boolean)
    
    Dim sNull As String
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    Dim lngRetVal As Long
    Dim strCommandLine As String
    Dim strUFlag As String
    Dim strSFlag As String
    
    Dim strShortRawFilename As String
    Dim strShortFullFile As String
    Dim strShortPath As String

    strSFlag = ""
    strUFlag = ""
    
    If Not blnRegister_in Then strUFlag = "/u "
    If m_blnSilentMode Then strSFlag = "/s "

    With sInfo
        .cbReserved2 = 0
        .lpReserved2 = vbNull
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .cb = Len(sInfo)
    End With

    strCommandLine = m_strCommandCom & " regsvr32.exe " & strSFlag & strUFlag & """" & strRelativePathFile_in & """"
    lngRetVal = CreateProcess(sNull, strCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    lngRetVal = WaitForSingleObject(pInfo.hProcess, INFINITE)

End Sub
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
Private Function FileExists(ByVal Filename As String, ByVal DisplayMsgBox As Boolean, ByVal MBCaption As String) As Boolean
    
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
