Attribute VB_Name = "modFuncs"
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

' set this to True when running GB from VB, False when compiling for the public!
Public Const g_blnInVBIDE As Boolean = False
Public Const gc_strAPP_PATH_WBS_WHEN_IN_VBIDE = "C:\GameBase\Build\"

Public g_intLogFileNum As Integer
Public g_strLogFileTrace As String
Public ga_strLogFileTrace() As String
Public Const gc_LOG_FILE_TRACE_SEPARATOR = "    "

Public g_blnScriptSubclassed As Boolean
Public g_blnArchiveSubclassed As Boolean

Public g_strShowKV As String

Public Type ArchiveType
    strArchivePathFile As String
    strTempPathWBS As String
    a_strRunnableFileTypes() As String
    blnOKPressed As Boolean
    strRunnableFile As String
End Type
Public g_udtArchiveBrowse As ArchiveType

Public Type EmulatorPropsType
    strName As String
    strFullPath As String
    strScript As String
    blnUseShort As Boolean
    strFiles() As String
    lngNumFiles As Long
    strTypes As String
    blnUseDefaultTypes As Boolean
    strGBName As String
    strEmuFile As String
    lngWait As Long
End Type
Public g_udtEmulatorProps As EmulatorPropsType
Public g_blnEmusManaged As Boolean
Public g_blnFileTypesChanged As Boolean

Public g_blnRezippable As Boolean

'the emulator setup wizard
Public Type WizardDATLineType
    strEmulatorName As String
    strURL As String
    strInfo As String
    strFilename As String
    a_strCRC32() As String
    lngTotalCRCs As Long
    a_strFileTypes() As String
    lngTotalFileTypes As Long
    blnUseShort As Boolean
    strScriptFile As String
    a_strAssFiles() As String
    lngTotalAssFiles As Long
    strFound As String
End Type
Public Type WizardDATSectionType
    strName As String
    a_udtEmulators() As WizardDATLineType
    lngTotalEmulators As Long
End Type
Public Type WizardDATFileType
    strMessagesFile As String
    blnAborted As Boolean
    blnGlobalDAT As Boolean
    strDATFile As String
    strEmuFile As String
    strGBFolder As String
    strScriptFolder As String
    sngDATVersion As Single
    strLastUpdated As String
    strAuthor As String
    strHomePage As String
    a_udtSections() As WizardDATSectionType
    lngTotalSections As Long
End Type
Public g_udtWizard As WizardDATFileType
Public g_strBrowseForFolder As String

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

Public g_strScriptSaveVar As String

Public Type ScriptSelectorType
    blnOKClicked As Boolean
    strSelectedFile As String
    strInitPath As String
End Type
Public g_udtScriptSelectorType As ScriptSelectorType

Public g_strManageEmuGBName As String
Public g_strManageEmuFile As String
Public g_strFileTypesGBName As String
Public g_strFileTypesEmuFile As String

Public Type ManualGEMUSFieldsFriend
    lngControl As Long
    lngPlayers As Long
    lngPalNTSC As Long
    blnTDE As Boolean
    strComment As String
    strVComment As String
    strNotes As String
    blnOKClicked As Boolean
    blnDefaults As Boolean
End Type
Public g_udtManualFields As ManualGEMUSFieldsFriend
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public defScriptWindowProc As Long
Public lngScriptWindowMinX As Long
Public lngScriptWindowMinY As Long
Public lngScriptWindowMaxX As Long
Public lngScriptWindowMaxY As Long
Public defArchiveWindowProc As Long
Public lngArchiveWindowMinX As Long
Public lngArchiveWindowMinY As Long
Public lngArchiveWindowMaxX As Long
Public lngArchiveWindowMaxY As Long
Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24
Public Type POINTAPI
    x As Long
    Y As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub LogStatus(strStatus As String)

    On Error GoTo Err_Handler
    
    If g_intLogFileNum > 0 Then
        Print #g_intLogFileNum, g_strLogFileTrace & strStatus
    End If

    Exit Sub

Err_Handler:
    g_intLogFileNum = 0

End Sub
Public Sub PushLogFileTrace(ByVal strFuncName As String)

    On Error GoTo Err_Handler

    Dim lngNumItems As Long
    Dim lngIdx As Long
    
    lngNumItems = UBound(ga_strLogFileTrace)
    ReDim Preserve ga_strLogFileTrace(lngNumItems + 1)
    ga_strLogFileTrace(lngNumItems + 1) = strFuncName

    g_strLogFileTrace = ""
    For lngIdx = 1 To lngNumItems + 1
        g_strLogFileTrace = g_strLogFileTrace & gc_LOG_FILE_TRACE_SEPARATOR
    Next lngIdx
    
    Call LogStatus("-----------------" & strFuncName & "-----------------")
    
Err_Handler:

End Sub

Public Sub PopLogFileTrace()

    On Error GoTo Err_Handler

    Dim lngNumItems As Long
    Dim lngIdx As Long
    
    lngNumItems = UBound(ga_strLogFileTrace)
    Call LogStatus("-----------------" & String(Len(ga_strLogFileTrace(lngNumItems)), "-") & "-----------------")
    
    ReDim Preserve ga_strLogFileTrace(lngNumItems - 1)

    g_strLogFileTrace = ""
    For lngIdx = 1 To lngNumItems - 1
        g_strLogFileTrace = g_strLogFileTrace & gc_LOG_FILE_TRACE_SEPARATOR
    Next lngIdx

Err_Handler:

End Sub
Sub WaitSeconds(ByVal lngNumSeconds_in As Long)

    Dim lngFinishTime As Long

    lngFinishTime = GetTickCount + (lngNumSeconds_in * 1000)
    While GetTickCount() < lngFinishTime
        DoEvents
    Wend

End Sub


Public Sub ScriptSubClass(hWnd As Long)
    
    'assign our own window message procedure (ScriptWindowProc)
  
    On Error Resume Next
    defScriptWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ScriptWindowProc)
    g_blnScriptSubclassed = True
    
End Sub
Public Sub ArchiveSubClass(hWnd As Long)
    
    'assign our own window message procedure (ArchiveWindowProc)
  
    On Error Resume Next
    defArchiveWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ArchiveWindowProc)
    g_blnArchiveSubclassed = True
    
End Sub
Public Sub ScriptUnSubClass(hWnd As Long)
    
    'restore the default message handling before exiting
    If defScriptWindowProc Then
        SetWindowLong hWnd, GWL_WNDPROC, defScriptWindowProc
        defScriptWindowProc = 0
        g_blnScriptSubclassed = False
    End If
    
End Sub
Public Sub ArchiveUnSubClass(hWnd As Long)
    
    'restore the default message handling before exiting
    If defArchiveWindowProc Then
        SetWindowLong hWnd, GWL_WNDPROC, defArchiveWindowProc
        defArchiveWindowProc = 0
        g_blnArchiveSubclassed = False
    End If
    
End Sub
Public Function ScriptWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hWnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmEmulatorProps.hWnd
            On Error Resume Next
            'our form-specific handler
            Select Case uMsg
                Case WM_GETMINMAXINFO
                Dim MMI As MINMAXINFO
                CopyMemory MMI, ByVal lParam, LenB(MMI)
                'set the MINMAXINFO data to the
                'minimum and maximum values set
                'by the option choice
                With MMI
                    .ptMinTrackSize.x = lngScriptWindowMinX
                    .ptMinTrackSize.Y = lngScriptWindowMinY
                    .ptMaxTrackSize.x = lngScriptWindowMaxX
                    .ptMaxTrackSize.Y = lngScriptWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                ScriptWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                ScriptWindowProc = CallWindowProc(defScriptWindowProc, hWnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function
Public Function ArchiveWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hWnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmArchive.hWnd
            On Error Resume Next
            'our form-specific handler
            Select Case uMsg
                Case WM_GETMINMAXINFO
                Dim MMI As MINMAXINFO
                CopyMemory MMI, ByVal lParam, LenB(MMI)
                'set the MINMAXINFO data to the
                'minimum and maximum values set
                'by the option choice
                With MMI
                    .ptMinTrackSize.x = lngArchiveWindowMinX
                    .ptMinTrackSize.Y = lngArchiveWindowMinY
                    .ptMaxTrackSize.x = lngArchiveWindowMaxX
                    .ptMaxTrackSize.Y = lngArchiveWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                ArchiveWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                ArchiveWindowProc = CallWindowProc(defArchiveWindowProc, hWnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function
Public Function GetINIMessageFromFile(strSection_in As String, strKey_in As String, strDefault_in As String, strFile_in As String) As String

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
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, strFile_in)
    GetINIMessageFromFile = Left$(strReturned, lngRes)

End Function
