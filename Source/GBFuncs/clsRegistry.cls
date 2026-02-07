VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
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

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_CREATED_NEW_KEY = &H1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Function GetRootKey(ByVal strRoot As String) As Long
    
    'Returns the Constant of the Root Key to read/write from the string passed
    
    strRoot = UCase$(strRoot)
    Select Case strRoot
        Case "HKEY_CLASSES_ROOT"
            GetRootKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetRootKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetRootKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetRootKey = HKEY_USERS
        Case Else
            GetRootKey = 0 ' Fail
    End Select
    
End Function


Public Function GetSetting_HKCU_With_HKLM_Fallback(sRootKey As String, sKeyName As String, sValueName As String, vValue As Variant, _
                                bStringType As Boolean, Optional vDefault As Variant = "--DONT-SET--") As Boolean

    'safely get the value from HKCU falling backto HKLM

    GetSetting_HKCU_With_HKLM_Fallback = GetSetting("HKEY_CURRENT_USER", sKeyName, sValueName, vValue, bStringType, vDefault)

    If vValue = vDefault Then
        GetSetting_HKCU_With_HKLM_Fallback = GetSetting("HKEY_LOCAL_MACHINE", sKeyName, sValueName, vValue, bStringType, vDefault)
    End If

End Function


Public Function GetSetting_HKCU_Fallback_HKLM(sKeyName As String, sValueName As String, vValue As Variant, _
                                bStringType As Boolean, Optional vDefault As Variant = "--DONT-SET--") As Boolean

    'safely get the value from HKCU falling backto HKLM
    
    Dim lngDoneFallback As Long
    Dim vTempValue As Variant

'    lngDoneFallback = 1
'    Call GetSetting("HKEY_CURRENT_USER", "Software\GB64\GameBase\fallback\" & sKeyName, sValueName, lngDoneFallback, False, 0)
'
'    If lngDoneFallback = 0 Then
'        'not done fallback so get the setting from HKLM...
'        GetSetting_HKCU_Fallback_HKLM = GetSetting("HKEY_LOCAL_MACHINE", sKeyName, sValueName, vTempValue, bStringType, "--DONT-SET--")
'        If vTempValue <> "--NOT-FOUND--" Then
'            'write it to HKCU...
'            SaveSetting "HKEY_CURRENT_USER", sKeyName, sValueName, vTempValue, bStringType
'            vValue = vTempValue
'        End If
'        'and log that fallback is now done...
'        lngDoneFallback = 1
'        SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\fallback\" & sKeyName, sValueName, lngDoneFallback, False
'
'    Else
        'done fallback, so just get the value from HKCU
        GetSetting_HKCU_Fallback_HKLM = GetSetting("HKEY_CURRENT_USER", sKeyName, sValueName, vValue, bStringType, vDefault)
'    End If
   
End Function
Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    
    On Error GoTo QueryValueExError
       
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
        Select Case lType
        ' For strings
            Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = Left$(sValue, cch - 1)
                Else
                    vValue = Empty
                End If
                ' For DWORDS
            Case REG_DWORD:
                lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                If lrc = ERROR_NONE Then vValue = lValue
            Case Else
               'all other data types not supported
               lrc = -1
       End Select
QueryValueExExit:
       QueryValueEx = lrc
       Exit Function
QueryValueExError:
       Resume QueryValueExExit
End Function

Public Function SaveSetting(sRootKey As String, sKeyName As String, sValueName As String, vValueSetting As Variant, bStringType As Boolean) As Boolean
    Dim hKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    Dim lValueType As Long
    Dim RootKey As Long
    
    RootKey = GetRootKey(sRootKey)
    If RootKey = 0 Then
        SaveSetting = False
        Exit Function
    End If
    
    lRetVal = RegCreateKeyEx(RootKey, sKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
    If lRetVal <> ERROR_NONE Then
        RegCloseKey (hKey)
        SaveSetting = False
        Exit Function
    End If
    
    If bStringType = True Then
        lValueType = REG_SZ
    Else
        lValueType = REG_DWORD
    End If
    
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
    If lRetVal <> ERROR_NONE Then
        SaveSetting = False
    Else
        SaveSetting = True
    End If
End Function

Public Function GetSetting(sRootKey As String, sKeyName As String, sValueName As String, vValue As Variant, bStringType As Boolean, Optional vDefault As Variant = "--DONT-SET--") As Boolean
    
    Dim lRetVal As Long      'result of the API functions
    Dim Created As Long
    Dim hKey As Long         'handle of opened key
    Dim lValueType As Long
    Dim RootKey As Long
    
    RootKey = GetRootKey(sRootKey)
    If RootKey = 0 Then
        GetSetting = False
        Exit Function
    End If
    
    lRetVal = RegCreateKeyEx(RootKey, sKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, Created)
    If lRetVal <> ERROR_NONE Then
        GetSetting = False
        RegCloseKey (hKey)
        Exit Function
    End If
        
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    If lRetVal <> ERROR_NONE Then
        If vDefault = "--DONT-SET--" Then
            vValue = "--NOT-FOUND--"
            GetSetting = True
        Else
            If bStringType = True Then
                lValueType = REG_SZ
            Else
                lValueType = REG_DWORD
            End If
            lRetVal = SetValueEx(hKey, sValueName, lValueType, vDefault)
            lRetVal = QueryValueEx(hKey, sValueName, vValue)
            If lRetVal <> ERROR_NONE Then
                GetSetting = False
            Else
                GetSetting = True
            End If
        End If
    Else
        GetSetting = True
    End If
    RegCloseKey (hKey)

End Function

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function


