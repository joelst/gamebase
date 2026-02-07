Attribute VB_Name = "modAPI"
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

'Subclass APIs for restriction the form Size
Public defWindowProc As Long
Public lngWindowMinX As Long
Public lngWindowMinY As Long
Public lngWindowMaxX As Long
Public lngWindowMaxY As Long
Public defCatWindowProc As Long
Public lngCatWindowMinX As Long
Public lngCatWindowMinY As Long
Public lngCatWindowMaxX As Long
Public lngCatWindowMaxY As Long
Public defPictureSelectorWindowProc As Long
Public lngPictureSelectorWindowMinX As Long
Public lngPictureSelectorWindowMinY As Long
Public lngPictureSelectorWindowMaxX As Long
Public lngPictureSelectorWindowMaxY As Long
Public defGameSelectorWindowProc As Long
Public lngGameSelectorWindowMinX As Long
Public lngGameSelectorWindowMinY As Long
Public lngGameSelectorWindowMaxX As Long
Public lngGameSelectorWindowMaxY As Long
Public defExtrasWindowProc As Long
Public lngExtrasWindowMinX As Long
Public lngExtrasWindowMinY As Long
Public lngExtrasWindowMaxX As Long
Public lngExtrasWindowMaxY As Long
Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Browse For Folder API
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'for windows messages
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LVM_FIRST = &H1000
Public Const LVM_DELETEALLITEMS = LVM_FIRST + 9

'timing (duh!)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'for playing sounds
Public Const SND_ASYNC = &H1
Public Const SND_SYNC = &H0
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndPlaySoundMem Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long

'for getting drive types
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_FIXED = 3
Public Const DRIVE_CDROM = 5

'for shutting down sid players etc..
Public Const PROCESS_TERMINATE = &H1
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'for getting windows folder (for c64 fonts)
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_ID As Long = &H2
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MIIM_TYPE As Long = &H10
Private Const MIIM_DATA As Long = &H20
Private Const MFT_RADIOCHECK As Long = &H200
Private Const MENU_IDENTIFIER As Long = &H1
Private Const MFT_STRING As Long = &H0

'for changing menu checkmarks into radio buttons and adding bitmaps to menus
Public Const MF_BYPOSITION = &H400&
Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

'for displaying long strings with ellipses (e.g. in status bar)
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Public Const DT_CALCRECT = &H400
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000&
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_EXPANDTABS = &H40
Public Const DT_NOPREFIX = &H800
Public Const DT_WORDBREAK = &H10
Public Const DT_TABSTOP = &H80
Public Const DT_WORD_ELLIPSIS = &H40000

'ini file functions
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'for waiting for progs to end
Public Const INFINITE = &HFFFF
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'for copying a file
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Function GetINIMessage(strSection_in As String, strKey_in As String, strDefault_in As String) As String

    '------------------------------------------------------------------------
    'Reads a text item from the chosen GameBase Text.ini file
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
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, g_strMessagesFile)
    GetINIMessage = Left$(strReturned, lngRes)

End Function
Public Function EllipseText(ByVal strItem_in As String, sngDestWidthTwips_in As Long) As String

    Dim rc As RECT

    With Screen
        rc.Right = (sngDestWidthTwips_in / .TwipsPerPixelX) * 0.95
        rc.Bottom = 255 / .TwipsPerPixelY
    End With
    Call DrawText(frmMain.hDC, strItem_in, -1, rc, DT_CALCRECT Or DT_MODIFYSTRING Or DT_PATH_ELLIPSIS)
    EllipseText = strItem_in

End Function

Public Function EllipseTextEnd(ByVal strItem_in As String, sngDestWidthTwips_in As Long) As String

    Dim rc As RECT

    With Screen
        rc.Right = (sngDestWidthTwips_in / .TwipsPerPixelX) * 0.95
        rc.Bottom = 255 / .TwipsPerPixelY
    End With
    Call DrawText(frmMain.hDC, strItem_in, -1, rc, DT_CALCRECT Or DT_MODIFYSTRING Or DT_END_ELLIPSIS)
    EllipseTextEnd = strItem_in

End Function
Public Sub SetRadioMenuChecks(frmWithMenu As Form, lngMenuBarIndex As Long, ByVal lngMenuItem As Long)

    Dim lngMenuHandle As Long
    Dim lngSubMenuHandle As Long
    Dim udtMenuInfo As MENUITEMINFO

    'get the menu handles
    lngMenuHandle = GetMenu(frmWithMenu.hwnd)
    lngSubMenuHandle = GetSubMenu(lngMenuHandle, lngMenuBarIndex)

    'fill a structure to retrieve the current item menu string by first calling
    'GetMenuItemInfo passing a null string. The required size is returned in
    'udtMenuInfo.cch. Add 1 to accommodate the null that will be added when called.
    With udtMenuInfo
        .cbSize = Len(udtMenuInfo)
        .fMask = MIIM_TYPE
        .fType = MFT_STRING
        .dwTypeData = vbNullString
        .cch = Len(udtMenuInfo.dwTypeData)
        'get the needed buffer size
        Call GetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo)
        'set the buffer
        .dwTypeData = Space$(udtMenuInfo.cch + 1)
        .cch = Len(udtMenuInfo.dwTypeData)
    End With

    'get the data
    If GetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo) <> 0 Then
        'copy it's attributes, changing the checkmark to a radiobutton
        With udtMenuInfo
            .cbSize = Len(udtMenuInfo)
            .fType = MFT_RADIOCHECK
            .fMask = MIIM_TYPE
        End With
        'modify the menu item
        Call SetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo)
    End If

End Sub
Public Sub SetRadioMenuChecks2(lngSubMenuHandle As Long, ByVal lngMenuItem As Long)

    Dim udtMenuInfo As MENUITEMINFO

    'fill a structure to retrieve the current item menu string by first calling
    'GetMenuItemInfo passing a null string. The required size is returned in
    'udtMenuInfo.cch. Add 1 to accommodate the null that will be added when called.
    With udtMenuInfo
        .cbSize = Len(udtMenuInfo)
        .fMask = MIIM_TYPE
        .fType = MFT_STRING
        .dwTypeData = vbNullString
        .cch = Len(udtMenuInfo.dwTypeData)
        'get the needed buffer size
        Call GetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo)
        'set the buffer
        .dwTypeData = Space$(udtMenuInfo.cch + 1)
        .cch = Len(udtMenuInfo.dwTypeData)
    End With

    'get the data
    If GetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo) <> 0 Then
        'copy it's attributes, changing the checkmark to a radiobutton
        With udtMenuInfo
            .cbSize = Len(udtMenuInfo)
            .fType = MFT_RADIOCHECK
            .fMask = MIIM_TYPE
        End With
        'modify the menu item
        Call SetMenuItemInfo(lngSubMenuHandle, lngMenuItem, MENU_IDENTIFIER, udtMenuInfo)
    End If

End Sub

Public Function GetTypeOfDrive(strRootPath_in As String) As PathLocationEnum

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
Public Function FARPROC(pfn As Long) As Long
    
    FARPROC = pfn
    
End Function






Public Sub SubClass(hwnd As Long)
    
    'assign our own window message procedure (WindowProc)
  
    On Error Resume Next
    defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    g_blnSubclassed = True
    
End Sub
Public Sub CatSubClass(hwnd As Long)
    
    'assign our own window message procedure (CatWindowProc)
  
    On Error Resume Next
    defCatWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf CatWindowProc)
    g_blnCatSubclassed = True
    
End Sub

Public Sub PictureSelectorSubClass(hwnd As Long)
    
    'assign our own window message procedure (PicureSelectorWindowProc)
  
    On Error Resume Next
    defPictureSelectorWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf PictureSelectorWindowProc)
    g_blnPictureSelectorSubclassed = True
    
End Sub
Public Sub GameSelectorSubClass(hwnd As Long)
    
    'assign our own window message procedure (GameSelectorWindowProc)
  
    On Error Resume Next
    defGameSelectorWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf GameSelectorWindowProc)
    g_blnGameSelectorSubclassed = True
    
End Sub
Public Sub ExtrasSubClass(hwnd As Long)
    
    'assign our own window message procedure (ExtrasWindowProc)
  
    On Error Resume Next
    defExtrasWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf ExtrasWindowProc)
    g_blnExtrasSubclassed = True
    
End Sub
Public Sub UnSubClass(hwnd As Long)
    
    'restore the default message handling before exiting
    If defWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
        defWindowProc = 0
        g_blnSubclassed = False
    End If
    
End Sub
Public Sub CatUnSubClass(hwnd As Long)
    
    'restore the default message handling before exiting
    If defCatWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, defCatWindowProc
        defCatWindowProc = 0
        g_blnCatSubclassed = False
    End If
    
End Sub
Public Sub PictureSelectorUnSubClass(hwnd As Long)
    
    'restore the default message handling before exiting
    If defPictureSelectorWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, defPictureSelectorWindowProc
        defPictureSelectorWindowProc = 0
        g_blnPictureSelectorSubclassed = False
    End If
    
End Sub
Public Sub GameSelectorUnSubClass(hwnd As Long)
    
    'restore the default message handling before exiting
    If defGameSelectorWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, defGameSelectorWindowProc
        defGameSelectorWindowProc = 0
        g_blnGameSelectorSubclassed = False
    End If
    
End Sub
Public Sub ExtrasUnSubClass(hwnd As Long)
    
    'restore the default message handling before exiting
    If defExtrasWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, defExtrasWindowProc
        defExtrasWindowProc = 0
        g_blnExtrasSubclassed = False
    End If
    
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hwnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmMain.hwnd
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
                    .ptMinTrackSize.X = lngWindowMinX
                    .ptMinTrackSize.Y = lngWindowMinY
                    .ptMaxTrackSize.X = lngWindowMaxX
                    .ptMaxTrackSize.Y = lngWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                WindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                WindowProc = CallWindowProc(defWindowProc, hwnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function

Public Function CatWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hwnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmCategories.hwnd
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
                    .ptMinTrackSize.X = lngCatWindowMinX
                    .ptMinTrackSize.Y = lngCatWindowMinY
                    .ptMaxTrackSize.X = lngCatWindowMaxX
                    .ptMaxTrackSize.Y = lngCatWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                CatWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                CatWindowProc = CallWindowProc(defCatWindowProc, hwnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function
Public Function PictureSelectorWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hwnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmScrnShotSelector.hwnd
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
                    .ptMinTrackSize.X = lngPictureSelectorWindowMinX
                    .ptMinTrackSize.Y = lngPictureSelectorWindowMinY
                    .ptMaxTrackSize.X = lngPictureSelectorWindowMaxX
                    .ptMaxTrackSize.Y = lngPictureSelectorWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                PictureSelectorWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                PictureSelectorWindowProc = CallWindowProc(defPictureSelectorWindowProc, hwnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function

Public Function GameSelectorWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hwnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmGameSelector.hwnd
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
                    .ptMinTrackSize.X = lngGameSelectorWindowMinX
                    .ptMinTrackSize.Y = lngGameSelectorWindowMinY
                    .ptMaxTrackSize.X = lngGameSelectorWindowMaxX
                    .ptMaxTrackSize.Y = lngGameSelectorWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                GameSelectorWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                GameSelectorWindowProc = CallWindowProc(defGameSelectorWindowProc, hwnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function

Public Function ExtrasWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Callback window message procedure

    On Error Resume Next

    Select Case hwnd
        'If the handle returned is to our form,
        'perform form-specific message handling
        'to deal with the notifications. If it
        'is a general system message, pass it
        'on to the default window procedure.
        Case frmExtraProps.hwnd
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
                    .ptMinTrackSize.X = lngExtrasWindowMinX
                    .ptMinTrackSize.Y = lngExtrasWindowMinY
                    .ptMaxTrackSize.X = lngExtrasWindowMaxX
                    .ptMaxTrackSize.Y = lngExtrasWindowMaxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'the MSDN tells us that if we process
                'the message, to return 0
                ExtrasWindowProc = 0
            Case Else
                'this takes care of all the other messages
                'coming to the form and not specifically
                'handled above.
                ExtrasWindowProc = CallWindowProc(defExtrasWindowProc, hwnd, uMsg, wParam, lParam)
          End Select
    End Select
    
End Function
