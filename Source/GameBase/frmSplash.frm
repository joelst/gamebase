VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgSplash 
      Height          =   1125
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmSplash"
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

Dim FCall As cFuncCall

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

'for preventing duplicate database instances
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const GW_HWNDPREV = 3
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'apis for blitting the desktop onto the form to create the illusion of a transparent form
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub BlitDesktopOntoForm()

    '------------------------------------------------------------
    'grabs the portion of the desktop that is behind the form
    'and blits it to the form's background to create the illusion
    'of a transparent form when displaying a splash screen
    'that contains transparent pixels
    '------------------------------------------------------------

    Dim lnghDesktopWindow As Long   'handle to desktop window
    Dim lnghDesktopDC As Long       'handle to desktop device context
    Dim lngWidth As Long            'width of area to blit
    Dim lngHeight As Long           'height of area to blit
    Dim lngSourceX As Long          'X start of blit source (desktop)
    Dim lngSourceY As Long          'Y start of blit source (desktop)
    
    'get the desktop's window handle
    lnghDesktopWindow = GetDesktopWindow()
    
    'get the desktop's device context handle
    lnghDesktopDC = GetDC(lnghDesktopWindow)
    
    'get x/y/width/height values for blit
    With Screen
        lngSourceX = ((.Width / 2) - (Width / 2)) / .TwipsPerPixelX
        lngSourceY = ((.Height / 2) - (Height / 2)) / .TwipsPerPixelY
        lngWidth = Width / .TwipsPerPixelX
        lngHeight = Height / .TwipsPerPixelY
    End With

    'blit from desktop to form
    Call BitBlt(hDC, 0, 0, lngWidth, lngHeight, lnghDesktopDC, lngSourceX, lngSourceY, SRCCOPY)
    
    'release the desktop's device context
    Call ReleaseDC(lnghDesktopWindow, lnghDesktopDC)

End Sub


Private Function LoadRandomSplashScreen() As Boolean

    '-------------------------------------------------------------------
    'checks if there are any splash screens in the \splash folder
    'and loads a random one to replace the default one (if it finds any)
    '
    'Remarks:   The random select will always include the default
    '           splash screen
    '-------------------------------------------------------------------

    Dim a_strImages() As String             'array for splash screen filenames
    Dim lngImageCount As Long               'number of images to test
    Dim a_strTypes(5) As String             'array for image types
    Dim lngTypeIdx As Long                  'loop index
    Dim lngSelectedImage As Long            'index of selected splash screen to use
    Dim strNextFile As String               'filename storage for each found splash screen image
    Dim lngLowerBound As Long               'for random select
    Dim lngUpperBound As Long               'for random select
    Dim strImageFile As String

    ''''''''''''''''''''''''''''
    'read in the image filenames
    ''''''''''''''''''''''''''''

    'set the image types allowed (VB ImageBox control must support all types specified here)
    a_strTypes(0) = ".gif"
    a_strTypes(1) = ".jpg"
    a_strTypes(2) = ".bmp"
    a_strTypes(3) = ".dib"
    a_strTypes(4) = ".wmf"
    a_strTypes(5) = ".emf"

    'init variables
    ReDim a_strImages(0)
    lngImageCount = 0
    a_strImages(0) = ""

    'read in the filenames
    For lngTypeIdx = 0 To UBound(a_strTypes)
        strNextFile = Dir(g_strGameBasePathWBS & "Splash\*" & a_strTypes(lngTypeIdx))
        Do While strNextFile <> ""
            ReDim Preserve a_strImages(lngImageCount)
            a_strImages(lngImageCount) = strNextFile
            lngImageCount = lngImageCount + 1
            strNextFile = Dir
        Loop
    Next lngTypeIdx

    '''''''''''''''''''''''''''''
    'set the random splash screen
    '''''''''''''''''''''''''''''

    'only select a random image if there is at least one in the \splash folder
    If a_strImages(0) <> "" Then
        lngImageCount = UBound(a_strImages)
        If lngImageCount > 0 Then
            lngLowerBound = 0
            lngUpperBound = lngImageCount
            Randomize
            lngSelectedImage = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)
            strImageFile = g_strGameBasePathWBS & "Splash\" & a_strImages(lngSelectedImage)
            imgSplash.Picture = LoadPicture(strImageFile)
        Else
            strImageFile = g_strGameBasePathWBS & "Splash\" & a_strImages(0)
            imgSplash.Picture = LoadPicture(g_strGameBasePathWBS & "Splash\" & a_strImages(0))
        End If
        g_lngStartupSplashSecondsToWait = GetSecondsToWaitFromFilename(strImageFile)
        If g_lngStartupSplashSecondsToWait = 0 Then g_lngStartupSplashSecondsToWait = 2 'default
        LoadRandomSplashScreen = True
    Else
        LoadRandomSplashScreen = False
    End If

End Function
Private Sub PlayRandomSound()

    '-------------------------------------------------------------------
    'checks if there are any wavs in the \splash folder
    'and loads a random one to play (if it finds any)
    '-------------------------------------------------------------------

    Dim a_strWavs() As String               'array for wav filenames
    Dim lngWavCount As Long                 'number of wavs to test
    Dim lngSelectedWav As Long              'index of selected wav to use
    Dim strNextFile As String               'filename storage for each found wav
    Dim lngLowerBound As Long               'for random select
    Dim lngUpperBound As Long               'for random select
    Dim strWavPathFile As String

    ''''''''''''''''''''''''''''
    'read in the sound filenames
    ''''''''''''''''''''''''''''

    'init variables
    ReDim a_strWavs(0)
    lngWavCount = 0
    a_strWavs(0) = ""

    'read in the filenames
    strNextFile = Dir(g_strGameBasePathWBS & "Splash\*.wav")
    Do While strNextFile <> ""
        ReDim Preserve a_strWavs(lngWavCount)
        a_strWavs(lngWavCount) = strNextFile
        lngWavCount = lngWavCount + 1
        strNextFile = Dir
    Loop

    ''''''''''''''''''''''
    'play the random sound
    ''''''''''''''''''''''

    'only select a random image if there is at least one in the \splash folder
    If a_strWavs(0) <> "" Then
        lngWavCount = UBound(a_strWavs)
        lngLowerBound = 0
        lngUpperBound = lngWavCount
        Randomize
        lngSelectedWav = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)
        g_blnIsWavOnStart = True
        strWavPathFile = g_strGameBasePathWBS & "Splash\" & a_strWavs(lngSelectedWav)
        g_lngStartupWavSecondsToWait = GetSecondsToWaitFromFilename(strWavPathFile)
        Call sndPlaySound(strWavPathFile, SND_NODEFAULT Or SND_ASYNC)
        If g_lngStartupWavSecondsToWait > 0 Then
            Set g_xobjStartupWavClock = New clsClock
            Call g_xobjStartupWavClock.StartClock
        End If
    End If

End Sub

Private Function GetSecondsToWaitFromFilename(ByVal strPathFile As String) As Long

    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim lngDotPos As Long

    GetSecondsToWaitFromFilename = 0

    strTemp = Left$(strPathFile, Len(strPathFile) - 4)
    lngDotPos = InStrRev(strTemp, ".")
    If lngDotPos >= Len(strTemp) Then Exit Function
    GetSecondsToWaitFromFilename = Val(Right$(strTemp, Len(strTemp) - lngDotPos))

    Exit Function

Err_Handler:
    GetSecondsToWaitFromFilename = 0

End Function

' Returns 0 on success, error code on error
Private Function RegisterGenericFile(ByVal DoRegister As Boolean, ByVal FileToRegister As String) As Long
    Dim Result As Long
    Dim OpenFileStructure As OFSTRUCT
    Const OF_EXIST = &H4000
    Const FILE_NOT_FOUND = 2
   
    Set FCall = New cFuncCall
    FCall.LibraryName = FileToRegister

    Result = OpenFile(FileToRegister, OpenFileStructure, OF_EXIST)
    
    If OpenFileStructure.nErrCode = FILE_NOT_FOUND Then
        Result = -1
    Else
        If DoRegister Then
            FCall.FunctionName = "DllRegisterServer"
        Else
            FCall.FunctionName = "DllUnregisterServer"
        End If
        FCall.CallFunction
        Result = 0
    End If
    
    RegisterGenericFile = Result

    Set FCall = Nothing
End Function

Private Sub Form_Load()

    Dim blnOtherSplash As Boolean
    Dim blnLoadStartupForm As Boolean
    Dim lngOldHWnd As Long
    Dim lngResult As Long
    Dim lngSplash As Long
    Dim strCommand As String
    Dim blnSkipVCheck As Boolean

    g_blnIsWavOnStart = False
    g_lngStartupWavSecondsToWait = 0
    g_lngStartupSplashSecondsToWait = 2
    Set g_xobjStartupClock = New GBFuncs.clsClock

    'set the GameBase application path (with added backslash)
    g_strAppPathWithBackslash = App.Path
    If Right$(g_strAppPathWithBackslash, 1) <> "\" Then g_strAppPathWithBackslash = g_strAppPathWithBackslash & "\"

    If g_blnInVBIDE = True Then
        '==============================================
        '
        '           FOR DEBUG TESTING ONLY
        '
        g_strAppPathWithBackslash = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
        '==============================================
    Else
        'auto register dlls
        lngResult = RegisterGenericFile(True, "gbfuncs.dll")
        If lngResult <> 0 Then
            End
        End If
        lngResult = RegisterGenericFile(True, "gbsql.dll")
        If lngResult <> 0 Then
            End
        End If
        lngResult = RegisterGenericFile(True, "gemus.dll")
        If lngResult <> 0 Then
            End
        End If
    End If

    Set g_xobjRegistry = New GBFuncs.clsRegistry
    Set g_xobjFunc = New GBFuncs.clsFunctions
    'Set g_xobjPNG = New cPNG  'removed

    g_blnSkipVersionCheck = False
    If Command$ = "" Then
        blnLoadStartupForm = True
    Else
        'want to skip the version check if swapping from one GB to another
        strCommand = Trim$(Command$)
        If Left$(strCommand, 3) = "-v " Then
            g_blnSkipVersionCheck = True
            strCommand = Right$(strCommand, Len(strCommand) - 3)
        End If
        If g_xobjFunc.DirExists(g_strAppPathWithBackslash & strCommand) = True Then
            App.Title = "wporiwoperi"
            lngOldHWnd = FindWindow("ThunderRT6Main", "GB_" & strCommand)
            If lngOldHWnd > 0 Then
                lngOldHWnd = GetWindow(lngOldHWnd, GW_HWNDPREV)
                lngResult = OpenIcon(lngOldHWnd)
                lngResult = SetForegroundWindow(lngOldHWnd)
                Set g_xobjRegistry = Nothing
                Set g_xobjFunc = Nothing
                'Set g_xobjPNG = Nothing  'removed
                End
            Else
                If g_xobjFunc.FileExists(g_strAppPathWithBackslash & strCommand & "\" & strCommand & ".mdb", False, "") = True Then
                    g_strGameBasePath = g_strAppPathWithBackslash & strCommand
                    g_strGameBasePathWBS = g_strAppPathWithBackslash & strCommand & "\"
                    g_strGameBaseName = strCommand
                    g_strMessagesFile = g_strAppPathWithBackslash & strCommand & "\Text.ini"
                    g_strGameBaseDBPath = g_strAppPathWithBackslash & strCommand & "\" & strCommand & ".mdb"
                    g_strEmuFile = g_strAppPathWithBackslash & strCommand & "\Emulators.ini"
                    g_strMusicFile = g_strAppPathWithBackslash & strCommand & "\MusicPlayers.ini"
                    g_strPathFile = g_strAppPathWithBackslash & strCommand & "\Paths.ini"
                    g_strSettingsFile = g_strAppPathWithBackslash & strCommand & "\Settings.ini"
                    g_strGameBaseDotINI = g_strAppPathWithBackslash & "GameBase.ini"
                    blnLoadStartupForm = False
                    App.Title = "GB_" & strCommand
                Else
                    blnLoadStartupForm = True
                End If
            End If
        Else
            blnLoadStartupForm = True
        End If
    End If

    If blnLoadStartupForm = True Then
        'no duplicate selector window instances
        App.Title = "plftioeutioef"
        lngOldHWnd = FindWindow("ThunderRT6Main", "GameBase_")
        If lngOldHWnd > 0 Then
            lngOldHWnd = GetWindow(lngOldHWnd, GW_HWNDPREV)
            lngResult = OpenIcon(lngOldHWnd)
            lngResult = SetForegroundWindow(lngOldHWnd)
            Set g_xobjRegistry = Nothing
            Set g_xobjFunc = Nothing
            'Set g_xobjPNG = Nothing  'removed
            End
        End If
        App.Title = "GameBase_"
        Load frmStartup
        If frmStartup.m_blnUnload = True Then
            MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_NoGameBaseFolders1", "There are no GameBase folders to choose from!") & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Error Messages", "Error_NoGameBaseFolders2", "Please use the GameBase Toolbox to create a new GameBase."), vbExclamation, gc_strAPPNAME
            If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "GBToolbox.exe", False, "") = True Then
                Call Shell(g_strAppPathWithBackslash & "GBToolbox.exe -NOMSGBOX", vbNormalFocus)
            End If
            Set g_xobjRegistry = Nothing
            Set g_xobjFunc = Nothing
            'Set g_xobjPNG = Nothing  'removed
            End
        ElseIf g_strGameBaseName = "" Then
            Dim bValid As Long
            
            'frmStartup.Show vbModal, Me  'modified to fake modal
            frmStartup.Show vbModeless, Me
            bValid = True
            Do While bValid = True
                If (frmStartup.Visible = False) Then
                    bValid = False
                End If
                DoEvents
                If (g_strGameBaseName <> "") Then
                    bValid = False
                End If
            Loop
        Else
            Unload frmStartup
        End If
    End If

    If (g_strGameBaseName = "") Or (g_blnAbort = True) Then
        Set g_xobjRegistry = Nothing
        Set g_xobjFunc = Nothing
        'Set g_xobjPNG = Nothing  'removed
        End
    End If
    If g_xobjFunc.FileExists(g_strGameBaseDBPath, False, gc_strAPPNAME) = False Then
        MsgBox "Could not find the database file:-" & vbCrLf & vbCrLf & g_strGameBaseDBPath
        Set g_xobjRegistry = Nothing
        Set g_xobjFunc = Nothing
        'Set g_xobjPNG = Nothing  'removed
        End
    End If

    'no duplicate database instances
    App.Title = "dofiudiofusiof"
    lngOldHWnd = FindWindow("ThunderRT6Main", "GB_" & g_strGameBaseName)
    If lngOldHWnd > 0 Then
        lngOldHWnd = GetWindow(lngOldHWnd, GW_HWNDPREV)
        lngResult = OpenIcon(lngOldHWnd)
        lngResult = SetForegroundWindow(lngOldHWnd)
        Set g_xobjRegistry = Nothing
        Set g_xobjFunc = Nothing
        'Set g_xobjPNG = Nothing  'removed
        End
    End If
    App.Title = "GB_" & g_strGameBaseName

    'load a random splash screen from the \splash folder if any exist
    blnOtherSplash = LoadRandomSplashScreen()

    'set the caption and version info
    g_strAPPVERSION = Format$(Val(App.Major), "0") & "." & Format$(Val(App.Minor), "0")
    g_strWINDOWCAPTION = g_strGameBaseName
    g_strDatabaseName = g_strGameBaseName

    'set the form's width and height to that of the splash screen
    'NOTE: this implicitly sets the top and left properties of the form
    'as the form's StartUpPosition is set to "2 - CenterScreen"
    With imgSplash
        DoEvents
        Width = (Width - ScaleWidth) + .Width
        Height = (Height - ScaleHeight) + .Height
    End With

    'show the splash screen
    Call BlitDesktopOntoForm
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "ShowSplash", lngSplash, False, 1
    If lngSplash > 0 Then
        Call PlayRandomSound
        Show
        g_xobjStartupClock.StartClock
    End If

    'allow to show before loading frmMain
    DoEvents
    'load the main form
    'NOTE: this form is unloaded at the end of frmMain_Load()
    Load frmMain

    Set g_xobjStartupClock = Nothing
    
End Sub

Private Sub imgSplash_Click()
    g_lngStartupWavSecondsToWait = 0
    g_lngStartupSplashSecondsToWait = 0
End Sub
