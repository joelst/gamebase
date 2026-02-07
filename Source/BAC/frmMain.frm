VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Borg Assimilation Console"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iLstToolbar 
      Left            =   4440
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "diff"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0853
            Key             =   "master"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1429
      ButtonWidth     =   1058
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "iLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Master"
            Key             =   "master"
            ImageKey        =   "master"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Diff"
            Key             =   "diff"
            ImageKey        =   "diff"
         EndProperty
      EndProperty
   End
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
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
'       Borg Assimilation Console (BAC)
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


Private Function GetFile(ByVal Filename As String, ByVal Prompt As String, LastDirRegKKey As String) As String

    Dim Folder As String
    Dim InitDir As String
    Dim JustFile As String
    Dim JustPath As String

    On Error GoTo Exit_Here
    
    g_xobjRegistry.GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\LastDirs", LastDirRegKKey, InitDir, True, App.Path
    If g_xobjFunc.DirExists(InitDir) = False Then
        InitDir = App.Path
    End If
    
    GetFile = ""
    
    With CDialog
        .CancelError = True
        .DialogTitle = Prompt
        .Filename = Filename
        .Filter = Filename & "|" & Filename
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNNoReadOnlyReturn
        .InitDir = InitDir
        .ShowOpen
        GetFile = .Filename
    End With
    
    g_xobjFunc.GetFileAndPath GetFile, JustPath, JustFile
    g_xobjRegistry.SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\LastDirs", LastDirRegKKey, JustPath, True
    
Exit_Here:

End Function
Private Sub GetNewDiffDatabase()

    Dim strFile As String
    Dim sngVersion As Single

    'get new database file
    strFile = GetFile("*.mdb", "Select the DIFF Database", "Diff")
    If strFile <> "" Then
        'make sure database version is valid
        sngVersion = GetDatabaseVersion(strFile)
        If sngVersion = gc_sngREQ_DB_VERSION Then
            g_strDiff = strFile
            'close current database
            If g_blnDiffOpen = True Then
                g_dbDiff.Close
                Set g_dbDiff = Nothing
            End If
            'open new database
            Set g_dbDiff = OpenDatabase(g_strDiff)
            g_blnDiffOpen = True
            'read into the diff list
            Call frmDiffList.LoadList
        Else
            'wrong database version
            MsgBox "Incompatible database version (v" & Format$(sngVersion, "0.0") & "). Expected v" & Format$(gc_sngREQ_DB_VERSION, "0.0") & ".", vbInformation, "BAC"
        End If
    End If

End Sub

Private Sub GetNewMasterDatabase()

    Dim strFile As String
    Dim sngVersion As Single

    'get new database file
    strFile = GetFile("*.mdb", "Select the MASTER Database", "Master")
    If strFile <> "" Then
        'make sure database version is valid
        sngVersion = GetDatabaseVersion(strFile)
        If sngVersion = gc_sngREQ_DB_VERSION Then
            g_strMaster = strFile
            'close current database
            If g_blnMasterOpen = True Then
                g_dbMaster.Close
                Set g_dbMaster = Nothing
            End If
            'open new database
            Set g_dbMaster = OpenDatabase(g_strMaster)
            g_blnMasterOpen = True
            'read into the master list
            Call frmMasterList.LoadList
        Else
            'wrong database version
            MsgBox "Incompatible database version (v" & Format$(sngVersion, "0.0") & "). Expected v" & Format$(gc_sngREQ_DB_VERSION, "0.0") & ".", vbInformation, "BAC"
        End If
    End If

End Sub
Private Sub MDIForm_Load()

    Dim sngRegWidth As Single
    Dim sngRegHeight As Single
    Dim sngRegTop As Single
    Dim sngRegLeft As Single
    Dim blnMaximized As Boolean

    'create external objects
    Set g_xobjFunc = New GBFuncs.clsFunctions
    Set g_xobjRegistry = New GBFuncs.clsRegistry

    'set window size/position
    With g_xobjRegistry
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainWidth", sngRegWidth, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainHeight", sngRegHeight, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainTop", sngRegTop, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainLeft", sngRegLeft, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainMaximized", blnMaximized, False, False
    End With
    If sngRegWidth < 500 Then
        sngRegWidth = 4770
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight < 500 Then
        sngRegHeight = 4230
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (Height / 2)
        sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    Move sngRegLeft, sngRegTop, sngRegWidth, sngRegHeight
    If blnMaximized Then
        WindowState = vbMaximized
    End If

    'init variables
    g_strMaster = ""
    g_strDiff = ""
    g_blnMasterOpen = False
    g_blnDiffOpen = False

    'unload the splash screen
    Do Until g_xobjStartupClock.TimeTaken > 1.75
        DoEvents
    Loop
    Unload frmSplash
    Set frmSplash = Nothing
    Show
    Visible = True

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

    'save window size/position
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainWidth", Width, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainHeight", Height, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainTop", Top, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainLeft", Left, False
        End If
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "MainMaximized", (WindowState = vbMaximized), False
    End With

    'close databases
    If g_blnMasterOpen = True Then
        g_dbMaster.Close
        Set g_dbMaster = Nothing
    End If
    If g_blnDiffOpen = True Then
        g_dbDiff.Close
        Set g_dbDiff = Nothing
    End If

    'release external objects
    Set g_xobjFunc = Nothing
    Set g_xobjRegistry = Nothing

End Sub


Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "master"
            Call GetNewMasterDatabase
        Case "diff"
            Call GetNewDiffDatabase
    End Select

End Sub

Private Function GetDatabaseVersion(strDBFile_in As String) As Single

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
    GetDatabaseVersion = CSng(lngDBMajorVersion & "." & lngDBMinorVersion)
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

