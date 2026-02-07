VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GameBase Toolbox"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4620
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox FileBox 
      Height          =   480
      Left            =   4560
      Pattern         =   "*.ini"
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.StatusBar sbGameBase 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3285
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Right-Click each Node for options"
            TextSave        =   "Right-Click each Node for options"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstGameBase 
      Left            =   4560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "gamebase"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0823
            Key             =   "textfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BBA
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F60
            Key             =   "inifile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1307
            Key             =   "access"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvGameBase 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilstGameBase"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New GameBase..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileNewRomCenter2 
         Caption         =   "Import &RomCenter 2.00 Datafile..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRoot 
      Caption         =   "Root"
      Visible         =   0   'False
      Begin VB.Menu mnuRootNew 
         Caption         =   "&New GameBase..."
      End
      Begin VB.Menu mnuRootNewRomCenter2 
         Caption         =   "Import &RomCenter 2.00 Datafile..."
      End
   End
   Begin VB.Menu mnuGameBase 
      Caption         =   "GBRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuGameBaseDelete 
         Caption         =   "&Delete..."
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "DB"
      Visible         =   0   'False
      Begin VB.Menu mnuDBCompact 
         Caption         =   "&Compact"
      End
      Begin VB.Menu mnuDBFix 
         Caption         =   "&Scan and Fix Errors"
      End
      Begin VB.Menu mnuDBBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBImportRC2 
         Caption         =   "Import &RomCenter 2.00 Datafile..."
      End
      Begin VB.Menu mnuDBBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBGenres 
         Caption         =   "&Genres Wizard..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDBBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBFirstLoad 
         Caption         =   "&Database Settings..."
      End
      Begin VB.Menu mnuDBBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBClearExtras 
         Caption         =   "Clear all Game &Extras..."
      End
      Begin VB.Menu mnuDBRemoveViews 
         Caption         =   "&Remove All Custom Views..."
      End
      Begin VB.Menu mnuDBDeleteRecords 
         Caption         =   "Delete Rec&ords from Table..."
      End
   End
   Begin VB.Menu mnuScriptRoot 
      Caption         =   "ScriptRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuScriptNew 
         Caption         =   "&New Script..."
         Enabled         =   0   'False
      End
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

Private m_blnExpert As Boolean


Private Function GetFile(ByVal InitDir As String, Filter As String, ByVal Prompt As String) As String

    Dim xobjFunc As GBFuncs.clsFunctions
    Dim Folder As String

    On Error GoTo Exit_Here
    
    GetFile = ""

    Set xobjFunc = New GBFuncs.clsFunctions
    If xobjFunc.DirExists(InitDir) = False Then
        Folder = App.Path
    Else
        Folder = InitDir
    End If
    Set xobjFunc = Nothing
    
    With cDialog
        .CancelError = True
        .DialogTitle = Prompt
        .Filename = ""
        .Filter = Filter
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNNoReadOnlyReturn
        .InitDir = Folder
        .ShowOpen
        GetFile = .Filename
    End With
    
Exit_Here:

End Function
Private Sub CompactDatabase(strFilename_in As String, strGBFolderWBS_in As String)

    Dim Func As GBFuncs.clsFunctions
    Dim dbEng As DAO.DBEngine
    Dim strFilename As String
    Dim strTempFilename As String
    Dim f As Integer
    Dim lngOldSize As Long
    Dim lngNewSize As Long
    Dim dbGameBase As DAO.Database
    Dim rsGameBase As DAO.Recordset
    Dim lngOldVal As Long
    Dim lngNewVal As Long

    If Not (tvGameBase.SelectedItem Is Nothing) Then
        strFilename = strFilename_in
        strTempFilename = strGBFolderWBS_in & "gbtemp.mdb"
        If FileExists(strFilename) Then
            Screen.MousePointer = vbHourglass
            DoEvents
            On Error GoTo Err_Handler
            'get a test variable
            lngOldVal = -3
            lngNewVal = -45
            Set dbGameBase = OpenDatabase(strFilename)
            Set rsGameBase = dbGameBase.OpenRecordset("SELECT * FROM Games", dbOpenSnapshot)
            If rsGameBase.EOF = True Then
                lngOldVal = 0
            Else
                rsGameBase.MoveLast
                rsGameBase.MoveFirst
                lngOldVal = rsGameBase.RecordCount
            End If
            rsGameBase.Close
            Set rsGameBase = Nothing
            Set dbGameBase = Nothing
            'get old file size
            Set Func = New GBFuncs.clsFunctions
            f = FreeFile()
            Open strFilename For Binary As #f
            lngOldSize = LOF(f)
            Close #f
            'delete old backup if present
            Func.DeletePathSilent strTempFilename
            'compact
            Set dbEng = New DBEngine
            DBEngine.CompactDatabase strFilename, strTempFilename
            Set DBEngine = Nothing
            'test the new db is ok
            Set dbGameBase = OpenDatabase(strTempFilename)
            Set rsGameBase = dbGameBase.OpenRecordset("SELECT * FROM Games", dbOpenSnapshot)
            If rsGameBase.EOF = True Then
                lngNewVal = 0
            Else
                rsGameBase.MoveLast
                rsGameBase.MoveFirst
                lngNewVal = rsGameBase.RecordCount
            End If
            rsGameBase.Close
            Set rsGameBase = Nothing
            Set dbGameBase = Nothing
            'ok to continue?
            If lngOldVal = lngNewVal Then
                'rename the temp one
                Func.DeletePathSilent strFilename
                Name strTempFilename As strFilename
                f = FreeFile()
                Open strFilename For Binary As #f
                lngNewSize = LOF(f)
                Close #f
                Screen.MousePointer = vbDefault
                MsgBox "Compacted Database from " & Format$(lngOldSize / 1024, "0") & "Kb to " & Format$(lngNewSize / 1024, "0") & "Kb.", vbInformation, "GameBase"
            Else
                MsgBox "Compact Failed" & vbCrLf & vbCrLf & "This is one of those crap and annoying error messages that tells you absolutely nothing about WHY the operation failed, because the author probably doesn't have a clue why it failed!", vbInformation, "GameBase"
            End If
        Else
            MsgBox "Could not find file: " & strFilename, vbInformation, "GameBase"
        End If
    End If
    
    Exit Sub

Err_Handler:
    MsgBox "Unable to compact database '" & strFilename & "'." & vbCrLf & vbCrLf & Err.Description & " (" & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")", vbCritical, "GB-Toolbox"

End Sub



Private Sub CreateNewDatabase()

    ' Resource IDs for resource file: -
    '101 = Text.ini
    '102 = Gfx\gfx.txt
    '103 = help.htm
    '104 = credits.ini

    On Error GoTo Err_Handler

    Dim Func As GBFuncs.clsFunctions
    Dim strGBName As String
    Dim strFolder As String
    Dim intFileNum As Integer
    Dim strTextFile As String
    Dim Noddy As Node

    With g_udtInputBox
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 64
        .strCancelButtonText = "&Cancel"
        .strCaption = "Create New GameBase"
        .strInvalidChars = "\/:*?""<>|"
        .strOKButtonText = "&OK"
        .strPrompt = "Enter the new GameBase name:"
        .strReturnedText = ""
        .strStartingText = ""
        frmInputBox.Show vbModal, Me
        If .blnOKClicked = True Then
            strGBName = .strReturnedText
            strFolder = g_strAppPathWithBackSlash & .strReturnedText
            Set Func = New GBFuncs.clsFunctions
            If Func.DirExists(strFolder) = False Then
                If Func.FileExists(g_strAppPathWithBackSlash & "Empty.mdb", False, "GB-Toolbox") = True Then
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    Call MkDir(strFolder)
                    Call MkDir(strFolder & "\Gfx")
                    Call MkDir(strFolder & "\Scripts")
                    Call MkDir(strFolder & "\Sounds")
                    Call MkDir(strFolder & "\Splash")
                    Call FileCopy(g_strAppPathWithBackSlash & "Empty.mdb", strFolder & "\" & strGBName & ".mdb")
                    Call CreateFile(strFolder & "\Credits.txt", "Created by GBToolbox v" & Trim$(App.Major) & "." & Trim$(App.Minor) & " on " & Format$(Now, "dddd dd mmmm yyyy") & " by Unknown!" & vbCrLf & vbCrLf & "If you're the creator of this GameBase database, put your credits in this file (credits.txt).")
                    strTextFile = GetResourceFileAsString(101)
                    Call CreateFile(strFolder & "\Text.ini", strTextFile)
                    strTextFile = GetResourceFileAsString(102)
                    Call CreateFile(strFolder & "\Gfx\gfx.txt", strTextFile)
                    strTextFile = GetResourceFileAsString(103)
                    Call CreateFile(strFolder & "\help.htm", strTextFile)
                    strTextFile = GetResourceFileAsString(104)
                    Call CreateFile(strFolder & "\Credits.ini", strTextFile)
                    Call CreateFile(strFolder & "\Sounds\sounds.txt", "Place your own Classic Game WAV files here.  See the GameBase documentation for more details.")
                    Call CreateFile(strFolder & "\Splash\Splash.txt", "Place your own splash screen(s) here.  See the GameBase documentation for more details.")
                    Call RefreshList
                    Set Noddy = tvGameBase.Nodes(strFolder & "\")
                    If Not (Noddy Is Nothing) Then
                        Set tvGameBase.SelectedItem = Noddy
                        tvGameBase.SelectedItem.Selected = True
                        tvGameBase.SelectedItem.Expanded = True
                        tvGameBase.SelectedItem.EnsureVisible
                    End If
                Else
                    MsgBox "Could not find Empty.mdb in the GameBase installation folder." & vbCrLf & vbCrLf & "You can download another Empty.mdb file from the GameBase web site.", vbInformation, "GB-Toolbox"
                End If
            Else
                MsgBox "There is already a GameBase folder with that name!", vbInformation, "GB-Toolbox"
            End If
            Set Func = Nothing
        End If
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

Err_Handler:
    MsgBox "Unable to create GameBase '" & strGBName & "'." & vbCrLf & vbCrLf & Err.Description & " (" & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")", vbCritical, "GB-Toolbox"
    Call RefreshList
    Screen.MousePointer = vbDefault

End Sub
Private Function FileExists(strPathAndFile_in As String) As Boolean

    '-------------------------------------------------------
    'Determines whether a file exists or not
    '
    'Inputs:    strPathAndFile:     The Path\Filename of the
    '                               file to check.
    '
    'Returns:   True if File Exists
    '           False otherwise
    '-------------------------------------------------------
    
    'set error handler just incase
    On Error GoTo NoFile:
    
    'check if file exists
    FileExists = (Dir(strPathAndFile_in) <> "")
    
    'release error handler and exit
    On Error GoTo 0
    Exit Function
    
NoFile:
    'release error handler and exit
    On Error GoTo 0
    FileExists = False

End Function





Private Sub ImportRomCenter2()

    Dim xobjFunc As GBFuncs.clsFunctions
    Dim xobjReg As GBFuncs.clsRegistry
    Dim strInitDir As String
    Dim strPath As String
    Dim strFile As String
    Dim strDataFile As String
    Dim strVersion As String
    Dim Noddy As Node

    Set xobjFunc = New GBFuncs.clsFunctions
    Set xobjReg = New GBFuncs.clsRegistry

    'get the datafile
    xobjReg.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GBToolbox\BrowsePaths", "RomCenter2DataFilePath", strInitDir, True, App.Path
    strDataFile = GetFile(strInitDir, "RomCenter 2.0 Datafiles (*.dat)|*.dat", "Select the RC-2.0 Datafile...")
    If strDataFile <> "" Then
        'save path for next time
        Call xobjFunc.GetFileAndPath(strDataFile, strPath, strFile)
        xobjReg.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GBToolbox\BrowsePaths", "RomCenter2DataFilePath", strPath, True
        'check the version
        strVersion = GetINIMessage(strDataFile, "DAT", "version", "")
        
        If strVersion = "2" Or strVersion = "2." Or strVersion = "2.0" Or strVersion = "2.00" Then
            'get credits
            g_strRC2Author = GetINIMessage(strDataFile, "CREDITS", "Author", "")
            g_strRC2Version = GetINIMessage(strDataFile, "CREDITS", "Version", "")
            g_strRC2Comment = GetINIMessage(strDataFile, "CREDITS", "Comment", "")
            'show the import form
            g_strRC2DataFile = strDataFile
            frmRomCenter2.Show vbModal, Me
            If g_blnRC2Imported = True Then
                If g_blnRC2ExistingDB = False Then
                    'update the list and show the new database
                    Call RefreshList
                    Set Noddy = tvGameBase.Nodes(g_strRC2GBFolderWBS)
                    If Not (Noddy Is Nothing) Then
                        Set tvGameBase.SelectedItem = Noddy
                        tvGameBase.SelectedItem.Selected = True
                        tvGameBase.SelectedItem.Expanded = True
                        tvGameBase.SelectedItem.EnsureVisible
                    End If
                End If
                MsgBox "You will need to run ""Scan and Fix Errors"" on this database file before you can use it."
            End If
        Else
            MsgBox "This datafile was not found to be version 2.00!" & vbCrLf & vbCrLf & "Reported version is: [" & strVersion & "]", vbInformation, "GBToolbox"
        End If
    End If

    Set Noddy = Nothing
    Set xobjFunc = Nothing
    Set xobjReg = Nothing

End Sub

Private Sub RefreshList()

    Dim Func As GBFuncs.clsFunctions
    Dim lngIdx As Long
    Dim LItem As Node
    Dim strDir As String
    Dim lngAttr As Long
    Dim lngFileIdx As Long

    Screen.MousePointer = vbHourglass
    DoEvents

    Set Func = New GBFuncs.clsFunctions

    With tvGameBase
        .Nodes.Clear
        .Nodes.Add , , "ROOT", "Gamebases", "gamebase", "gamebase"
        .Nodes("ROOT").Bold = True
    End With
    With ilstGameBase
        For lngIdx = .ListImages.Count To 6 Step -1
            .ListImages.Remove lngIdx
        Next lngIdx
    End With

    'load the icons
    lngIdx = 6
    strDir = Dir(g_strAppPathWithBackSlash, vbDirectory)
    Do Until strDir = ""
        lngAttr = GetAttr(g_strAppPathWithBackSlash & strDir)
        If lngAttr And vbDirectory Then
            If (strDir <> ".") And (strDir <> "..") And (UCase$(strDir) <> "DOCS") Then
                If Func.FileExists(g_strAppPathWithBackSlash & strDir & "\" & strDir & ".mdb", False, "") = True Then
                    Set LItem = tvGameBase.Nodes.Add("ROOT", tvwChild, g_strAppPathWithBackSlash & strDir & "\", strDir)
                    If Func.FileExists(g_strAppPathWithBackSlash & strDir & "\Gfx\Startup_32x32.gif", False, "") = False Then
                        LItem.Image = 1
                    Else
                        ilstGameBase.ListImages.Add lngIdx, , LoadPicture(g_strAppPathWithBackSlash & strDir & "\Gfx\Startup_32x32.gif")
                        LItem.Image = lngIdx
                        lngIdx = lngIdx + 1
                    End If
                    Set LItem = Nothing
                    'database
                    tvGameBase.Nodes.Add g_strAppPathWithBackSlash & strDir & "\", tvwChild, g_strAppPathWithBackSlash & strDir & "\" & strDir & ".mdb", strDir & ".mdb", "access", "access"
                    'inifile folder
                    tvGameBase.Nodes.Add g_strAppPathWithBackSlash & strDir & "\", tvwChild, strDir & " INI Files", "INI Files", "folder", "folder"
                    'script folder
                    tvGameBase.Nodes.Add g_strAppPathWithBackSlash & strDir & "\", tvwChild, strDir & " GEMUS Scripts", "GEMUS Scripts", "folder", "folder"
                    'textfile folder
                    tvGameBase.Nodes.Add g_strAppPathWithBackSlash & strDir & "\", tvwChild, strDir & " Text Files", "Text Files", "folder", "folder"
                    'ini files
                    FileBox.Pattern = "*.ini"
                    FileBox.Path = g_strAppPathWithBackSlash & strDir
                    FileBox.Refresh
                    For lngFileIdx = 0 To FileBox.ListCount - 1
                        tvGameBase.Nodes.Add strDir & " INI Files", tvwChild, g_strAppPathWithBackSlash & strDir & "\" & FileBox.List(lngFileIdx), FileBox.List(lngFileIdx), "inifile", "inifile"
                    Next lngFileIdx
                    'text files
                    FileBox.Pattern = "*.txt"
                    FileBox.Path = g_strAppPathWithBackSlash & strDir
                    FileBox.Refresh
                    For lngFileIdx = 0 To FileBox.ListCount - 1
                        tvGameBase.Nodes.Add strDir & " Text Files", tvwChild, g_strAppPathWithBackSlash & strDir & "\" & FileBox.List(lngFileIdx), FileBox.List(lngFileIdx), "textfile", "textfile"
                    Next lngFileIdx
                    'script files
                    If Func.DirExists(g_strAppPathWithBackSlash & strDir & "\Scripts") = False Then
                        Call MkDir(g_strAppPathWithBackSlash & strDir & "\Scripts")
                    End If
                    FileBox.Pattern = "*.txt"
                    FileBox.Path = g_strAppPathWithBackSlash & strDir & "\Scripts"
                    FileBox.Refresh
                    For lngFileIdx = 0 To FileBox.ListCount - 1
                        tvGameBase.Nodes.Add strDir & " GEMUS Scripts", tvwChild, g_strAppPathWithBackSlash & strDir & "\Scripts\" & FileBox.List(lngFileIdx), FileBox.List(lngFileIdx), "textfile", "textfile"
                    Next lngFileIdx
                End If
            End If
        End If
        strDir = Dir()
    Loop

    tvGameBase.Nodes("ROOT").Expanded = True
    sbGameBase.Panels(1).Text = "Right-Click each Node for options"
    
    Set Func = Nothing

    Screen.MousePointer = vbDefault

End Sub

Private Sub ReplaceNullData()

    Dim dbGameBase As DAO.Database
    Dim strFilename As String
    Dim sngDatabaseVersion As Single
    
    If Not (tvGameBase.SelectedItem Is Nothing) Then
        strFilename = tvGameBase.SelectedItem.Key
        If FileExists(strFilename) Then
            sngDatabaseVersion = GetDatabaseVersion(strFilename)
            If sngDatabaseVersion >= 2 Then
                Screen.MousePointer = vbHourglass
                DoEvents
                Set dbGameBase = OpenDatabase(strFilename)
                With dbGameBase
                    .Execute "UPDATE Games SET Name = """" WHERE Name IS NULL"
                    .Execute "UPDATE Games SET Filename = """" WHERE Filename IS NULL"
                    .Execute "UPDATE Games SET ScrnshotFilename = """" WHERE ScrnshotFilename IS NULL"
                    .Execute "UPDATE Games SET SidFilename = """" WHERE SidFilename IS NULL"
                    .Execute "UPDATE Games SET DateLastPlayed = """" WHERE DateLastPlayed IS NULL"
                    .Execute "UPDATE Games SET HighScore = """" WHERE HighScore IS NULL"
                    .Execute "UPDATE Games SET MemoText = """" WHERE MemoText IS NULL"
                    .Execute "UPDATE Games SET Comment = """" WHERE Comment IS NULL"
                    .Execute "UPDATE Games SET V_Comment = """" WHERE V_Comment IS NULL"
                    .Execute "UPDATE Games SET FileToRun = """" WHERE FileToRun IS NULL"
                    .Execute "UPDATE Music SET Name = """" WHERE Name IS NULL"
                    .Execute "UPDATE Music SET Filename = """" WHERE Filename IS NULL"
                    .Execute "UPDATE Musicians SET Grp = """" WHERE Grp IS NULL"
                    .Execute "UPDATE Musicians SET Nick = """" WHERE Nick IS NULL"
                    .Execute "UPDATE Musicians SET Photo = """" WHERE Photo IS NULL"
                    .Execute "UPDATE Crackers SET Cracker = ""Rename Me!"" WHERE Cracker IS NULL"
                    .Execute "UPDATE Difficulty SET Difficulty = ""Rename Me!"" WHERE Difficulty IS NULL"
                    .Execute "UPDATE Genres SET Genre = ""Sub-Genre"" WHERE Genre IS NULL"
                    .Execute "UPDATE Languages SET Language = ""Rename Me!"" WHERE Language IS NULL"
                    .Execute "UPDATE Musicians SET Musician = ""Rename Me!"" WHERE Musician IS NULL"
                    .Execute "UPDATE PGenres SET ParentGenre = ""Parent-Genre"" WHERE ParentGenre IS NULL"
                    .Execute "UPDATE Programmers SET Programmer = ""Rename Me!"" WHERE Programmer IS NULL"
                    .Execute "UPDATE Publishers SET Publisher = ""Rename Me!"" WHERE Publisher IS NULL"
                    If sngDatabaseVersion > 2# Then
                        .Execute "UPDATE Games SET CRC = """" WHERE CRC IS NULL"
                        .Execute "UPDATE Extras SET Name = """" WHERE Name IS NULL"
                        .Execute "UPDATE Extras SET Path = """" WHERE Path IS NULL"
                        .Execute "UPDATE Extras SET Data = """" WHERE Data IS NULL"
                        .Execute "UPDATE Games SET Filesize = 0 WHERE Filesize IS NULL"
                    End If
                    If sngDatabaseVersion > 2.2 Then
                        .Execute "UPDATE Games SET Gemus = """" WHERE Gemus IS NULL"
                    End If
                    If sngDatabaseVersion > 2.4 Then
                        .Execute "UPDATE Config SET FirstLoadMessage = """" WHERE FirstLoadMessage IS NULL"
                        .Execute "UPDATE Games SET V_LengthType = 19 WHERE V_LengthType < 0"
                        .Execute "UPDATE Games SET V_LengthType = 19 WHERE V_LengthType > 19"
                    End If
                    If sngDatabaseVersion > 2.5 Then
                        .Execute "UPDATE Games SET WebLink_Name = """" WHERE WebLink_Name IS NULL"
                        .Execute "UPDATE Games SET WebLink_URL = """" WHERE WebLink_URL IS NULL"
                        .Execute "UPDATE Games SET V_WebLink_Name = """" WHERE V_WebLink_Name IS NULL"
                        .Execute "UPDATE Games SET V_WebLink_URL = """" WHERE V_WebLink_URL IS NULL"
                        .Execute "UPDATE Artists SET Artist = ""Rename Me!"" WHERE Artist IS NULL"
                        .Execute "UPDATE Developers SET Developer = ""Rename Me!"" WHERE Developer IS NULL"
                        .Execute "UPDATE Licenses SET License = ""Rename Me!"" WHERE License IS NULL"
                        .Execute "UPDATE Rarities SET Rarity = ""Rename Me!"" WHERE Rarity IS NULL"
                    End If
                    If sngDatabaseVersion > 2.7 Then
                        .Execute "UPDATE Extras SET FileToRun = """" WHERE FileToRun IS NULL"
                    End If
                End With
                Set dbGameBase = Nothing
                Screen.MousePointer = vbDefault
                MsgBox "Fixed: All NULL Data Removed!", vbInformation, "GameBase"
            Else
                MsgBox "Version of database is less than 2.0!", vbInformation, "GameBase"
            End If
        Else
            MsgBox "Could not find file: " & strFilename, vbInformation, "GameBase"
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call RefreshList
    End If
End Sub

Private Sub Form_Load()

    Dim strCommandLine As String
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim Reg As GBFuncs.clsRegistry

    'get current size/position/state variables
    Set Reg = New GBFuncs.clsRegistry
    With Reg
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GBToolbox\Window", "MainWidth", sngRegWidth, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GBToolbox\Window", "MainHeight", sngRegHeight, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GBToolbox\Window", "MainTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GBToolbox\Window", "MainLeft", sngRegLeft, False, -1
    End With
    Set Reg = Nothing

    'make sure not to big or small...
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
    '...or out of full screen view
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (Height / 2)
        sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    'set the position and size
    Top = sngRegTop
    Left = sngRegLeft
    Width = sngRegWidth
    Height = sngRegHeight

    strCommandLine = UCase$(Command$)
    m_blnExpert = (InStr(strCommandLine, "-EXPERT") > 0)
    
    If (InStr(strCommandLine, "-NOMSGBOX") = 0) And m_blnExpert = False Then
        MsgBox "Please make sure the GameBase frontend is NOT running when using this tool.", vbInformation, "GameBase"
    End If
    

    g_strAppPathWithBackSlash = App.Path
    If Right$(g_strAppPathWithBackSlash, 1) <> "\" Then
        g_strAppPathWithBackSlash = g_strAppPathWithBackSlash & "\"
    End If
    
    If g_blnInVBIDE = True Then
        g_strAppPathWithBackSlash = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    
    Caption = "GameBase Toolbox v" & Format$(App.Major, "0") & "." & Format$(App.Minor, "0")
    Call RefreshList

End Sub

Private Sub Form_Resize()

    If ScaleHeight >= sbGameBase.Height Then
        sbGameBase.Panels(1).Width = ScaleWidth
        tvGameBase.Move 0, 0, ScaleWidth, ScaleHeight - sbGameBase.Height
    End If

End Sub




Private Sub Form_Unload(Cancel As Integer)

    Dim Reg As GBFuncs.clsRegistry

    'save current window size/position
    Set Reg = New GBFuncs.clsRegistry
    With Reg
        'main window size/position
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GBToolbox\Window", "MainWidth", Width, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GBToolbox\Window", "MainHeight", Height, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GBToolbox\Window", "MainTop", Top, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GBToolbox\Window", "MainLeft", Left, False
    End With
    Set Reg = Nothing

End Sub


Private Sub mnuDBClearExtras_Click()

    Dim Func As GBFuncs.clsFunctions
    Dim Reg As GBFuncs.clsRegistry
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strTemp As String
    Dim sngDatabaseVersion As Single

    sngDatabaseVersion = GetDatabaseVersion(tvGameBase.SelectedItem.Key)
    If sngDatabaseVersion < 2.1 Then
        MsgBox "Database must be v2.1 or higher to perform this operation!", vbInformation, "GBToolbox"
        Exit Sub
    End If

    Set Func = New GBFuncs.clsFunctions
    Set Reg = New GBFuncs.clsRegistry
    If Func.FileExists(tvGameBase.SelectedItem.Key, False, "GBToolbox") = True Then
        If MsgBox("Do you really want to clear all the Game Extras linked this database?" & vbCrLf & vbCrLf & "(This will not remove the files themselves.)", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, "Confirm Clearing All Game Extras") = vbYes Then
            Set db = OpenDatabase(tvGameBase.SelectedItem.Key)
            strTemp = sbGameBase.Panels(1).Text
            sbGameBase.Panels(1).Text = "Clearing Game Extras in '" & tvGameBase.SelectedItem.Text & "'..."
            Load frmWait
            frmWait.lblPrompt.Caption = "Clearing Game Extras..."
            frmWait.Show vbModeless, Me
            tvGameBase.Enabled = False
            db.Execute "UPDATE Games SET Extras = False WHERE Extras = True"
            db.Execute "DELETE * FROM Extras WHERE (1)"
            Unload frmWait
            tvGameBase.Enabled = True
            sbGameBase.Panels(1).Text = strTemp
            db.Close
            Set db = Nothing
            MsgBox "Cleared all Game Extras!", vbInformation, "GBToolbox"
        End If
    Else
        MsgBox "Huh? This database doesn't exist!"
        Call RefreshList
    End If
    Set Func = Nothing
    Set Reg = Nothing

End Sub

Private Sub mnuDBCompact_Click()

    Dim strTemp As String
    
    strTemp = sbGameBase.Panels(1).Text
    sbGameBase.Panels(1).Text = "Compacting " & strTemp & "..."
    Load frmWait
    frmWait.lblPrompt.Caption = "Compacting..."
    frmWait.Show vbModeless, Me
    tvGameBase.Enabled = False
    Call CompactDatabase(tvGameBase.SelectedItem.Key, tvGameBase.SelectedItem.Parent.Key)
    Unload frmWait
    tvGameBase.Enabled = True
    sbGameBase.Panels(1).Text = strTemp

End Sub

Private Sub mnuDBDeleteRecords_Click()

    g_strRemoveDB = tvGameBase.SelectedItem.Key
    frmRemoveRec.Show vbModal, Me
    
End Sub

Private Sub mnuDBFirstLoad_Click()

    Dim Func As GBFuncs.clsFunctions

    Set Func = New GBFuncs.clsFunctions
    If Func.FileExists(tvGameBase.SelectedItem.Key, False, "GBToolbox") = True Then
        g_strDatabase = tvGameBase.SelectedItem.Key
        frmFirstLoad.Show vbModal, Me
    Else
        MsgBox "Huh? This database doesn't exist!"
        Call RefreshList
    End If
    Set Func = Nothing

End Sub


Private Sub mnuDBFix_Click()

    Dim strTemp As String
    
    strTemp = sbGameBase.Panels(1).Text
    sbGameBase.Panels(1).Text = "Removing NULL Data..."
    Load frmWait
    frmWait.lblPrompt.Caption = "Removing NULL Data..."
    frmWait.Show vbModeless, Me
    tvGameBase.Enabled = False
    Call ReplaceNullData
    Unload frmWait
    tvGameBase.Enabled = True
    sbGameBase.Panels(1).Text = strTemp

End Sub

Private Sub mnuDBGenres_Click()

    Dim Func As GBFuncs.clsFunctions

    Set Func = New GBFuncs.clsFunctions
    If Func.FileExists(tvGameBase.SelectedItem.Key, False, "GBToolbox") = True Then
        g_strDatabase = tvGameBase.SelectedItem.Key
        frmGenres.Show vbModal, Me
    Else
        MsgBox "Huh? This database doesn't exist!"
        Call RefreshList
    End If
    Set Func = Nothing

End Sub

Private Sub mnuDBImportRC2_Click()
    g_strRC2DBFile = tvGameBase.SelectedItem.Key
    g_strRC2DBName = Left$(tvGameBase.SelectedItem.Text, Len(tvGameBase.SelectedItem.Text) - 4)
    g_blnRC2ExistingDB = True
    Call ImportRomCenter2
End Sub

Private Sub mnuDBRemoveViews_Click()

    Dim Func As GBFuncs.clsFunctions
    Dim Reg As GBFuncs.clsRegistry
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strTemp As String

    Set Func = New GBFuncs.clsFunctions
    Set Reg = New GBFuncs.clsRegistry
    If Func.FileExists(tvGameBase.SelectedItem.Key, False, "GBToolbox") = True Then
        If MsgBox("Do you really want to delete all the custom views from this database?", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, "Confirm Removing All Custom Views") = vbYes Then
            strTemp = sbGameBase.Panels(1).Text
            sbGameBase.Panels(1).Text = "Removing Custom Views in '" & tvGameBase.SelectedItem.Text & "'..."
            Load frmWait
            frmWait.lblPrompt.Caption = "Removing Custom Views..."
            frmWait.Show vbModeless, Me
            tvGameBase.Enabled = False
            Set db = OpenDatabase(tvGameBase.SelectedItem.Key)
            Set rs = db.OpenRecordset("SELECT * FROM ViewData", dbOpenDynaset)
            With rs
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        .Delete
                        .MoveNext
                    Loop
                End If
                .Close
            End With
            Set rs = Nothing
            Set rs = db.OpenRecordset("SELECT * FROM ViewFilters", dbOpenDynaset)
            With rs
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        .Delete
                        .MoveNext
                    Loop
                End If
                .Close
            End With
            Set rs = Nothing
            db.Close
            Set db = Nothing
            Reg.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & tvGameBase.SelectedItem.Parent.Text & "\QuickFilter", "NumFilters", 0, False
            Unload frmWait
            tvGameBase.Enabled = True
            sbGameBase.Panels(1).Text = strTemp
            MsgBox "Removed All Custom Views!", vbInformation, "GBToolbox"
        End If
    Else
        MsgBox "Huh? This database doesn't exist!"
        Call RefreshList
    End If
    Set Func = Nothing
    Set Reg = Nothing

End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    Call CreateNewDatabase
End Sub

Private Sub mnuFileNewRomCenter2_Click()
    g_blnRC2ExistingDB = False
    Call ImportRomCenter2
End Sub

Private Sub mnuGameBaseDelete_Click()

    Dim Func As GBFuncs.clsFunctions
    Dim strFolder As String

    strFolder = Left$(tvGameBase.SelectedItem.Key, Len(tvGameBase.SelectedItem.Key) - 1)
    Set Func = New GBFuncs.clsFunctions
    If Func.DirExists(strFolder) = True Then
        If MsgBox("Are you sure you want to permanently delete this GameBase folder?" & vbCrLf & vbCrLf & "'" & strFolder & "'" & vbCrLf & vbCrLf & "WARNING: This action is irreversable.  Once it's gone, it's gone!", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, "Really Delete '" & tvGameBase.SelectedItem.Text & "'?") = vbYes Then
            sbGameBase.Panels(1).Text = "Deleting '" & strFolder & "'..."
            Load frmWait
            frmWait.lblPrompt.Caption = "Deleting..."
            frmWait.Show vbModeless, Me
            tvGameBase.Enabled = False
            Call Func.DeletePathSilent(strFolder)
            Unload frmWait
            tvGameBase.Enabled = True
            Call RefreshList
        End If
    Else
        MsgBox "Huh? This folder doesn't exist!"
        Call RefreshList
    End If
    
    Set Func = Nothing

End Sub


Private Sub mnuRootNew_Click()

    Call CreateNewDatabase

End Sub

Private Sub mnuRootNewRomCenter2_Click()
    Call ImportRomCenter2
End Sub

Private Sub tvGameBase_DblClick()

    Dim Func As GBFuncs.clsFunctions
    Dim Nod As Node

    Set Nod = tvGameBase.SelectedItem
    Set Func = New GBFuncs.clsFunctions
    
    If Not (Nod Is Nothing) Then
    
        If Nod.Key = "ROOT" Then
        
            'Root - DO NOTHING

        ElseIf Nod.Parent.Key = "ROOT" Then
        
            'GameBase Root - DO NOTHING

        ElseIf Right$(Nod.Key, 9) = " INI Files" Then
        
            'INI Files Folder - DO NOTHING
                
        ElseIf Right$(Nod.Key, 10) = " Text Files" Then
        
            'Text Files Folder - DO NOTHING
                
        ElseIf LCase$(Right$(Nod.Key, 4)) = ".mdb" Then
        
            'Database File

            If Func.FileExists(Nod.Key, False, "GameBase") = True Then
                If m_blnExpert = False Then
                    MsgBox "WARNING: Only EXPERTS should change the GameBase database in MS Access!" & vbCrLf & vbCrLf & "If you change something by hand then your database doesn't work, you've only yourself to blame!", vbInformation, "GBToolbox"
                End If
                Call Func.ShellExecuteFile(Nod.Key, "", False)
            End If
                
        ElseIf Right$(Nod.Parent.Key, 10) = " INI Files" Then
            
            'INI File
            
            If Func.FileExists(Nod.Key, False, "GameBase") = True Then
                Call Func.ShellExecuteFile(Nod.Key, "", False)
            End If
            
        ElseIf Right$(Nod.Parent.Key, 11) = " Text Files" Then

            'Text File
            
            If Func.FileExists(Nod.Key, False, "GameBase") = True Then
                Call Func.ShellExecuteFile(Nod.Key, "", False)
            End If
            
        ElseIf Right$(Nod.Parent.Key, 14) = " GEMUS Scripts" Then

            'Script File
            
            If Func.FileExists(Nod.Key, False, "GameBase") = True Then
                Call Func.ShellExecuteFile(Nod.Key, "", False)
            End If

        End If
    End If
    Set Nod = Nothing
    Set Func = Nothing

End Sub

Private Sub tvGameBase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Nod As Node

    If Button = vbRightButton Then
        Set Nod = tvGameBase.HitTest(X, Y)
        If Not (Nod Is Nothing) Then
        
            If Nod.Key = "ROOT" Then
            
                'Root
                
                PopupMenu mnuRoot
        
            ElseIf Nod.Parent.Key = "ROOT" Then
            
                'GameBase Root
                
                PopupMenu mnuGameBase
            
            ElseIf Right$(Nod.Key, 10) = " INI Files" Then
            
                'INI Files Folder - DO NOTHING
                    
            ElseIf Right$(Nod.Key, 11) = " Text Files" Then

                    
            ElseIf Right$(Nod.Key, 14) = " GEMUS Scripts" Then
            
                'Script Files Folder
                
                PopupMenu mnuScriptRoot
                    
            ElseIf LCase$(Right$(Nod.Key, 4)) = ".mdb" Then
            
                'Database File

                PopupMenu mnuDB
                    
            ElseIf Right$(Nod.Parent.Key, 10) = " INI Files" Then
                
                'INI File
                
                
                
            ElseIf Right$(Nod.Parent.Key, 11) = " Text Files" Then


            ElseIf Right$(Nod.Parent.Key, 14) = " GEMUS Scripts" Then

                
            End If
        End If
        Set Nod = Nothing
    End If

End Sub

Private Sub tvGameBase_NodeClick(ByVal Node As MSComctlLib.Node)

    Select Case Node.Key
        Case "ROOT"
            sbGameBase.Panels(1).Text = "Right-Click each Node for options"
        Case Else
            sbGameBase.Panels(1).Text = Node.Key
    End Select
    

End Sub


