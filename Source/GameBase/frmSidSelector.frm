VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSidSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sid Selector"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmSidSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAll 
      Caption         =   "&Apply to all games/sids in current view"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin TabDlg.SSTab tabSearch 
      Height          =   3615
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Manual Search"
      TabPicture(0)   =   "frmSidSelector.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FileBox"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DirBox"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Auto Search"
      TabPicture(1)   =   "frmSidSelector.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstTypes"
      Tab(1).Control(1)=   "cmdFind"
      Tab(1).Control(2)=   "txtFind"
      Tab(1).Control(3)=   "lstFound"
      Tab(1).Control(4)=   "lblSearching"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox lstTypes 
         Height          =   315
         Left            =   -71400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70320
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   3
         Top             =   600
         Width           =   3345
      End
      Begin VB.ListBox lstFound 
         Height          =   2205
         ItemData        =   "frmSidSelector.frx":0902
         Left            =   -74760
         List            =   "frmSidSelector.frx":0904
         TabIndex        =   6
         Top             =   1080
         Width           =   5055
      End
      Begin VB.DirListBox DirBox 
         Height          =   1215
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   5055
      End
      Begin VB.FileListBox FileBox 
         Height          =   1455
         Hidden          =   -1  'True
         Left            =   240
         Pattern         =   "*.sid"
         System          =   -1  'True
         TabIndex        =   1
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label lblSearching 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Searching..."
         Height          =   195
         Left            =   -74760
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Sid Info"
      Height          =   1140
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   5535
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   795
         UseMnemonic     =   0   'False
         Width           =   4305
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   540
         UseMnemonic     =   0   'False
         Width           =   4305
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   285
         UseMnemonic     =   0   'False
         Width           =   4305
      End
      Begin VB.Label lblPCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   795
         UseMnemonic     =   0   'False
         Width           =   705
      End
      Begin VB.Label lblPAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   540
         UseMnemonic     =   0   'False
         Width           =   510
      End
      Begin VB.Label lblPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   280
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame fraNone 
      Caption         =   "Sid Info"
      Height          =   1140
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   5535
      Begin VB.Label lblError 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "None selected or file is not a valid PSid"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   540
         UseMnemonic     =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Label lblSid 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      UseMnemonic     =   0   'False
      Width           =   5505
   End
End
Attribute VB_Name = "frmSidSelector"
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

Private m_strLastPath As String 'last valid path selected
Private WithEvents m_xobjFunc As GBFuncs.clsFunctions
Attribute m_xobjFunc.VB_VarHelpID = -1

Private Sub SetSidInFileBox(strFile_in As String)

    '-------------------------------------------
    'finds and selects a sid file in the filebox
    'that matches the passed in sid file
    '
    'Inputs:    strFile_in:     Filename to find
    '-------------------------------------------

    Dim strUpper As String  'upper case filename
    Dim lngFileIdx As Long  'loop index

    'find the file in the filebox
    strUpper = UCase(strFile_in)
    With FileBox
        For lngFileIdx = 0 To (.ListCount - 1)
            If strUpper = UCase(.List(lngFileIdx)) Then
                'found! - select it
                .ListIndex = lngFileIdx
                Exit For
            End If
        Next lngFileIdx
    End With

End Sub

Private Function WithinRootSidPath(strPath_in As String) As Boolean

    '------------------------------------------------
    'tests whether the passed in string is within the
    'root sid music path
    '
    'Inputs:    strPath_in:     String to test.
    '
    'Returns:   True if the path is within the root
    '           False otherwise
    '------------------------------------------------

    WithinRootSidPath = (UCase(Left$(strPath_in, Len(g_udtRegPathSid.strRealPath))) = UCase(g_udtRegPathSid.strRealPath))

End Function

Private Function FilenameForDatabase(strFullPath_in As String) As String

    '------------------------------------------------
    'returns the sid filename for the database from
    'the passed in full path\filename
    '
    'Inputs:    strFullPath_in:     String to test.
    '
    'Returns:   Path\filename for the database
    '
    'Assumptions:   The passed in string must include
    '               the root sid path
    '------------------------------------------------

    FilenameForDatabase = Right$(strFullPath_in, Len(strFullPath_in) - Len(g_udtRegPathSid.strWithBackSlash))

End Function
Private Sub cmdCancel_Click()

    'do not update g_strSidSelectorFilename
    Unload Me

End Sub

Private Sub cmdClear_Click()

    'no sid selected now...
    FileBox.ListIndex = gc_lngNOLISTINDEX
    lblSid = g_strNoneSelected
    fraInfo.Visible = False
    fraNone.Visible = True
    cmdClear.Enabled = False

End Sub

Private Sub cmdFind_Click()

    '---------------------------------------------------
    'perform a search on the text in the txtFind textbox
    '---------------------------------------------------

    Dim blnOldClearState As Boolean     'state of Clear button
    Dim strSearchPath As String         'path for recursive search
    Dim lngSearchIdx As Long            'loop index

    'disable the controls while searching
    DirBox.Enabled = False
    FileBox.Enabled = False
    txtFind.Enabled = False
    cmdFind.Enabled = False
    blnOldClearState = cmdClear.Enabled
    cmdClear.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    lstFound.Clear
    lstFound.Visible = False
    lstFound.Enabled = False
    lblSearching.Visible = True
    MousePointer = vbHourglass
    DoEvents

    'do the search
    strSearchPath = g_udtRegPathSid.strRealPath 'passing byref so copy first as argument changes in find function
    m_xobjFunc.FindFiles strSearchPath, "*" & txtFind & "*." & lstTypes.List(lstTypes.ListIndex)

    're-enable the controls
    DirBox.Enabled = True
    FileBox.Enabled = True
    txtFind.Enabled = True
    cmdFind.Enabled = True
    cmdClear.Enabled = blnOldClearState
    cmdOK.Enabled = True
    cmdOK.Default = True
    cmdCancel.Enabled = True
    lblSearching.Visible = False
    lstFound.Visible = True
    lstFound.Enabled = True
    If lstFound.ListCount = 0 Then
        txtFind.SetFocus
    Else
        lstFound.SetFocus
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub lstFound_Click()

    Dim strPathFile As String
    Dim strPath As String
    Dim strFile As String

    If lstFound.ListIndex = gc_lngNOLISTINDEX Then
        Exit Sub
    Else
        'select the sid
         strPathFile = g_udtRegPathSid.strWithBackSlash & lstFound.List(lstFound.ListIndex)
         g_xobjFunc.GetFileAndPath strPathFile, strPath, strFile
         DirBox.Path = strPath
         DoEvents 'allow filebox to fill
         Call SetSidInFileBox(strFile)
         cmdOK.Default = True
    End If

End Sub

Private Sub lstFound_DblClick()

    'just play the selected sid!
    FileBox_DblClick

End Sub

Private Sub tabSearch_Click(PreviousTab As Integer)

    If tabSearch.Tab <> PreviousTab Then
        If tabSearch.Tab = 0 Then
            If FileBox.ListCount > 0 Then
                FileBox.SetFocus
            Else
                DirBox.SetFocus
            End If
            cmdOK.Default = True
        Else
            txtFind.SetFocus
        End If
    End If

End Sub

Private Sub txtFind_Change()

    If txtFind.Text <> "" Then
        cmdFind.Enabled = True
        cmdFind.Default = True
    Else
        cmdOK.Enabled = False
    End If


End Sub


Private Sub txtFind_GotFocus()

    With txtFind
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    cmdFind.Default = True

End Sub
Private Sub cmdOK_Click()

    'confirmation
    If chkAll.Value = vbChecked Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "AreYouSure", "Are you sure you want to update ALL GAMES AND MUSIC in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If

    'return the new chosen sid
    If lblSid = g_strNoneSelected Then
        g_strSidSelectorFilename = g_strNone
        g_strSidSelectorSidName = ""
    Else
        g_strSidSelectorFilename = lblSid
        g_strSidSelectorSidName = lblName
    End If
    g_blnApplySidFileToAll = (chkAll.Value = vbChecked)
    Unload Me

End Sub

Private Sub DirBox_Change()

On Error GoTo Err_Handler

    'test if new selected path is within the root sid path
    If WithinRootSidPath(DirBox.Path) Then
        m_strLastPath = DirBox.Path
        FileBox.Path = m_strLastPath
    Else
        DirBox.Path = m_strLastPath
    End If
    Exit Sub

Err_Handler:
    MsgBox Err.Description & " (" & Trim$(Err.Number) & ")", vbExclamation, gc_strAPPNAME
    On Error GoTo 0
    Unload Me

End Sub

Private Sub DirBox_Click()
    
    DirBox.Path = DirBox.List(DirBox.ListIndex)
    
End Sub


Private Sub FileBox_Click()

    Dim strPathWBS As String        'path with backslash
    Dim strSidName As String        'the sid's internal name
    Dim strSidAuthor As String      'the sid's internal author
    Dim strSidCopyright As String   'the sid's internal copyright
    Dim strSidFile As String        'the full sid path\filename
    Dim lngTotalSongs As Long       'the sid's internal no. songs
    Dim lngDefaultSong As Long      'the sid's internal default song no.

    With FileBox
        If .ListIndex > gc_lngNOLISTINDEX Then
            'set the database filename label
            strPathWBS = FileBox.Path
            If Right$(strPathWBS, 1) <> "\" Then
                strPathWBS = strPathWBS & "\"
            End If
            strSidFile = strPathWBS & .List(.ListIndex)
            lblSid = FilenameForDatabase(strSidFile)
            cmdClear.Enabled = True
            'load the sid info
            If g_xobjFunc.GetSidInfo(strSidFile, strSidName, strSidAuthor, strSidCopyright, lngTotalSongs, lngDefaultSong) Then
                lblName = strSidName
                lblAuthor = strSidAuthor
                lblCopyright = strSidCopyright
                fraInfo.Visible = True
                fraNone.Visible = False
            Else
                fraInfo.Visible = False
                fraNone.Visible = True
            End If
        Else
            'nothing in the filebox
            lblSid = g_strNoneSelected
            fraInfo.Visible = False
            fraNone.Visible = True
            cmdClear.Enabled = False
        End If
    End With
    
End Sub

Private Sub FileBox_DblClick()

    Dim strPathWBS As String        'path with backslash
    Dim strSidName As String        'the sid's internal name
    Dim strSidAuthor As String      'the sid's internal author
    Dim strSidCopyright As String   'the sid's internal copyright
    Dim strSidFile As String        'the full sid path\filename
    Dim lngTotalSongs As Long       'the sid's internal no. songs
    Dim lngDefaultSong As Long      'the sid's internal default song no.
    Dim lngPlayerIdx As Long        'loop index
    Dim dblProcessID As Double      'process ID of shelled player
    Dim lngHandle As Long

    '------------------------------------------------------------------------
    'play the selected sid with the first valid player in the list of players
    '------------------------------------------------------------------------

    With FileBox
        If .ListIndex > gc_lngNOLISTINDEX Then
            'set the database filename label
            strPathWBS = FileBox.Path
            If Right$(strPathWBS, 1) <> "\" Then
                strPathWBS = strPathWBS & "\"
            End If
            strSidFile = strPathWBS & .List(.ListIndex)
            'play the sid
            If g_lngTotalSidPlayers > 0 Then
                'find first valid player
                lngPlayerIdx = GetAppropriatePlayer(strSidFile)
                If lngPlayerIdx > -1 Then
                    With ga_udtSidPlayers(lngPlayerIdx)
                        If g_xobjFunc.FileExists(.strFullPathAndFilename, False, "") Then
                            'now shell the player
                            If Not .blnUseShortFilenames Then
                                dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, strSidFile, False), vbNormalFocus)
                            Else
                                dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, strSidFile, True), vbNormalFocus)
                            End If
                            'store the process handle
                            lngHandle = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                            Call AddToSidPlayerHandles(lngHandle)
                        End If
                    End With
                Else
                    Call MsgBox(g_xobjFunc.GetINIMessage("Error_Messages", "Error_NoMusicPlayersForType", "There were no Music Players found that support this Music File.  You can set the supported file types in Options!"), vbInformation, gc_strAPPNAME)
                End If
            End If
        End If
    End With

End Sub


Private Sub Form_Load()

    Dim strLastBrowsePath As String             'from the registry
    Dim strSidFile As String                    'the full path\filename of the current sid file
    Dim strPath As String                       'just the sid's path
    Dim strFile As String                       'just the sid's filename
    Dim lngFTIdx As Long                        'loop index
    Dim strPattern As String                    'filebox pattern
    Dim a_strTypes() As String                  'for getting all the supported music filetypes
    Const DesignScaleHeight As Single = 6375    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5775     'The required scalewidth for the form
    Dim lngTotalPatterns As Long

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'load the finder object
    Set m_xobjFunc = New GBFuncs.clsFunctions

    'set the filters for the filebox
    lngTotalPatterns = GetSingleMusicTypesArray(a_strTypes)
    strPattern = GetFileBoxPatternStringFromArray(a_strTypes)
    FileBox.Pattern = strPattern

    'load them into the types search listbox
    With lstTypes
        For lngFTIdx = 0 To (lngTotalPatterns - 1)
            .AddItem UCase$(a_strTypes(lngFTIdx))
        Next lngFTIdx
        .ListIndex = 0
    End With

    'load the default text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    cmdClear.Caption = g_strButtonClearText
    With g_xobjFunc
        cmdFind.Caption = .GetINIMessage("Music Selector", "FindButton", "&Find")
        chkAll.Caption = .GetINIMessage("Music Selector", "ApplyToAll", "&Apply to all games/music in current view")
        Caption = .GetINIMessage("Music Selector", "Caption", "Music Selector")
        tabSearch.TabCaption(0) = .GetINIMessage("Music Selector", "ManualSearch", "&Manual Search")
        tabSearch.TabCaption(1) = .GetINIMessage("Music Selector", "AutoSearch", "&Auto Search")
        lblError = .GetINIMessage("Music Selector", "NoneSelected", "None selected")
        fraInfo.Caption = .GetINIMessage("Music Selector", "Tab_Sid_Label_InfoTitle", "Sid-File Info")
        fraNone.Caption = fraInfo.Caption
        lblPName = .GetINIMessage("Music Selector", "Tab_Sid_Label_InfoName", "Name:")
        lblPAuthor = .GetINIMessage("Music Selector", "Tab_Sid_Label_InfoAuthor", "Author:")
        lblPCopyright = .GetINIMessage("Music Selector", "Tab_Sid_Label_InfoCopyright", "Copyright:")
        lblSearching = .GetINIMessage("Music Selector", "Searching", "Searching") & "..."
    End With

    'are we adding?
    If g_blnSidSelectorNewItem = True Then
        chkAll.Enabled = False
    End If

    g_strSidSelectorSidName = ""
    
    '====================================================================
    'set the initial file and folder  --  ROOT HVSC Path assumed to exist
    '====================================================================

    If g_strSidSelectorFilename = g_strNone Then
        'no filename
        fraInfo.Visible = False
        fraNone.Visible = True
        cmdClear.Enabled = False
        'set starting folder
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastMusicSelectorPath", strLastBrowsePath, True, g_udtRegPathSid.strRealPath
        If WithinRootSidPath(strLastBrowsePath) Then
            If g_xobjFunc.DirExists(strLastBrowsePath) Then
                m_strLastPath = strLastBrowsePath
            Else
                m_strLastPath = g_udtRegPathSid.strRealPath
            End If
            DirBox.Path = m_strLastPath
        Else
            m_strLastPath = g_udtRegPathSid.strRealPath
            DirBox.Path = g_udtRegPathSid.strRealPath
        End If
        lblSid = g_strNoneSelected
    Else
        'filename set - see if exists
        strSidFile = g_udtRegPathSid.strWithBackSlash & g_strSidSelectorFilename
        g_xobjFunc.GetFileAndPath strSidFile, strPath, strFile
        If g_xobjFunc.FileExists(strSidFile, False, "") Then
            'file exists
            m_strLastPath = strPath
            DirBox.Path = strPath
            DoEvents 'allow to update file listbox
            Call SetSidInFileBox(strFile)
        Else
            'file doesn't exist
            MsgBox g_xobjFunc.GetINIMessage("Sid Selector", "FileNotFound", "The Sid file was not found."), vbInformation, gc_strAPPNAME
            If g_xobjFunc.DirExists(strPath) Then
                'we know it's within the root folder
                m_strLastPath = strPath
            Else
                'set the root folder as the sub folder doesn't exist
                m_strLastPath = g_udtRegPathSid.strRealPath
            End If
            DirBox.Path = m_strLastPath
            lblSid = FilenameForDatabase(strSidFile)
            fraInfo.Visible = False
            fraNone.Visible = True
            cmdClear.Enabled = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'release objects
    Set m_xobjFunc = Nothing

End Sub


Private Sub m_xobjFunc_FoundFile(ByVal strPathAndFile As String)

    lstFound.AddItem FilenameForDatabase(strPathAndFile)

End Sub


