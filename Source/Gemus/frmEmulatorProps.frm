VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmulatorProps 
   Caption         =   "Emulator Properties"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmEmulatorProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8415
   Begin VB.CheckBox chkSyntax 
      Caption         =   "&Auto Syntax Check"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CheckBox chkWait 
      Caption         =   "&Wait until closed (if rezipping games)"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2700
      Width           =   3855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame fraTypes 
      Caption         =   "Supported Native File Types"
      Height          =   1455
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   4095
      Begin VB.TextBox txtTypes 
         Height          =   285
         Left            =   240
         MaxLength       =   4096
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox chkTypes 
         Caption         =   "Use &Defaults"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDelim 
         BackStyle       =   0  'Transparent
         Caption         =   "Delimiter text"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   3615
      End
   End
   Begin VB.CheckBox chkShort 
      Caption         =   "&Use Short (8.3) Filenames"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtEmulatorName 
      Height          =   285
      Left            =   120
      MaxLength       =   128
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Script Files"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3855
      Begin Gemus.ctlFileSelector ctlFiles 
         Height          =   1170
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
         _extentx        =   5953
         _extenty        =   2064
         browseregkey    =   "LastEmuFilePath"
      End
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   5280
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblScriptFile 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblPrompt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Line 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   18
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Emulator Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "GEMUS Script:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Emulator Path\Filename:"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblEmulatorExe 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmEmulatorProps"
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

Private Gemus As clsGemus
Private m_strGBName As String
Private m_strEmuFile As String
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_xobjRegistry As GBFuncs.clsRegistry
Private m_strMessFile As String
Private m_strLineWord As String
Private m_strGemusError As String
Private m_strAppPathWithBackSlash As String
Private m_strPreChangedScript As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_SETSEL = &HB1

Private Const mc_strDEFAULT_EMU_FILETYPES As String = ""
Private m_lngLastLine As Long                       'the last line the user was on
Private m_blnInFocus As Boolean
Private m_blnChangingScriptLineByCode As Boolean    'flag for changing script lines by code
Private m_strDefaultTypes As String
Private m_blnDoneFirstTime As Boolean

Private Const mc_strDefaultScript As String = "Add_CLP(%gamepathfile%)" & vbCrLf & "Run_Emulator()" & vbCrLf
Private Sub ProcessLineChange()

    '---------------------------------------------------------
    'determines whether we need to validate the last left line
    'if it has changed
    '---------------------------------------------------------

    Dim lngNewLine As Long
    Dim lngSelStart As Long
    Dim lngSelLen As Long
    Dim strError As String

    'get new line number
    lngNewLine = SendMessage(txtScript.hWnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    If lngNewLine > 0 Then
        lblPrompt(3) = m_strLineWord & " " & lngNewLine
    Else
        lblPrompt(3) = ""
    End If
    DoEvents

    'validate the last line
    If chkSyntax.Value = vbChecked Then
        If (lngNewLine <> m_lngLastLine) And (m_lngLastLine > 0) Then
            strError = Gemus.ValidateScriptLine(Gemus.GetScriptLineText(txtScript.Text, m_lngLastLine))
            If strError <> "" Then
                MsgBox strError & vbCrLf & vbCrLf & m_strLineWord & " " & m_lngLastLine, vbExclamation, m_strGemusError
                txtScript.SetFocus
            Else
                m_lngLastLine = lngNewLine
            End If
        Else
            m_lngLastLine = lngNewLine
        End If
    Else
        m_lngLastLine = lngNewLine
    End If

End Sub

Private Sub SetOKButton()

    '-------------------------------
    'sets the state of the OK button
    '-------------------------------

    Dim lngIdx As Long

    If (txtEmulatorName <> "") And (lblEmulatorExe <> "") Then
        With txtTypes
            If InStr(.Text, ";;") > 0 Then
                cmdOK.Enabled = False
                Exit Sub
            End If
            For lngIdx = 1 To Len(.Text)
                If Mid$(.Text, lngIdx, 1) <> ";" Then
                    cmdOK.Enabled = True
                    Exit Sub
                End If
            Next lngIdx
            cmdOK.Enabled = False
        End With
    Else
        cmdOK.Enabled = False
    End If

End Sub

Private Sub chkShort_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        chkWait.SetFocus
    End If

End Sub

Private Sub chkSyntax_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        If cmdOK.Enabled = True Then
            cmdOK.SetFocus
        Else
            cmdCancel.SetFocus
        End If
    End If

End Sub


Private Sub chkWait_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        cmdLoad.SetFocus
    End If

End Sub

Private Sub chkTypes_Click()

    With txtTypes
        If chkTypes.Value = vbChecked Then
            .Text = m_strDefaultTypes
            .Enabled = False
            .BackColor = fraTypes.BackColor
            Call SetOKButton
        Else
            .Enabled = True
            .BackColor = txtEmulatorName.BackColor
            .SetFocus
        End If
    End With

End Sub

Private Sub chkTypes_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        If txtTypes.Enabled = True Then
            txtTypes.SetFocus
        Else
            chkShort.SetFocus
        End If
    End If

End Sub


Private Sub cmdBrowse_Click()

    '----------------------------------
    'get the filename of the sid player
    '----------------------------------

    Dim strLastDir As String    'last dir traversed
    Dim strFile As String       'for splitting into path\file vars
    Dim strPath As String       'for splitting into path\file vars

    On Error GoTo Cancel_Pressed

    'set up the open dialog
    With cDialog
        .CancelError = True
        .DefaultExt = ".exe"
        .DialogTitle = GetINIMessageFromFile("Emulator Property Window", "FilePrompt", "Select Emulator", m_strMessFile)
        .Filter = "Emulators (*.exe)|*.exe|Batch Files (*.bat;*.cmd)|*.bat;*.cmd|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        m_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\BrowsePaths", "LastEmulatorPath", strLastDir, True, App.Path
        If lblEmulatorExe = "" Then
            If m_xobjFunc.DirExists(strLastDir) = True Then
                .InitDir = strLastDir
            Else
                .InitDir = App.Path
            End If
            .Filename = ""
        ElseIf m_xobjFunc.FileExists(lblEmulatorExe, False, "") = True Then
            .Filename = lblEmulatorExe
        ElseIf m_xobjFunc.DirExists(strLastDir) = True Then
            .InitDir = strLastDir
            .Filename = ""
        Else
            .InitDir = App.Path
            .Filename = ""
        End If
    End With
    
    'get the file
    cDialog.ShowOpen
    lblEmulatorExe = cDialog.Filename

    'save the last path
    m_xobjFunc.GetFileAndPath cDialog.Filename, strPath, strFile
    m_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & m_strGBName & "\BrowsePaths", "LastEmulatorPath", strPath, True
    Call ctlFiles.SetLastTraversedPath(strPath)

Cancel_Pressed:

    'set the ok button
    Call SetOKButton

End Sub

Private Sub cmdBrowse_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        ctlFiles.SetFocus
    End If

End Sub


Private Sub cmdCancel_Click()

    If m_strPreChangedScript <> txtScript.Text Then
        If MsgBox("Discard Script changes?", vbQuestion Or vbYesNo Or vbDefaultButton2, "GEMUS") = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me

End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        txtEmulatorName.SetFocus
    End If

End Sub







Private Sub cmdHelp_Click()

    Call m_xobjFunc.ShellExecuteFile("http://www.bu22.com/wiki/gemus_script_reference/home", "", False)
    
End Sub

Private Sub cmdHelp_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        chkSyntax.SetFocus
    End If

End Sub




Private Sub cmdLoad_Click()

    Dim f As Integer
    Dim strFolder As String
    Dim Func As GBFuncs.clsFunctions
    
    On Error GoTo Err_Handler
    
    strFolder = m_strAppPathWithBackSlash & m_strGBName & "\Scripts\"
    Set Func = New GBFuncs.clsFunctions
    If Func.DirExists(Left$(strFolder, Len(strFolder) - 1)) = False Then
        MkDir strFolder
    End If
    Set Func = Nothing
    g_udtScriptSelectorType.blnOKClicked = False
    g_udtScriptSelectorType.strInitPath = strFolder
    g_udtScriptSelectorType.strSelectedFile = lblScriptFile.Caption
    frmScriptSelector.Show vbModal, Me
    If g_udtScriptSelectorType.blnOKClicked = True Then
        txtScript.Text = ""
        f = FreeFile()
        Open (strFolder & g_udtScriptSelectorType.strSelectedFile) For Input As #f
        If LOF(f) > 0 Then
            If LOF(f) > 65526 Then
                'too large for textbox
                txtScript.Text = Input$(65526, f)
            Else
                'size ok
                txtScript.Text = Input$(LOF(f), f)
            End If
            txtScript.SelStart = 0
        End If
        Close #f
        lblScriptFile.Caption = g_udtScriptSelectorType.strSelectedFile
        txtScript.SetFocus
        m_strPreChangedScript = txtScript.Text
    End If
    Exit Sub

Err_Handler:
    On Error GoTo 0
    MsgBox "Unable to load a GEMUS Script.  " & Err.Description & " (" & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")", vbCritical, "GEMUS"

End Sub

Private Sub cmdLoad_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        txtScript.SetFocus
    End If

End Sub


Private Sub cmdOK_Click()

    Dim lngFileIdx As Long      'loop index
    Dim strError As String      'error result
    Dim lngErrorLine As Long    'error line
    Dim strReplace As String    'replacement text
    Dim lngLineIdx As Long      'loop index
    Dim lngLastLineIdx As Long  'last actual line of script
    Dim f As Integer
    Dim strFolder As String
    Dim strFilename As String
    Dim strScriptFile As String

    On Error GoTo Err_Handler

    strReplace = ""
    strScriptFile = ""
    If txtScript <> "" Then

        'get and validate the script lines
        strError = Gemus.ValidateAndGetScriptLines(txtScript)
        If strError <> "" Then
            'bad line(s) found in script
            MsgBox strError, vbExclamation, m_strGemusError
            txtScript.SetFocus
            Exit Sub
        ElseIf Gemus.TotalScriptLines = 0 Then
            'no lines found in script!
            If MsgBox(GetINIMessageFromFile("Emulator Script", "Error_NoScript", "There is no script set for this emulator!  Save details anyway? (games won't run!)", m_strMessFile), vbYesNo Or vbQuestion, m_strGemusError) = vbNo Then
                txtScript.SetFocus
                Exit Sub
            Else
                strReplace = txtScript.Text
            End If
        Else
            'validate the script structure
            strError = Gemus.ValidateScriptStructure(txtScript.Text, lngErrorLine)
            If strError <> "" Then
                'invalid script structure
                If lngErrorLine > -1 Then
                    strError = strError & vbCrLf & vbCrLf & m_strLineWord & " " & lngErrorLine
                End If
                MsgBox strError, vbExclamation, m_strGemusError
                txtScript.SetFocus
                Exit Sub
            Else
                'format the script lines
                For lngLineIdx = (Gemus.TotalScriptLines - 1) To 0 Step -1
                    If Gemus.ScriptLines(lngLineIdx) <> "" Then
                        lngLastLineIdx = lngLineIdx
                        Exit For
                    End If
                Next lngLineIdx
                strReplace = ""
                For lngLineIdx = 0 To lngLastLineIdx
                    strReplace = strReplace & Gemus.ScriptLinesPre(lngLineIdx) & Gemus.ScriptLines(lngLineIdx) & vbCrLf
                Next lngLineIdx
            End If
        End If
    End If

    'save the script file
    If strReplace <> "" Or lblScriptFile.Caption <> "" Then
        strFolder = m_strAppPathWithBackSlash & m_strGBName & "\Scripts\"
        If m_xobjFunc.DirExists(Left$(strFolder, Len(strFolder) - 1)) = False Then
            MkDir strFolder
        End If
        If lblScriptFile.Caption = "" Then
            'new file
            g_strScriptSaveVar = strFolder
            frmGetFilename.Show vbModal, Me
            If g_strScriptSaveVar = "" Then
                Exit Sub
            Else
                strFilename = strFolder & g_strScriptSaveVar
                strScriptFile = g_strScriptSaveVar
            End If
        Else
            'replace file
            strFilename = strFolder & lblScriptFile.Caption
            strScriptFile = lblScriptFile.Caption
        End If
        'write file
        m_xobjFunc.DeletePathSilent strFilename
        f = FreeFile()
        Open strFilename For Output As #f
        Print #f, strReplace
        Close #f
    End If

    'Save info and close window if here!
    With g_udtEmulatorProps
        .blnUseShort = (chkShort = vbChecked)
        .strScript = strScriptFile
        .strFullPath = lblEmulatorExe
        .strName = txtEmulatorName
        .blnUseDefaultTypes = (chkTypes = vbChecked)
        .strTypes = txtTypes
        .lngWait = IIf((chkWait.Value = vbChecked), 1, 0)
        .lngNumFiles = ctlFiles.ListCount
        ReDim .strFiles(0)
        If .lngNumFiles > 0 Then
            ReDim .strFiles(.lngNumFiles - 1)
            For lngFileIdx = 0 To (.lngNumFiles - 1)
                .strFiles(lngFileIdx) = ctlFiles.List(lngFileIdx)
            Next lngFileIdx
        End If
    End With
    Unload Me
    Exit Sub

Err_Handler:
    On Error GoTo 0
    g_udtEmulatorProps.strName = ""
    MsgBox "Unable to save emulator details!  " & Err.Description & " (" & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")", vbCritical, "GEMUS"

End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        cmdCancel.SetFocus
    End If

End Sub

Private Sub ctlFiles_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        If chkTypes.Enabled = True Then
            chkTypes.SetFocus
        Else
            txtTypes.SetFocus
        End If
    End If

End Sub


Private Sub Form_Activate()

    If m_blnDoneFirstTime = False Then
        'txtScript.SetFocus
        txtEmulatorName.SetFocus
        m_blnDoneFirstTime = True
    End If

End Sub

Private Sub Form_Load()

    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips
    Const DesignScaleHeight As Single = 6255    'The required minimum scaleheight for the form
    Const DesignScaleWidth As Single = 8415     'The required minimum scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim blnRegMaximized As Boolean              'last set window maximized state
    Dim strLastFilePath As String               'last emulator file path traversed
    Dim strFile As String                       'last files file
    Dim strPath As String                       'last files path
    Dim lngFileIdx As Long                      'loop index
    Dim f As Integer
    Dim lngAutoSyntaxCheck As Long

    'create the references
    Set Gemus = New clsGemus
    Set m_xobjFunc = New GBFuncs.clsFunctions
    Set m_xobjRegistry = New GBFuncs.clsRegistry

    'init vars
    m_blnDoneFirstTime = False
    m_strAppPathWithBackSlash = App.Path
    If g_blnInVBIDE Then
        m_strAppPathWithBackSlash = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    If Right$(m_strAppPathWithBackSlash, 1) <> "\" Then m_strAppPathWithBackSlash = m_strAppPathWithBackSlash & "\"
    m_strMessFile = m_strAppPathWithBackSlash & "GemusText.ini"
    m_strLineWord = GetINIMessageFromFile("Emulator Script", "Line", "Line", m_strMessFile)
    m_strGemusError = GetINIMessageFromFile("Emulator Script", "GEMUSError", "GEMUS Error", m_strMessFile)
    Call Gemus.Init(g_udtEmulatorProps.strGBName)
    m_strGBName = g_udtEmulatorProps.strGBName
    m_strEmuFile = g_udtEmulatorProps.strEmuFile

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngScriptWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngScriptWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngScriptWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngScriptWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With m_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ScriptWidth", sngRegWidth, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ScriptHeight", sngRegHeight, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ScriptTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ScriptLeft", sngRegLeft, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ScriptMaximized", blnRegMaximized, False, False
        'auto syntax check
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "AutoSyntaxCheck", lngAutoSyntaxCheck, False, vbChecked
        If lngAutoSyntaxCheck <> vbChecked And lngAutoSyntaxCheck <> vbUnchecked Then
            lngAutoSyntaxCheck = vbChecked
        End If
        chkSyntax.Value = lngAutoSyntaxCheck
    End With

    'make sure not to big or small...
    If sngRegWidth < sngMinTwipsWidth Then
        sngRegWidth = sngMinTwipsWidth
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight < sngMinTwipsHeight Then
        sngRegHeight = sngMinTwipsHeight
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    '...or out of full screen view
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (Height / 2)
        sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    'set the position and size
    If blnRegMaximized = False Then
        Top = sngRegTop
        Left = sngRegLeft
    Else
        WindowState = vbMaximized
    End If
    Width = sngRegWidth
    Height = sngRegHeight

    'subclass the window
    If g_blnInVBIDE = False Then
        Call ScriptSubClass(Me.hWnd)
    End If

    'set default button text
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", m_strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", m_strMessFile)
    cmdBrowse.Caption = GetINIMessageFromFile("General Buttons", "Browse", "&Browse", m_strMessFile)
    cmdHelp.Caption = GetINIMessageFromFile("General Buttons", "Help", "&Help", m_strMessFile)

    'get ini file text
    cmdLoad.Caption = GetINIMessageFromFile("General Buttons", "Load", "&Load", m_strMessFile)
    Caption = GetINIMessageFromFile("Emulator Property Window", "Caption", "Emulator Properties", m_strMessFile) & "..."
    lblPrompt(2) = GetINIMessageFromFile("Emulator Property Window", "NamePrompt", "Emulator &Name", m_strMessFile) & ":"
    lblPrompt(0) = GetINIMessageFromFile("Emulator Property Window", "PathPrompt", "Emulator Path\Filename", m_strMessFile) & ":"
    fraFiles.Caption = GetINIMessageFromFile("Emulator Property Window", "FilesPrompt", "Associated Emulator Files", m_strMessFile)
    fraTypes.Caption = GetINIMessageFromFile("Emulator Property Window", "FileTypesPrompt", "Supported Native File Types", m_strMessFile)
    lblPrompt(1) = GetINIMessageFromFile("Emulator Property Window", "ScriptPrompt", "GEMUS &Script", m_strMessFile) & ":"
    chkShort.Caption = GetINIMessageFromFile("Emulator Property Window", "ShortFilenamePrompt", "&Use Short (8.3) Filenames", m_strMessFile)
    chkWait.Caption = GetINIMessageFromFile("Emulator Property Window", "WaitFinishPrompt", "&Wait until closed (if rezipping games)", m_strMessFile)
    chkTypes.Caption = GetINIMessageFromFile("Emulator Property Window", "UseDefaultFileTypesPrompt", "Use &Defaults", m_strMessFile)
    lblDelim.Caption = GetINIMessageFromFile("Emulator Property Window", "DelimText", "Delimit with", m_strMessFile) & " ;"
    chkSyntax.Caption = GetINIMessageFromFile("Emulator Property Window", "AutoSyntaxCheck", "&Auto Syntax Check", m_strMessFile)

    m_strDefaultTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", mc_strDEFAULT_EMU_FILETYPES, m_strEmuFile)

    'fill the properties
    m_blnChangingScriptLineByCode = False
    With g_udtEmulatorProps
        txtEmulatorName = .strName
        lblEmulatorExe = .strFullPath
        If .strScript <> "" Then
            If m_xobjFunc.FileExists(m_strAppPathWithBackSlash & m_strGBName & "\Scripts\" & .strScript, False, "GEMUS") = True Then
                f = FreeFile()
                Open (m_strAppPathWithBackSlash & m_strGBName & "\Scripts\" & .strScript) For Input As #f
                If LOF(f) > 65526 Then
                    txtScript.Text = Input$(65526, f)
                Else
                    txtScript.Text = Input$(LOF(f), f)
                End If
                Close #f
                lblScriptFile.Caption = .strScript
            Else
                lblScriptFile.Caption = ""
                txtScript.Text = ""
            End If
        Else
            lblScriptFile.Caption = ""
            txtScript.Text = mc_strDefaultScript
        End If
        
        chkShort = Abs(.blnUseShort)
        chkWait = IIf((.lngWait = 1), vbChecked, vbUnchecked)
        'blank the name (means user pressed cancel)
        .strName = ""
        For lngFileIdx = 0 To (.lngNumFiles - 1)
            ctlFiles.AddItem .strFiles(lngFileIdx)
        Next lngFileIdx

        If .blnUseDefaultTypes = True Then
            txtTypes.Text = m_strDefaultTypes
            chkTypes.Value = vbChecked
        Else
            txtTypes.Text = .strTypes
        End If
        ctlFiles.GameBaseName = m_strGBName
    End With

    'load last selected emulator file path
    m_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & m_strGBName & "\BrowsePaths", ctlFiles.BrowseRegKey, strLastFilePath, True, App.Path
    If m_xobjFunc.FileExists(lblEmulatorExe, False, "") Then
        m_xobjFunc.GetFileAndPath lblEmulatorExe, strPath, strFile
        ctlFiles.SetLastTraversedPath strPath
    Else
        ctlFiles.SetLastTraversedPath strLastFilePath
    End If

    'setup variables / controls
    m_lngLastLine = 0
    txtScript.MaxLength = 65526
    m_blnInFocus = False
    m_strPreChangedScript = txtScript.Text

    'set the ok button
    Call SetOKButton

End Sub


Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        'horizontal
        cmdCancel.Left = ScaleWidth - 120 - cmdCancel.Width
        cmdOK.Left = cmdCancel.Left - 120 - cmdOK.Width
        txtScript.Width = ScaleWidth - 120 - txtScript.Left
        lblPrompt(3).Left = (txtScript.Left + txtScript.Width) - lblPrompt(3).Width
    
        'vertical
        cmdCancel.Top = ScaleHeight - 120 - cmdCancel.Height
        cmdOK.Top = cmdCancel.Top
        cmdHelp.Top = cmdCancel.Top
        chkSyntax.Top = cmdHelp.Top
        txtScript.Height = cmdCancel.Top - 120 - txtScript.Top
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnScriptSubclassed = True Then
            Call ScriptUnSubClass(Me.hWnd)
        End If
    End If

    'save window sizes etc...
    With m_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ScriptMaximized", True, False
        End If
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "AutoSyntaxCheck", chkSyntax.Value, False
    End With

    'last emulator file path
    m_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & m_strGBName & "\BrowsePaths", ctlFiles.BrowseRegKey, ctlFiles.GetLastTraversedPath, True

    'free class references
    Set Gemus = Nothing
    Set m_xobjFunc = Nothing
    Set m_xobjRegistry = Nothing

End Sub

Private Sub txtEmulatorName_Change()

    Call SetOKButton

End Sub

Private Sub txtEmulatorName_GotFocus()

    With txtEmulatorName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtEmulatorName_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        cmdBrowse.SetFocus
    End If

End Sub


Private Sub txtScript_GotFocus()

    cmdOK.Default = False
    
    On Local Error Resume Next
    'get line number
    m_lngLastLine = SendMessage(txtScript.hWnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    If m_lngLastLine > 0 Then
        lblPrompt(3) = m_strLineWord & " " & m_lngLastLine
    Else
        lblPrompt(3) = ""
    End If
    m_blnInFocus = True

End Sub


Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim lngNewLine As Long      'the new line

    'get new line number
    lngNewLine = SendMessage(txtScript.hWnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    If lngNewLine > 0 Then
        lblPrompt(3) = m_strLineWord & " " & lngNewLine
    Else
        lblPrompt(3) = ""
    End If
    DoEvents

End Sub

Private Sub txtScript_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim lngNewLine As Long      'the new line
    
    On Local Error Resume Next
    If KeyCode = 65 And Shift = 2 Then
        txtScript.SelStart = 0
        txtScript.SelLength = Len(txtScript.Text)
    ElseIf (KeyCode = 13) Or (KeyCode = 10) Or (KeyCode = 37) Or (KeyCode = 38) Or (KeyCode = 39) Or (KeyCode = 40) Or (KeyCode = 8) Then
        Call ProcessLineChange
    End If

End Sub

Private Sub txtScript_LostFocus()

    lblPrompt(3) = ""
    cmdOK.Default = True
    m_blnInFocus = False

End Sub

Private Sub txtScript_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If m_blnInFocus = True Then
        Call ProcessLineChange
    End If

End Sub


Private Sub txtTypes_Change()

    Call SetOKButton

End Sub

Private Sub txtTypes_GotFocus()

    With txtTypes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtTypes_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = vbTab Then
        KeyAscii = 0
        chkShort.SetFocus
    ElseIf InStr(" \/:*?""<>|", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub


