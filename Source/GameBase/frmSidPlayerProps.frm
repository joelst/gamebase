VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSidPlayerProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sid Player Properties..."
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmSidPlayerProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFT 
      Height          =   285
      Left            =   120
      MaxLength       =   128
      TabIndex        =   10
      Top             =   2520
      Width           =   4815
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox chkShort 
      Caption         =   "&Use Short (8.3) Filenames"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox txtCLP 
      Height          =   285
      Left            =   120
      MaxLength       =   2048
      TabIndex        =   6
      Text            =   "%sidfile%"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      MaxLength       =   128
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblDelimit 
      BackStyle       =   0  'Transparent
      Caption         =   "Delimit with ;  (e.g. SID;MP3;MID...)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblFTPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Supported File Types:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label lblCLPPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Command &Line Parameters:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblPathPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Path/Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblNamePrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Player &Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmSidPlayerProps"
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

Private Sub SetOKButton()

    '-------------------------------
    'sets the state of the OK button
    '-------------------------------

    cmdOK.Enabled = ((txtName <> "") And (txtCLP <> "") And (lblPath <> "") And (txtFT.Text <> ""))

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
        .DialogTitle = g_xobjFunc.GetINIMessage("Music Player Property Window", "FilePrompt", "Select Music Player")
        .Filter = "Music Players (*.exe)|*.exe|Batch Files (*.bat;*.cmd)|*.bat;*.cmd|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastMusicPlayerPath", strLastDir, True, App.Path
        If lblPath = "" Then
            If g_xobjFunc.DirExists(strLastDir) = True Then
                .InitDir = strLastDir
            Else
                .InitDir = App.Path
            End If
            .Filename = ""
        ElseIf g_xobjFunc.FileExists(lblPath, False, "") = True Then
            .Filename = lblPath
        ElseIf g_xobjFunc.DirExists(strLastDir) = True Then
            .InitDir = strLastDir
            .Filename = ""
        Else
            .InitDir = App.Path
            .Filename = ""
        End If
    End With
    
    'get the file
    cDialog.ShowOpen
    lblPath = cDialog.Filename

    'save the last path
    g_xobjFunc.GetFileAndPath cDialog.Filename, strPath, strFile
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastMusicPlayerPath", strPath, True

Cancel_Pressed:

    'set the ok button
    Call SetOKButton

End Sub


Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdDefault_Click()

    With txtCLP
        .Text = "%musicfile%"
        .SetFocus
    End With

End Sub

Private Sub cmdOK_Click()

    Dim a_strTest() As String
    Dim lngTotal As Long

    'must have a file type at least
    Call SetSidFileTypes(txtFT.Text, a_strTest(), lngTotal)
    If lngTotal = 0 Then
        Call MsgBox(g_xobjFunc.GetINIMessage("Error Messages", "Error_InvalidFileTypes", "Please enter at least one valid file-type!"), vbInformation, gc_strAPPNAME)
        txtFT.SetFocus
        Exit Sub
    End If

    'check that the %musicfile% parameter is in the command line
    If InStr(UCase(txtCLP), "%MUSICFILE%") = 0 Then
        If MsgBox(g_xobjFunc.GetINIMessage("Music Player Property Window", "Warning", "You have not specified the %musicfile% parameter in the command line.  Are you sure you want to continue?"), vbQuestion Or vbYesNo, gc_strAPPNAME) = vbNo Then
            txtCLP.SetFocus
            Exit Sub
        End If
    End If

    'save info and close window
    With g_udtSidPlayerProps
        .blnUseShort = (chkShort = 1)
        .strCommandLine = txtCLP
        .strFullPath = lblPath
        .strName = txtName
        Call SetSidFileTypes(txtFT.Text, .a_strTypes(), .lngTotalTypes)
    End With
    g_blnChangedSidplayerSettings = True
    Unload Me

End Sub


Private Sub Form_Load()

    Dim lngIdx As Long
    Const DesignScaleHeight As Single = 3735    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5040     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the button text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    cmdBrowse.Caption = g_strButtonBrowseText
    cmdDefault.Caption = g_strButtonDefaultText

    'get ini file text
    With g_xobjFunc
        Caption = .GetINIMessage("Music Player Property Window", "Caption", "Music Player Properties") & "..."
        lblNamePrompt = .GetINIMessage("Music Player Property Window", "NamePrompt", "Player &Name") & ":"
        lblPathPrompt = .GetINIMessage("Music Player Property Window", "PathPrompt", "Player Path/Filename") & ":"
        lblCLPPrompt = .GetINIMessage("Music Player Property Window", "CommandLinePrompt", "Command &Line Parameters") & ":"
        chkShort.Caption = .GetINIMessage("Music Player Property Window", "ShortFilenamePrompt", "&Use Short (8.3) Filenames")
        lblFTPrompt = .GetINIMessage("Music Player Property Window", "SupportedFileTypes", "&Supported File Types") & ":"
    End With

    'fill the properties
    With g_udtSidPlayerProps
        txtName = .strName
        lblPath = .strFullPath
        txtCLP = .strCommandLine
        chkShort = Abs(.blnUseShort)
        txtFT.Text = ""
        If .lngTotalTypes > 0 Then
            For lngIdx = 0 To UBound(.a_strTypes)
                txtFT.Text = txtFT.Text & .a_strTypes(lngIdx) & ";"
            Next lngIdx
        End If
        'blank the name (means user pressed cancel)
        .strName = ""
    End With

    'set the ok button's enabled state
    Call SetOKButton

End Sub


Private Sub txtCLP_Change()

    Call SetOKButton

End Sub

Private Sub txtCLP_GotFocus()

    With txtCLP
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtFT_Change()

    Call SetOKButton

End Sub

Private Sub txtFT_GotFocus()

    With txtFT
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtFT_KeyPress(KeyAscii As Integer)

    If InStr("\/?*:""<>|", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtName_Change()

    Call SetOKButton

End Sub

Private Sub txtName_GotFocus()

    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
