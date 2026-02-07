VERSION 5.00
Begin VB.Form frmRomCenter2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RomCenter 2.00 Datafile Import"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmRomCenter2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton cmdHelpGameFilenameField 
         Caption         =   "?"
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton cmdHelpGameNameField 
         Caption         =   "?"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdHelpTosec 
         Caption         =   "?"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton cmdHelpAddZip 
         Caption         =   "?"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton cmdHelpIgnoreDups 
         Caption         =   "?"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkAddZip 
         Caption         =   "&Add ZIP extension to filename"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox chkTosec 
         Caption         =   "&TOSEC Filenames (TNC v1)"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   3375
      End
      Begin VB.ComboBox lstFile 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
      End
      Begin VB.ComboBox lstName 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkIgnoreDups 
         Caption         =   "Ignore &duplicate Game Names"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game &Filename Field:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game &Name Field:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraScan 
      Height          =   1695
      Left            =   4200
      TabIndex        =   13
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton cmdHelpScan 
         Caption         =   "?"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdHelpMatch 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkMatch 
         Caption         =   "Only search &matching filename extensions"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdRomPath 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtRomPath 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkScan 
         Caption         =   "Scan for &Roms"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rom &Path:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   5760
      MaxLength       =   64
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblDataFile 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RomCenter Datafile:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&GameBase Name:"
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmRomCenter2"
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

Private m_xobjFunc As GBFuncs.clsFunctions
Private m_blnActivated As Boolean
Private Function GetGBNameFromRCFilename(strFilename_in As String) As String

    Dim strName As String
    Dim strTemp As String
    Dim lngIdx As Long
    Dim strChar As String
    Const cILLEGALCHARS As String = "\/:*?""<>|"
    
    'get rid of filename extension
    strName = Left$(strFilename_in, Len(strFilename_in) - 4)

    'remove all illegal chars
    strTemp = ""
    For lngIdx = 1 To Len(strName)
        strChar = Mid$(strName, lngIdx, 1)
        If InStr(cILLEGALCHARS, strChar) = 0 Then
            strTemp = strTemp & strChar
        End If
    Next lngIdx

    'limit to 64 chars
    GetGBNameFromRCFilename = Left$(strTemp, 64)

    'make sure not an empty string
    If Len(GetGBNameFromRCFilename) = 0 Then
        GetGBNameFromRCFilename = "New GameBase"
    End If

End Function

Private Sub chkScan_Click()

    If chkScan.Value = vbChecked Then
        lblPrompt(2).Enabled = True
        txtRomPath.Enabled = True
        txtRomPath.BackColor = lstName.BackColor
        cmdRomPath.Enabled = True
        chkMatch.Enabled = True
        cmdHelpMatch.Enabled = True
        txtRomPath.SetFocus
    Else
        lblPrompt(2).Enabled = False
        txtRomPath.Enabled = False
        txtRomPath.BackColor = BackColor
        cmdRomPath.Enabled = False
        chkMatch.Enabled = False
        cmdHelpMatch.Enabled = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelpAddZip_Click()

    MsgBox "This field should be used for MAME type setups where the Game Name Field you have chosen does not represent the actual rom-filename, but the name of the game.  GameBase will add .zip to the game name so that your zipped roms can be found and/or linked to the database.", vbInformation, "Add ZIP extension to filenames"
    chkAddZip.SetFocus

End Sub

Private Sub cmdHelpGameFilenameField_Click()

    MsgBox "This is the column in the datafile that you want to use as the Game Filename in GameBase.", vbInformation, "Game Filename Field"
    lstFile.SetFocus

End Sub

Private Sub cmdHelpGameNameField_Click()

    MsgBox "This is the column in the datafile that you want to use as the Game Name in GameBase.", vbInformation, "Game Name Field"
    lstName.SetFocus

End Sub

Private Sub cmdHelpIgnoreDups_Click()

    MsgBox "Some datafiles have the same game name listed many times (e.g. multiple rom-file games)." & vbCrLf & vbCrLf & "If the column you chose as the Game Name Field has duplicate entries, select this option to" & vbCrLf & "ignore them and only import games with unique names.", vbInformation, "Ignore Duplicate Game Names"
    chkIgnoreDups.SetFocus

End Sub

Private Sub cmdHelpMatch_Click()

    MsgBox "If you select this option, GameBase will only scan for files whose filename extension matches any of those in the datafile." & vbCrLf & vbCrLf & "E.g.  If the datafile only contains .TAP and .D64 files, only these filetypes will be scanned.  Setting this option can speed up the scanning process.", vbInformation, "Only search matching filename extensions"
    chkMatch.SetFocus

End Sub

Private Sub cmdHelpScan_Click()

    MsgBox "Select this option if you want GameBase to try to locate any roms or game files you may already have on your hard drive.  This scan will be performed on the datafile column you chose as the 'Game Filename Field'.", vbInformation, "Scan for Roms"
    chkScan.SetFocus

End Sub

Private Sub cmdHelpTosec_Click()

    MsgBox "If the column you chose as the Game Filename Field in the datafile uses TOSEC (TNC v1) compliant filenames, select this option.", vbInformation, "TOSEC Filenames (TNC v1)"
    chkTosec.SetFocus

End Sub

Private Sub cmdOK_Click()

    Dim strRomPath As String
    Dim strRomPathWBS As String
    Dim strGBFolder As String
    Dim strGBFolderWBS As String

    cmdOK.Enabled = False
    DoEvents
    strRomPath = ""
    strRomPathWBS = ""
    strGBFolder = ""
    strGBFolderWBS = ""

    'make sure we have a valid gamebase folder name
    If g_blnRC2ExistingDB = False Then
        If txtName.Text <> "" Then
            strGBFolder = g_strAppPathWithBackSlash & txtName.Text
            If m_xobjFunc.DirExists(strGBFolder) = False Then
                strGBFolderWBS = strGBFolder
                If Right$(strGBFolderWBS, 1) <> "\" Then strGBFolderWBS = strGBFolderWBS & "\"
            Else
                MsgBox "There is already a GameBase folder with that name!", vbInformation, "GBToolbox"
                txtName.SetFocus
                cmdOK.Enabled = True
                Exit Sub
            End If
        Else
            MsgBox "You must enter a GameBase name!", vbInformation, "GBToolbox"
            txtName.SetFocus
            cmdOK.Enabled = True
            Exit Sub
        End If
    End If

    'get a valid rom path if scanning
    If chkScan.Value = vbChecked Then
        If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Left$(txtRomPath.Text, 1)) = 0) Or (Mid$(txtRomPath.Text, 2, 1) <> ":") Or (Mid$(txtRomPath.Text, 3, 1) <> "\") Then
            MsgBox "Please enter an absolute (full) Rom Path!", vbInformation, "GBToolbox"
            If txtRomPath.Enabled = True Then
                txtRomPath.SetFocus
            End If
            cmdOK.Enabled = True
            Exit Sub
        ElseIf m_xobjFunc.DirExists(txtRomPath.Text) = False Then
            MsgBox "The Rom Path could not be found!", vbInformation, "GBToolbox"
            If txtRomPath.Enabled = True Then
                txtRomPath.SetFocus
            End If
            cmdOK.Enabled = True
            Exit Sub
        End If
    End If
    strRomPath = txtRomPath.Text
    If Right$(strRomPath, 1) = "\" Then
        strRomPathWBS = strRomPath
        If Len(strRomPath) > 3 Then
            strRomPath = Left$(strRomPath, Len(strRomPath) - 1)
        End If
    Else
        strRomPathWBS = strRomPath & "\"
    End If

    'bring up the wait window and process the import
    With g_udtRC2ImportOptions
        .blnAddZip = (chkAddZip.Value = vbChecked)
        .blnIgnoreDups = (chkIgnoreDups.Value = vbChecked)
        .blnMatch = (chkMatch.Value = vbChecked)
        .blnScan = (chkScan.Value = vbChecked)
        .blnTosec = (chkTosec.Value = vbChecked)
        .lngFileColIdx = lstFile.ItemData(lstFile.ListIndex)
        .lngNameColIdx = lstName.ItemData(lstName.ListIndex)
        .strGameBaseName = txtName.Text
        .strGameBaseFolder = strGBFolder
        .strGameBaseFolderWBS = strGBFolderWBS
        .strRomPath = strRomPath
        .strRomPathWBS = strRomPathWBS
    End With
    frmProgress.Show vbModal, Me
    'all done and all flags set
    If g_blnCancelPressed = False Then
        Unload Me
    Else
        cmdOK.Enabled = True
        cmdOK.SetFocus
    End If

End Sub


Private Sub cmdRomPath_Click()

    If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Left$(txtRomPath.Text, 1)) = 0) Or (Mid$(txtRomPath.Text, 2, 1) <> ":") Or (Mid$(txtRomPath.Text, 3, 1) <> "\") Then
        g_strBrowseForFolder = ""
    Else
        If m_xobjFunc.DirExists(txtRomPath.Text) = True Then
            g_strBrowseForFolder = txtRomPath.Text
        Else
            g_strBrowseForFolder = ""
        End If
    End If
    frmBrowseFolder.Show vbModal, Me
    If g_strBrowseForFolder <> "" Then
        txtRomPath.Text = g_strBrowseForFolder
        If txtRomPath.Enabled = True Then
            txtRomPath.SetFocus
        End If
    End If

End Sub


Private Sub cmdView_Click()

    On Error GoTo Err_Handler

    Call Shell("notepad.exe """ & g_strRC2DataFile & """", vbNormalFocus)
    
    Exit Sub
    
Err_Handler:
    MsgBox "Oops! Could not view file! Check that notepad.exe is on your path!", vbExclamation, "GBToolbox"

End Sub

Private Sub Form_Activate()

    If m_blnActivated = False Then
        If g_blnRC2ExistingDB = False Then
            txtName.SetFocus
        Else
            lstName.SetFocus
        End If
        m_blnActivated = True
    End If

End Sub

Private Sub Form_Load()

    Dim strFile As String
    Dim strPath As String
    Const DesignScaleHeight As Single = 3615    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 8175     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'create objects and init vars
    Set m_xobjFunc = New GBFuncs.clsFunctions
    m_blnActivated = False
    
    'set datafile name
    Call m_xobjFunc.GetFileAndPath(g_strRC2DataFile, strPath, strFile)
    lblDataFile = strFile
    If InStr(LCase$(lblDataFile), "tosec") > 0 Then
        chkTosec.Value = vbChecked
    End If

    'set the GameBase name
    If g_blnRC2ExistingDB = False Then
        txtName = GetGBNameFromRCFilename(strFile)
    Else
        txtName = g_strRC2DBName
        txtName.BackColor = BackColor
        txtName.Locked = True
    End If

    'fill listboxes
    With lstName
        .AddItem "Parent Name (col 1)"
        .ItemData(.NewIndex) = 1
        .AddItem "Parent Description (col 2)"
        .ItemData(.NewIndex) = 2
        .AddItem "Game Name (col 3)"
        .ItemData(.NewIndex) = 3
        .AddItem "Game Description (col 4)"
        .ItemData(.NewIndex) = 4
        .AddItem "Rom Name (col 5)"
        .ItemData(.NewIndex) = 5
        .AddItem "RomOf Name (col 8)"
        .ItemData(.NewIndex) = 8
        .ListIndex = 2
    End With
    With lstFile
        .AddItem "Parent Name (col 1)"
        .ItemData(.NewIndex) = 1
        .AddItem "Game Name (col 3)"
        .ItemData(.NewIndex) = 3
        .AddItem "Rom Name (col 5)"
        .ItemData(.NewIndex) = 5
        .AddItem "RomOf Name (col 8)"
        .ItemData(.NewIndex) = 7
        .ListIndex = 2
    End With
    

    'set globals vars
    g_blnRC2Imported = False
    g_strRC2GBFolderWBS = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'release objects
    Set m_xobjFunc = Nothing

End Sub


Private Sub txtName_Change()

    cmdOK.Enabled = Len(txtName.Text) > 0

End Sub

Private Sub txtName_GotFocus()

    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)

    If InStr("\/:*?""<>|", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRomPath_GotFocus()

    With txtRomPath
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


