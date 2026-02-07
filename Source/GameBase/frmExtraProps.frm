VERSION 5.00
Begin VB.Form frmExtraProps 
   Caption         =   "Extra Properties"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "frmExtraProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8295
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   3840
      MaxLength       =   255
      TabIndex        =   27
      Text            =   "http://"
      Top             =   360
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame fraGemus 
      Caption         =   "G&EMUS Fields"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   3495
      Begin VB.TextBox txtImageIndex 
         Height          =   285
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "0"
         Top             =   320
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edi&t..."
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optGemus 
         Caption         =   "Use T&hese Fields..."
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optGemus 
         Caption         =   "Use Cu&rrent Game Fields"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   3015
      End
      Begin VB.OptionButton optGemus 
         Caption         =   "Use GEMUS &Default Fields"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Image Index:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "&Type"
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3495
      Begin VB.OptionButton optType 
         Caption         =   "&URL"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   3135
      End
      Begin VB.OptionButton optType 
         Caption         =   "&MUSIC"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   3135
      End
      Begin VB.OptionButton optType 
         Caption         =   "&GEMUS"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   3135
      End
      Begin VB.OptionButton optType 
         Caption         =   "&STANDARD"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblURL 
         Caption         =   "The Extra is treated as a URL (web link) and is opened in a web browser."
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         UseMnemonic     =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblMusic 
         Caption         =   "Uses a Music Player to run the Extra. Treats the Extra as a Music File."
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblGemus 
         Caption         =   "Uses GEMUS to run the Extra. The Extra is treated as though it is a Game File."
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblNormal 
         Caption         =   "Uses File Associations to run the Extra. For example, if the Extra was a text file, it might be loaded into notepad.exe."
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   3015
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Frame fraFile 
      Caption         =   "&File"
      Height          =   6375
      Left            =   3720
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdChange 
         Caption         =   "Ch&ange"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   6000
         Width           =   855
      End
      Begin VB.FileListBox FileBox 
         Height          =   2235
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   4215
      End
      Begin VB.DirListBox DirBox 
         Height          =   1890
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox lstPaths 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblRunnable 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   6000
         Width           =   3375
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Runnable file in archive:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   5760
         UseMnemonic     =   0   'False
         Width           =   6240
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected File:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   5160
         UseMnemonic     =   0   'False
         Width           =   6240
      End
      Begin VB.Label lblFile 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5400
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   20
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "UR&L:"
      Height          =   315
      Index           =   3
      Left            =   3840
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.Label lblDummy 
      BackColor       =   &H80000011&
      Caption         =   "Label1"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Name:"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "frmExtraProps"
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

Private m_udtGemusFields As ManualGEMUSFields
Private m_blnByCode As Boolean
Private m_strLastPath As String
Private m_blnFirstTimeItemClick As Boolean
Private m_blnLoadBusy As Boolean

Private Sub ClickGemusFile()


    Dim blnIsRepackable As Boolean
    Dim blnSuccess As Boolean
    Dim blnUserMustBrowseArchive As Boolean
    Dim blnUserCancelled As Boolean

    Dim lngTotalValidFileTypes As Long

    Dim strFileType As String
    Dim strFile As String
    Dim strStatus As String
    Dim strPathFileToTest As String
    Dim strExtractResult As String
    Dim strSelectedPathFile As String
    Dim strRunnable As String
    Dim strError As String
    Dim strNotUsed As String
    Dim strUnpackToWBS As String
    Dim strExt As String
    Dim a_strFileTypes() As String

    Dim Unpacker As Gemus.clsUnpack

    'sure?
    If m_blnLoadBusy = True Then Exit Sub

    Set Unpacker = New Gemus.clsUnpack

    'get the proper file to display
    strFile = FileBox.List(FileBox.ListIndex)
    strPathFileToTest = DirBox.Path
    If Right$(strPathFileToTest, 1) <> "\" Then
        strPathFileToTest = strPathFileToTest & "\"
    End If
    strPathFileToTest = strPathFileToTest & strFile
    MousePointer = vbHourglass
    DoEvents
    lngTotalValidFileTypes = g_xobjGEMUS.GetAllValidFileTypesArray(a_strFileTypes())
    
    If m_blnFirstTimeItemClick Then
        strRunnable = g_udtExtrasProps.strRunnableFile
    Else
        strRunnable = ""
        lblRunnable.Caption = ""
    End If
    m_blnFirstTimeItemClick = False

    '********************************* try to unpack/display the chosen file *********************************
    blnSuccess = False
    blnUserCancelled = False
    strUnpackToWBS = g_udtPathGBGame.strWithBackSlash
    If Unpacker.Unpack(strPathFileToTest, strRunnable, a_strFileTypes, lngTotalValidFileTypes, _
                                    strUnpackToWBS, blnIsRepackable, strSelectedPathFile, _
                                    blnUserMustBrowseArchive, strError, strNotUsed) _
    Then
        If blnUserMustBrowseArchive Then
            'more than one compatible game file in archive, ask user to pick one
            If Unpacker.GetRunnableFileInArchiveFromUser(strPathFileToTest, strUnpackToWBS, a_strFileTypes, strRunnable) Then
                'try again now the user picked a runnable file
                If Unpacker.Unpack(strPathFileToTest, strRunnable, a_strFileTypes, lngTotalValidFileTypes, _
                                                strUnpackToWBS, blnIsRepackable, strSelectedPathFile, _
                                                blnUserMustBrowseArchive, strError, strNotUsed) _
                Then
                    If blnUserMustBrowseArchive Then
                        'should NEVER get here!
                        strError = "Unable to find runnable file in archive."
                    Else
                        'found runnable file
                        blnSuccess = True
                    End If
                End If
            Else
                blnUserCancelled = True
            End If
        Else
            'found runnable file
            blnSuccess = True
        End If
    End If
    
    '********************************************************************************************************

    If blnSuccess Then
        strExt = UCase$(Right$(strSelectedPathFile, Len(strSelectedPathFile) - InStrRev(strSelectedPathFile, ".")))
        'ok
        lblRunnable = strRunnable
        lblRunnable.Enabled = True
        cmdChange.Enabled = Unpacker.IsFileASupportedArchiveFileType(strPathFileToTest)
    ElseIf blnUserCancelled = True Then
        lblRunnable.Enabled = False
        lblRunnable.Caption = "UNSET"
        cmdChange.Enabled = Unpacker.IsFileASupportedArchiveFileType(strPathFileToTest)
    Else
        lblRunnable.Enabled = False
        lblRunnable.Caption = strError
        cmdChange.Enabled = False
    End If
    
    MousePointer = vbDefault
    Set Unpacker = Nothing

End Sub

Private Function GetSubPathFromFullPath(strFullPath_in) As String

    Dim strPath As String
    Dim lngIdx As Long

    strPath = ga_udtRegPathExtras(lstPaths.ItemData(lstPaths.ListIndex)).strWithBackSlash
    GetSubPathFromFullPath = Right$(strFullPath_in, Len(strFullPath_in) - Len(strPath))

End Function

Private Sub SetOKButton()

    If optType(0).Value = True Then
        cmdOK.Enabled = ((Len(txtName.Text) > 0)) _
                        And (Len(lblFile.Caption) > 0)
    ElseIf optType(1).Value = True Then
        cmdOK.Enabled = ((Len(txtImageIndex.Text) > 0) _
                        And (Len(txtName.Text) > 0)) _
                        And (Len(lblFile.Caption) > 0) _
                        And (lblRunnable.Enabled = True)
    ElseIf optType(2).Value = True Then
        cmdOK.Enabled = ((Len(txtName.Text) > 0)) _
                        And (Len(lblFile.Caption) > 0)
    Else
        cmdOK.Enabled = ((Len(txtName.Text) > 0)) _
                        And (Len(txtURL.Text) > 0)
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdChange_Click()

    Dim blnDidChoose As Boolean
    Dim lngTotalValidFileTypes As Long
    Dim strUnpackToWBS As String
    Dim strPathFileToTest As String
    Dim strFile As String
    Dim strRunnable As String
    Dim a_strFileTypes() As String
    Dim Unpacker As Gemus.clsUnpack
    Dim lngListIndex As Long

    'get the proper file to display
    strFile = FileBox.List(FileBox.ListIndex)
    strPathFileToTest = DirBox.Path
    If Right$(strPathFileToTest, 1) <> "\" Then
        strPathFileToTest = strPathFileToTest & "\"
    End If
    strPathFileToTest = strPathFileToTest & strFile
    lngTotalValidFileTypes = g_xobjGEMUS.GetAllValidFileTypesArray(a_strFileTypes())

    strUnpackToWBS = g_udtPathGBGame.strWithBackSlash

    Set Unpacker = New Gemus.clsUnpack
    blnDidChoose = Unpacker.GetRunnableFileInArchiveFromUser(strPathFileToTest, strUnpackToWBS, a_strFileTypes, strRunnable)
    Set Unpacker = Nothing

    If blnDidChoose Then
        g_udtExtrasProps.strRunnableFile = strRunnable
        m_blnFirstTimeItemClick = True
        lngListIndex = FileBox.ListIndex
        FileBox.ListIndex = -1
        FileBox.ListIndex = lngListIndex
        FileBox.SetFocus
    End If

End Sub

Private Sub cmdEdit_Click()

    Call g_xobjGEMUS.GetManualFieldsFromUser(m_udtGemusFields, False, g_strGameBaseName)
    optGemus(2).SetFocus

End Sub

Private Sub cmdOK_Click()

    'save the details
    With g_udtExtrasProps
        .blnGF_TDE = m_udtGemusFields.blnTDE
        If optGemus(0).Value = True Then
            .lngGemusOption = 0
        ElseIf optGemus(1).Value = True Then
            .lngGemusOption = 1
        Else
            .lngGemusOption = 2
        End If
        .lngGF_Control = m_udtGemusFields.lngControl
        .lngGF_ImageIndex = Val(txtImageIndex.Text)
        .lngGF_PalNTSC = m_udtGemusFields.lngPalNTSC
        .lngGF_Players = m_udtGemusFields.lngPlayers
        If optType(0).Value = True Then
            .lngType = 0
        ElseIf optType(1).Value = True Then
            .lngType = 1
        ElseIf optType(2).Value = True Then
            .lngType = 2
        Else
            .lngType = 3
        End If
        .strFile = IIf(optType(3).Value = False, lblFile, txtURL.Text)
        .strGF_Comment = m_udtGemusFields.strComment
        .strGF_Notes = m_udtGemusFields.strNotes
        .strGF_VComment = m_udtGemusFields.strVComment
        .strName = txtName.Text
        .strRunnableFile = lblRunnable.Caption
    End With

    'save the current directory for next time
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "ExtrasPropsPath", DirBox.Path, True

    'all done
    Unload Me

End Sub

Private Sub DirBox_Change()

    Dim lngIdx As Long

    'inside an extra path?
    With lstPaths
        For lngIdx = 0 To (.ListCount - 1)
            If UCase$(.List(lngIdx)) = UCase$(Left$(DirBox.Path, Len(.List(lngIdx)))) Then
                m_blnByCode = True
                .ListIndex = lngIdx
                m_blnByCode = False
                m_strLastPath = DirBox.Path
                FileBox.Path = DirBox.Path
                Exit Sub
            End If
        Next lngIdx
    End With

    'not allowed - outside screenshot paths
    DirBox.Path = m_strLastPath

End Sub

Private Sub DirBox_Click()

    DirBox.Path = DirBox.List(DirBox.ListIndex)

End Sub


Private Sub FileBox_Click()

    Dim strFullFile As String

    If FileBox.ListIndex > -1 Then
    
        If optType(1).Value = True Then
            Call ClickGemusFile
        End If
    
        strFullFile = FileBox.Path
        If Right$(strFullFile, 1) <> "\" Then strFullFile = strFullFile & "\"
        strFullFile = strFullFile & FileBox.List(FileBox.ListIndex)
        lblFile = GetSubPathFromFullPath(strFullFile)
    End If
    Call SetOKButton

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    MousePointer = vbHourglass
    DoEvents
    If KeyCode = 116 Then 'F5
        If UCase$(FileBox.Path) <> UCase$(DirBox.Path) Then
            FileBox.Path = DirBox.Path
        Else
            FileBox.Refresh
        End If
    End If
    MousePointer = vbDefault
    
End Sub


Private Sub Form_Load()

    Dim lngIdx As Long
    Dim strFullFile As String
    Dim strFile As String
    Const DesignScaleHeight As Single = 7215    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 8295     'The required scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim blnRegMaximized As Boolean              'last set window maximized state
    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngExtrasWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngExtrasWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngExtrasWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngExtrasWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "ExtrasWidth", sngRegWidth, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "ExtrasHeight", sngRegHeight, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "ExtrasTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "ExtrasLeft", sngRegLeft, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "ExtrasMaximized", blnRegMaximized, False, False
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
        Call ExtrasSubClass(Me.hwnd)
    End If

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        'without shortcuts
        Caption = .GetINIMessage("Extras Properties Window", "Caption", "Extras Properties for") & " " & g_udtGame.strName
        fraType.Caption = .GetINIMessage("Extras Properties Window", "Type", "&Type")
        lblPrompt(1).Caption = .GetINIMessage("Extras Properties Window", "SelectedFile", "Selected File") & ":"
        lblPrompt(4).Caption = .GetINIMessage("Extras Properties Window", "RunnableFile", "Runnable file in archive") & ":"
        lblNormal.Caption = .GetINIMessage("Extras Properties Window", "StandardDesc", "Uses File Associations to run the Extra. For example, if the Extra was a text file, it might be loaded into notepad.exe.")
        lblGemus.Caption = .GetINIMessage("Extras Properties Window", "GemusDesc", "Uses GEMUS to run the Extra. The Extra is treated as though it is a Game File.")
        lblMusic.Caption = .GetINIMessage("Extras Properties Window", "MusicDesc", "Uses a Music Player to run the Extra. Treats the Extra as a Music File.")
        lblURL.Caption = .GetINIMessage("Extras Properties Window", "URLDesc", "The Extra is treated as a URL (web link) and is opened in a web browser.")
        'with shortcuts
        cmdChange.Caption = .GetINIMessage("Extras Properties Window", "ChangeButton", "Ch&ange")
        lblPrompt(0).Caption = .GetINIMessage("Extras Properties Window", "NamePrompt", "&Name") & ":"
        fraType.Caption = .GetINIMessage("Extras Properties Window", "Type", "&Type")
        fraFile.Caption = .GetINIMessage("Extras Properties Window", "File", "&File")
        fraGemus.Caption = .GetINIMessage("Extras Properties Window", "GemusFields", "G&EMUS Fields")
        optType(0).Caption = .GetINIMessage("Extras Properties Window", "StandardOption", "&STANDARD")
        optType(1).Caption = .GetINIMessage("Extras Properties Window", "GemusOption", "&GEMUS")
        optType(2).Caption = .GetINIMessage("Extras Properties Window", "MusicOption", "&MUSIC")
        optType(3).Caption = .GetINIMessage("Extras Properties Window", "URLOption", "&URL")
        lblPrompt(2).Caption = .GetINIMessage("Extras Properties Window", "ImageIndex", "&Image Index") & ":"
        lblPrompt(3).Caption = .GetINIMessage("Extras Properties Window", "URLLabel", "UR&L") & ":"
        optGemus(0).Caption = .GetINIMessage("Extras Properties Window", "UseGemusDefaultFields", "Use GEMUS &Default Fields")
        optGemus(1).Caption = .GetINIMessage("Extras Properties Window", "UseCurrentGameFields", "Use Cu&rrent Game Fields")
        optGemus(2).Caption = .GetINIMessage("Extras Properties Window", "UseTheseFields", "Use T&hese Fields") & "..."
        cmdEdit.Caption = .GetINIMessage("Extras Properties Window", "EditButton", "Edi&t") & "..."
    End With

    'fill the default values
    m_blnFirstTimeItemClick = True
    m_blnByCode = True
    With g_udtExtrasProps
        txtName = .strName
        txtImageIndex = .lngGF_ImageIndex
        lblRunnable.Caption = .strRunnableFile
        optType(.lngType).Value = True
        optGemus(.lngGemusOption).Value = True
        m_udtGemusFields.blnTDE = .blnGF_TDE
        m_udtGemusFields.lngControl = .lngGF_Control
        m_udtGemusFields.lngPalNTSC = .lngGF_PalNTSC
        m_udtGemusFields.lngPlayers = .lngGF_Players
        m_udtGemusFields.strComment = .strGF_Comment
        m_udtGemusFields.strNotes = .strGF_Notes
        m_udtGemusFields.strVComment = .strGF_VComment
        'set the paths listbox
        m_blnLoadBusy = True
        For lngIdx = 0 To (g_lngTotalExtrasPaths - 1)
            If ga_udtRegPathExtras(lngIdx).ePathType <> e_PathLocation_Invalid Then
                lstPaths.AddItem ga_udtRegPathExtras(lngIdx).strRealPath
                lstPaths.ItemData(lstPaths.NewIndex) = lngIdx
                If lngIdx = .lngPathIdx Then
                    lstPaths.ListIndex = lstPaths.NewIndex
                End If
            End If
        Next lngIdx
        'init the dirbox
        DirBox.Path = .strInitPath
        If UCase$(DirBox.Path) <> UCase$(.strInitPath) Then
            FileBox.Path = DirBox.Path
        End If
        m_blnLoadBusy = False
        'select the file?
        If .lngType < 3 Then
            If .strFile <> "" Then
                strFullFile = FileBox.Path
                If Right$(strFullFile, 1) <> "\" Then strFullFile = strFullFile & "\"
                strFullFile = strFullFile & .strFile
                lblFile = GetSubPathFromFullPath(strFullFile)
                'try to find the file
                strFile = UCase(.strFile)
                With FileBox
                    For lngIdx = 0 To (.ListCount - 1)
                        If UCase$(.List(lngIdx)) = strFile Then
                            .ListIndex = lngIdx
                            Exit For
                        End If
                    Next lngIdx
                End With
            End If
        Else
            txtURL.Text = .strFile
        End If
    End With

    'done!
    Call SetOKButton
    g_udtExtrasProps.strName = "" 'means edit was cancelled
    m_blnByCode = False

End Sub

Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
    
        'width
        With fraFile
            .Width = ScaleWidth - (.Left + 120)
            txtURL.Width = .Width - 120
            lstPaths.Width = .Width - 240
        End With
        With lstPaths
            DirBox.Width = .Width
            FileBox.Width = .Width
            lblPrompt(1).Width = .Width
            lblPrompt(4).Width = .Width
            lblFile.Width = .Width
            lblRunnable.Width = .Width - cmdChange.Width
        End With
        cmdChange.Left = lblRunnable.Left + lblRunnable.Width
        With cmdCancel
            .Left = ScaleWidth - (120 + .Width)
            cmdOK.Left = .Left - (120 + cmdOK.Width)
        End With

        'height
        With cmdCancel
            .Top = ScaleHeight - (120 + .Height)
            cmdOK.Top = .Top
            fraFile.Height = .Top - (240 + fraFile.Top)
        End With
        With lblRunnable
            .Top = fraFile.Height - (120 + .Height)
            cmdChange.Top = .Top
            lblPrompt(4).Top = .Top - 240
        End With
        With lblFile
            .Top = lblPrompt(4).Top - 360
            lblPrompt(1).Top = .Top - 240
            FileBox.Height = .Top - (300 + FileBox.Top)
        End With

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnExtrasSubclassed = True Then
            Call ExtrasUnSubClass(Me.hwnd)
        End If
    End If

    'save reg settings
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "ExtrasMaximized", True, False
        End If
    End With

End Sub


Private Sub lstPaths_Click()

    If m_blnByCode = False Then
        With lstPaths
            If .ListIndex > -1 Then
                DirBox.Path = .List(.ListIndex)
            End If
        End With
    End If

End Sub


Private Sub optGemus_Click(Index As Integer)
    
    cmdEdit.Enabled = ((Index = 2) And optType(1).Value = True)

End Sub

Private Sub optType_Click(Index As Integer)

    Dim lngListIndex As Long

    fraGemus.Enabled = (Index = 1)
    lblPrompt(2).Enabled = (Index = 1)
    txtImageIndex.Enabled = (Index = 1)
    txtImageIndex.BackColor = IIf(Index = 1, txtName.BackColor, BackColor)
    optGemus(0).Enabled = (Index = 1)
    optGemus(1).Enabled = (Index = 1)
    optGemus(2).Enabled = (Index = 1)
    cmdEdit.Enabled = (Index = 1 And optGemus(2).Value = True)
    Line1.BorderColor = IIf(Index = 1, lblPrompt(0).ForeColor, lblDummy.BackColor)
    fraFile.Enabled = (Index <> 3)
    fraFile.Visible = fraFile.Enabled
    txtURL.Visible = (Index = 3)
    lblPrompt(3).Visible = (Index = 3)
    lblPrompt(4).Visible = (Index = 1)
    lblRunnable.Visible = (Index = 1)
    lblRunnable.Enabled = False
    cmdChange.Visible = (Index = 1)
    cmdChange.Enabled = False
    If Not m_blnByCode And Index < 3 And FileBox.Enabled = True And FileBox.Visible = True Then
        lngListIndex = FileBox.ListIndex
        FileBox.ListIndex = -1
        FileBox.ListIndex = lngListIndex
    End If
    Call SetOKButton

End Sub


Private Sub txtImageIndex_Change()

    Call SetOKButton

End Sub


Private Sub txtImageIndex_GotFocus()

    With txtImageIndex
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtImageIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If

    If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
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


Private Sub txtURL_Change()

    Call SetOKButton

End Sub


Private Sub txtURL_GotFocus()

    With txtURL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


