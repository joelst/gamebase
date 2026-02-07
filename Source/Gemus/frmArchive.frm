VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchive 
   Caption         =   "Select Runnable File from Archive..."
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frmArchive.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ilstArchive 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchive.frx":08CA
            Key             =   "archive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchive.frx":0962
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchive.frx":0AC5
            Key             =   "runnable"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchive.frx":0C20
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvArchive 
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   688
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilstArchive"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTick 
      Caption         =   "RE-PACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgTick 
      Height          =   360
      Left            =   120
      Picture         =   "frmArchive.frx":1036
      Top             =   3000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblDummy 
      Caption         =   "Label1"
      Enabled         =   0   'False
      ForeColor       =   &H80000010&
      Height          =   135
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmArchive"
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


Private Unpacker As clsUnpack
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_xobjRegistry As GBFuncs.clsRegistry
Private m_strAppPathWithBackSlash As String
Private m_strGBName As String
Private m_strMessFile As String
Private m_strFileTypes As String
Private m_strArchiveFileTypes As String
Private m_strTextRunnable As String
Private m_strTextExtracting As String
Private m_lngNextFreeExtractionFolderIdx As Long
Private m_str7ZipExtractorPathFile As String
Private m_strGemusINIFilename As String
Private m_strTextSelected As String
Private m_strTextMainArchive As String

Private Sub AddParentArchiveFolderIdxToAllImmediateSubArchiveTags(strArchiveKey As String, strFromKey As String, ByVal lngFolderIdx As Long)

    Dim n As node
    Dim lngIdx As Long
    Dim nFind As node
    Dim blnMatch As Boolean

    Set n = tvArchive.Nodes(strFromKey)
    
    If n.Children = 0 Then Exit Sub
    Set n = n.Child.FirstSibling
    For lngIdx = 1 To tvArchive.Nodes(strFromKey).Children
        If n.Children > 0 Then
            Call AddParentArchiveFolderIdxToAllImmediateSubArchiveTags( _
                    strArchiveKey, _
                    n.Key, _
                    lngFolderIdx _
                )
        End If
        If Right$(n.Key, 1) = ":" Then
            'node is an archive, so go trace up to the first archive node we find
            'and add the folder index to the tag if it matches strArchiveKey
            blnMatch = True
            Set nFind = n.Parent
            Do While nFind.Key <> strArchiveKey
                If Right$(nFind.Key, 1) = ":" Then
                    blnMatch = False
                    Exit Do
                End If
                Set nFind = nFind.Parent
            Loop
            If blnMatch Then
                Call AddToTag(n.Key, Trim$(lngFolderIdx))
            End If
        End If
        'move to next sibling
        Set n = n.Next
    Next lngIdx
    
    Set n = Nothing

End Sub

Private Sub AddToTag(strNodeKey As String, strValue As String)

    If tvArchive.Nodes(strNodeKey).Tag = "" Then
        tvArchive.Nodes(strNodeKey).Tag = tvArchive.Nodes(strNodeKey).Tag & ","
    End If

    tvArchive.Nodes(strNodeKey).Tag = tvArchive.Nodes(strNodeKey).Tag & strValue & ","

End Sub

Private Function CheckTag(strTag As String, strFind As String) As Boolean

    CheckTag = (InStr(strTag, "," & strFind & ",") > 0)

End Function
Private Function GetCurrentFileFromArchiveKey(strKey As String) As String

    Dim strFile As String
    Dim a_strCrumb() As String

    strFile = Left$(strKey, Len(strKey) - 1) 'remove trailing ":"
    a_strCrumb = Split(strFile, ":")
    GetCurrentFileFromArchiveKey = a_strCrumb(UBound(a_strCrumb))

End Function

Private Function GetFileKindFromFilenameExtension(strExt As String, blnIsRunnable_out As Boolean) As Long

    Dim strUpperExt As String
    Dim blnIsArchive As Boolean
    
    strUpperExt = ";" & UCase$(strExt) & ";"
    blnIsRunnable_out = (InStr(";" & m_strFileTypes & ";", strUpperExt) > 0)
    blnIsArchive = (InStr(";" & m_strArchiveFileTypes & ";", strUpperExt) > 0)
    
    If blnIsArchive Then
        GetFileKindFromFilenameExtension = 2 'archive (maybe runnable)
    ElseIf blnIsRunnable_out Then
        GetFileKindFromFilenameExtension = 1 'runnable file
    Else
        GetFileKindFromFilenameExtension = 0 'unrunnable file
    End If

End Function

Private Function GetFolderIdxFromTag(strTag As String) As Long

    Dim a_strTags() As String
    Dim lngNumTags As Long
    Dim lngIdx As Long
    Dim lngCharIdx As Long
    Dim strCurrentTag As String
    Dim blnInvalid As Boolean

    GetFolderIdxFromTag = 0
    If strTag = "" Or strTag = "," Or strTag = ",," Then Exit Function

    a_strTags = Split(strTag, ",")
    lngNumTags = UBound(a_strTags) + 1
    For lngIdx = 0 To lngNumTags - 1
        strCurrentTag = a_strTags(lngIdx)
        If strCurrentTag <> "" Then
            blnInvalid = False
            For lngCharIdx = 1 To Len(strCurrentTag)
                If InStr("0123456789", Mid$(strCurrentTag, lngCharIdx, 1)) = 0 Then
                    blnInvalid = True
                    Exit For
                End If
            Next lngCharIdx
            If blnInvalid = False Then
                GetFolderIdxFromTag = strCurrentTag
                Exit Function
            End If
        End If
    Next lngIdx

End Function

Private Sub PopulateArchiveListUnderNode(strParentNodeKey As String, ByVal lngNumFiles As Long, a_strFiles() As String)

    Dim lngIdx As Long
    Dim a_strParts() As String
    Dim lngNumParts As Long
    Dim node As node
    Dim strFile As String
    Dim strExt As String
    Dim lngDotPos As Long
    Dim lngLen As Long
    Dim lngFileKind As Long
    Dim lngPartIdx As Long
    Dim strKey As String
    Dim strParentKey As String
    Dim lngTempIdx As Long
    Dim a_strJoined() As String
    Dim a_strJoinedParent() As String
    Dim blnIsRunnable As Boolean

    For lngIdx = 0 To lngNumFiles - 1
    
        a_strParts = Split(a_strFiles(lngIdx), "\")
        lngNumParts = UBound(a_strParts) + 1
        
        For lngPartIdx = 0 To lngNumParts - 1
        
            strFile = a_strParts(lngPartIdx)
        
            'its a sub folder
            
            ReDim a_strJoined(lngPartIdx)
            For lngTempIdx = 0 To lngPartIdx
                a_strJoined(lngTempIdx) = a_strParts(lngTempIdx)
            Next lngTempIdx
            strKey = strParentNodeKey & Join(a_strJoined, "\")
            
            If lngPartIdx > 0 Then
                ReDim a_strJoinedParent(lngPartIdx - 1)
                For lngTempIdx = 0 To lngPartIdx - 1
                    a_strJoinedParent(lngTempIdx) = a_strParts(lngTempIdx)
                Next lngTempIdx
                strParentKey = strParentNodeKey & Join(a_strJoinedParent, "\")
            Else
                strParentKey = strParentNodeKey
            End If

            If Not NodeExists(strKey) Then
                If lngPartIdx < lngNumParts - 1 Then
                    'its a folder
                    Set node = tvArchive.Nodes.Add(strParentKey, tvwChild, strKey, strFile, "folder", "folder")
                    node.ForeColor = lblDummy.ForeColor
                Else
                    'its a file
                    lngLen = Len(strFile)
                    lngDotPos = InStrRev(strFile, ".")
                    If lngDotPos < lngLen Then
                        strExt = LCase$(Right$(strFile, lngLen - lngDotPos))
                        lngFileKind = GetFileKindFromFilenameExtension(strExt, blnIsRunnable)
                    Else
                        strExt = ""
                        lngFileKind = 0 'non-runnable file
                    End If
                    If lngFileKind = 0 Then
                        'unrunnable
                        Set node = tvArchive.Nodes.Add(strParentKey, tvwChild, strKey, strFile, "file", "file")
                        node.ForeColor = lblDummy.ForeColor
                    ElseIf lngFileKind = 1 Then
                        'runnable
                        Set node = tvArchive.Nodes.Add(strParentKey, tvwChild, strKey, strFile, "runnable", "runnable")
                        With node
                            .Bold = True
                            .ForeColor = &H8000&
                            .Tag = ",runnable,"
                        End With
                    Else
                        'archive
                        Set node = tvArchive.Nodes.Add(strParentKey, tvwChild, strKey & ":", strFile, "archive", "archive")
                        With node
                            .Tag = ",load,extract,"
                            If blnIsRunnable Then
                                .Tag = node.Tag & "runnable,"
                                .Bold = True
                                .ForeColor = &H8000&
                            Else
                                .ForeColor = &HFF0000
                            End If
                        End With
                        Set node = Nothing
                        Set node = tvArchive.Nodes.Add(strKey & ":", tvwChild, strKey & ":__dummy")
                    End If
                    Set node = Nothing
                End If
            End If

        Next lngPartIdx
    
    Next lngIdx

End Sub
Private Function NodeExists(strKey As String) As Boolean

    On Error GoTo Err_Handler

    Dim n As node
    Set n = tvArchive.Nodes(strKey)
    Set n = Nothing
    NodeExists = True
    Exit Function

Err_Handler:
    NodeExists = False

End Function

Private Sub RemoveFromTag(strNodeKey As String, strFind As String)

    Dim strNewTag As String
    strNewTag = Replace(tvArchive.Nodes(strNodeKey).Tag, "," & strFind & ",", ",", 1, 1)
    If strNewTag = "," Then strNewTag = ""
    tvArchive.Nodes(strNodeKey).Tag = strNewTag

End Sub
Private Sub SBU(strText As String)

    sbStatus.Panels(2).Text = strText
    sbStatus.Panels(2).ToolTipText = strText
    DoEvents

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    With g_udtArchiveBrowse
        .blnOKPressed = True
        .strRunnableFile = Right$(tvArchive.SelectedItem.Key, Len(tvArchive.SelectedItem.Key) - Len("ROOT:"))
        If Right$(.strRunnableFile, 1) = ":" Then .strRunnableFile = Left$(.strRunnableFile, Len(.strRunnableFile) - 1)
    End With

    Unload Me

End Sub


Private Sub Form_Load()

    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips
    Const DesignScaleHeight As Single = 3735    'The required minimum scaleheight for the form
    Const DesignScaleWidth As Single = 6255     'The required minimum scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim blnRegMaximized As Boolean              'last set window maximized state
    Dim strError As String
    Dim a_strTopLevelFiles() As String
    Dim lngTotalTopLevelFiles As Long
    Dim strArchiveFile As String
    Dim strArchivePath As String
    Dim node As node
    Dim lngLen As Long
    Dim lngDotPos As Long
    Dim strExt As String
    Dim lngFileKind As Long
    Dim blnIsRunnable As Boolean

    Set m_xobjFunc = New GBFuncs.clsFunctions
    Set m_xobjRegistry = New GBFuncs.clsRegistry

    m_strAppPathWithBackSlash = App.Path
    If g_blnInVBIDE Then
        m_strAppPathWithBackSlash = gc_strAPP_PATH_WBS_WHEN_IN_VBIDE
    End If
    If Right$(m_strAppPathWithBackSlash, 1) <> "\" Then m_strAppPathWithBackSlash = m_strAppPathWithBackSlash & "\"
    m_strMessFile = m_strAppPathWithBackSlash & "GemusText.ini"
    
    m_lngNextFreeExtractionFolderIdx = 1
    m_strGemusINIFilename = "__gemus.ini"

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngArchiveWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngArchiveWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngArchiveWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngArchiveWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With m_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ArchiveWidth", sngRegWidth, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ArchiveHeight", sngRegHeight, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ArchiveTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ArchiveLeft", sngRegLeft, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\Gemus\Window", "ArchiveMaximized", blnRegMaximized, False, False
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
        Call ArchiveSubClass(Me.hWnd)
    End If

    'get ini text
    Caption = GetINIMessageFromFile("Archive Browser", "Caption", "Select Runnable File from Archive", m_strMessFile) & "..."
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", m_strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", m_strMessFile)
    m_strTextRunnable = GetINIMessageFromFile("Archive Browser", "Runnable", "Runnable", m_strMessFile) & ": "
    m_strTextExtracting = GetINIMessageFromFile("Archive Browser", "Extracting", "Extracting", m_strMessFile) & " "
    m_strTextSelected = GetINIMessageFromFile("Archive Browser", "Selected", "Selected", m_strMessFile) & " "
    m_strTextMainArchive = GetINIMessageFromFile("Archive Browser", "MainArchive", "Main Archive", m_strMessFile) & " "

    Set Unpacker = New clsUnpack
    m_strArchiveFileTypes = UCase$(Unpacker.GetUnpackableFileTypeString())
    With g_udtArchiveBrowse
        Call m_xobjFunc.GetFileAndPath(.strArchivePathFile, strArchivePath, strArchiveFile)
        m_strFileTypes = UCase$(Join(.a_strRunnableFileTypes, ";"))
        strError = Unpacker.SevenZipList(.strArchivePathFile, .strTempPathWBS & "list.txt", lngTotalTopLevelFiles, a_strTopLevelFiles)
    End With
    If strError <> "" Then
        MsgBox strError, vbExclamation Or vbOKOnly, "GEMUS"
        cmdOK.Enabled = False
        cmdCancel.Enabled = True
        tvArchive.Enabled = False
        sbStatus.Panels(1).Text = strError
        Exit Sub
    End If
    m_str7ZipExtractorPathFile = Unpacker.GetExtractorPath("7z.exe")
    If m_str7ZipExtractorPathFile = "" Then
        strError = GetINIMessageFromFile("Archive Browser", "Error_No7z", "Unable to find 7z.exe", m_strMessFile)
        MsgBox strError, vbExclamation Or vbOKOnly, "GEMUS"
        cmdOK.Enabled = False
        cmdCancel.Enabled = True
        tvArchive.Enabled = False
        sbStatus.Panels(1).Text = strError
        Exit Sub
    End If

    'if here, fill the treeview with the archive's top level file list
    
    sbStatus.Panels(1).Text = m_strFileTypes
    sbStatus.Panels(1).ToolTipText = m_strFileTypes

    lngLen = Len(g_udtArchiveBrowse.strArchivePathFile)
    lngDotPos = InStrRev(g_udtArchiveBrowse.strArchivePathFile, ".")
    If lngDotPos < lngLen Then
        strExt = LCase$(Right$(g_udtArchiveBrowse.strArchivePathFile, lngLen - lngDotPos))
        lngFileKind = GetFileKindFromFilenameExtension(strExt, blnIsRunnable)
    End If
    With tvArchive
        .Nodes.Clear
        Set node = .Nodes.Add(, , "ROOT:", g_udtArchiveBrowse.strArchivePathFile, "archive", "archive")
        With node
            .Bold = True
            .Expanded = True
            .Tag = ",extract,"
            If blnIsRunnable Then
                .Tag = .Tag & "runnable,"
                .ForeColor = &H8000&
            End If
        End With
        Set node = Nothing
    End With
    Call PopulateArchiveListUnderNode("ROOT:", lngTotalTopLevelFiles, a_strTopLevelFiles)
    Call SBU(m_strTextMainArchive)
    cmdOK.Enabled = blnIsRunnable

End Sub

Private Sub Form_Resize()

    'width
    sbStatus.Panels(1).Width = ScaleWidth / 2
    sbStatus.Panels(2).Width = ScaleWidth / 2
    cmdCancel.Left = (ScaleWidth - 120) - cmdCancel.Width
    cmdOK.Left = cmdCancel.Left - 120 - cmdOK.Width
    tvArchive.Width = ScaleWidth
    lblTick.Left = imgTick.Left + imgTick.Width + 120

    'height
    cmdOK.Top = (ScaleHeight - sbStatus.Height) - 480
    cmdCancel.Top = cmdOK.Top
    tvArchive.Height = cmdOK.Top - 120
    imgTick.Top = tvArchive.Height + 120
    lblTick.Top = imgTick.Top

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnArchiveSubclassed = True Then
            Call ArchiveUnSubClass(Me.hWnd)
        End If
    End If

    'save window sizes etc...
    With m_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\Gemus\Window", "ArchiveMaximized", True, False
        End If
    End With

    Set m_xobjFunc = Nothing
    Set m_xobjRegistry = Nothing
    Set Unpacker = Nothing

End Sub


Private Sub tvArchive_Expand(ByVal node As MSComctlLib.node)

    On Error GoTo Err_Handler

    Dim strCrumb As String
    Dim a_strCrumb() As String
    Dim lngIdx As Long
    Dim strDelim As String
    Dim lngNumCrumbElements As Long
    Dim strArchiveKey As String
    Dim strArchivePathFile As String
    Dim strArchivePath As String
    Dim strArchiveFile As String
    Dim lngParentArchiveFolderIdx As Long
    Dim strError As String
    Dim lngNumFiles As Long
    Dim a_strFiles() As String
    Dim strFile As String
    Dim strOriginalNodeKey As String
    Dim lngFolderIdx As Long
    Dim lngTempFolderIdx As Long
    Dim strParentArchiveKey As String
    Dim blnTemp As Boolean
    Dim strTempSB As String
    Dim blnWasExtracted As Boolean

    'only interested in archives we need to load the list for
    If CheckTag(node.Tag, "load") = False Then Exit Sub

    'get a crumb back to the root node (main archive)
    strOriginalNodeKey = node.Key
    strDelim = "***___***"
    strCrumb = ""
    Do While node.Key <> "ROOT:"
        If Right$(node.Key, 1) = ":" Then
            'its an archive
            strCrumb = node.Key & strDelim & strCrumb
        End If
        Set node = node.Parent
    Loop
    strCrumb = "ROOT:" & strDelim & strCrumb
    strCrumb = Left$(strCrumb, Len(strCrumb) - Len(strDelim))
    a_strCrumb = Split(strCrumb, strDelim)
    lngNumCrumbElements = UBound(a_strCrumb) + 1

    'loop through each archive from root to this archive's parent archive, extracting
    'each one if necessary, so we can get to this archive and list its contents
    lngParentArchiveFolderIdx = 0
    lngFolderIdx = GetFolderIdxFromTag(tvArchive.Nodes(strOriginalNodeKey).Tag)
    If lngFolderIdx = 0 Then
        For lngIdx = 0 To lngNumCrumbElements - 2 'ignore the current archive, don't need to extract that to list its contents
            strArchiveKey = a_strCrumb(lngIdx)
            If CheckTag(tvArchive.Nodes(strArchiveKey).Tag, "extract") = True Then
            
                '************************************ need to extract this archive ************************************
    
                If lngIdx > 0 Then
                    'its a sub-archive (parent will have already been extracted so get path\file from there
                    lngTempFolderIdx = GetFolderIdxFromTag(tvArchive.Nodes(strArchiveKey).Tag)
                    strArchivePathFile = g_udtArchiveBrowse.strTempPathWBS _
                        & lngTempFolderIdx & "\" _
                        & GetCurrentFileFromArchiveKey(strArchiveKey)
                Else
                    'its the main archive
                    strArchivePathFile = g_udtArchiveBrowse.strArchivePathFile
                End If
                Call m_xobjFunc.GetFileAndPath(strArchivePathFile, strArchivePath, strArchiveFile)
                'extract it
                Screen.MousePointer = vbHourglass
                strTempSB = sbStatus.Panels(2).Text
                Call SBU("Extracting " & strArchiveFile & "...")
                blnTemp = cmdOK.Enabled
                cmdOK.Enabled = False
                tvArchive.Enabled = False
                Call m_xobjFunc.DeletePathSilent(g_udtArchiveBrowse.strTempPathWBS & m_lngNextFreeExtractionFolderIdx)
                blnWasExtracted = Unpacker.SevenZipUnpack( _
                        m_str7ZipExtractorPathFile, _
                        strArchivePath, _
                        strArchiveFile, _
                        g_udtArchiveBrowse.strTempPathWBS & m_lngNextFreeExtractionFolderIdx _
                    )
                cmdOK.Enabled = blnTemp
                tvArchive.Enabled = True
                Call SBU(strTempSB)
                Screen.MousePointer = vbNormal
                If Not blnWasExtracted Then
                    MsgBox GetINIMessageFromFile("Archive Browser", "Error_Extract", "Unable to extract this archive", m_strMessFile) _
                            & ":" & vbCrLf & vbCrLf _
                            & strArchivePathFile, _
                        vbExclamation Or vbOKOnly, "GEMUS"
                    Exit Sub
                End If
                'this archive is now extracted
                Call RemoveFromTag(tvArchive.Nodes(strArchiveKey).Key, "extract")
                'store the parent's data if its the last one
                If lngIdx = lngNumCrumbElements - 2 Then
                    lngParentArchiveFolderIdx = m_lngNextFreeExtractionFolderIdx
                    strParentArchiveKey = strArchiveKey
                End If
                'increment next free folder idx for extracting to
                m_lngNextFreeExtractionFolderIdx = m_lngNextFreeExtractionFolderIdx + 1
            End If
        Next lngIdx
    End If
    
    '************ get the contents list for the requested archive and add to the treeview ************
    
    'for all child nodes of the same parent node, add the folder index to the tag
    If lngParentArchiveFolderIdx > 0 Then
        Call AddParentArchiveFolderIdxToAllImmediateSubArchiveTags( _
                strParentArchiveKey, _
                strParentArchiveKey, _
                lngParentArchiveFolderIdx _
            )
        lngFolderIdx = lngParentArchiveFolderIdx
    End If

    'add the contents to the treeview
    strFile = g_udtArchiveBrowse.strTempPathWBS _
        & lngFolderIdx & "\" _
        & GetCurrentFileFromArchiveKey(strOriginalNodeKey)
    strError = Unpacker.SevenZipList(strFile, g_udtArchiveBrowse.strTempPathWBS & "list.txt", lngNumFiles, a_strFiles)
    If strError = "" Then
        Call PopulateArchiveListUnderNode(strOriginalNodeKey, lngNumFiles, a_strFiles)
    Else
        MsgBox GetINIMessageFromFile("Archive Browser", "Error_List", "Unable to list the contents of this archive", m_strMessFile) _
                & ":" & vbCrLf & vbCrLf _
                & strFile, _
            vbExclamation Or vbOKOnly, "GEMUS"
        Exit Sub
    End If

    'remove the dummy node
    tvArchive.Nodes.Remove tvArchive.Nodes(strOriginalNodeKey).Child.FirstSibling.Index

    'no longer need to load this archive's contents list
    Call RemoveFromTag(strOriginalNodeKey, "load")

    Exit Sub

Err_Handler:
    MsgBox GetINIMessageFromFile("Archive Browser", "Error_ExpandNode", "Unable to expand this node!", m_strMessFile) _
            & vbCrLf & vbCrLf _
            & Err.Description & " (Err " & Trim$(Err.Number) & "," & Trim$(Err.LastDllError) & ")", _
        vbExclamation Or vbOKOnly, "GEMUS"

End Sub
Private Sub tvArchive_NodeClick(ByVal node As MSComctlLib.node)
    
    'Debug.Print node.Key & " [" & node.Tag & "]"
    
    Dim a_strKeys() As String
    Dim lngNumKeys As Long
    Dim lngIdx As Long
    Dim blnOK As Boolean
    Dim strRunnableFile As String

    cmdOK.Enabled = CheckTag(node.Tag, "runnable")

    If CheckTag(node.Tag, "runnable") = False Then
        'only show re-pack notice if a runnable file is selected
        imgTick.Visible = False
        lblTick.Visible = False
        Call SBU("")
    Else
        If node.Key <> "ROOT:" Then
            strRunnableFile = Right(node.Key, Len(node.Key) - Len("ROOT:"))
            If Right$(strRunnableFile, 1) = ":" Then strRunnableFile = Left$(strRunnableFile, Len(strRunnableFile) - 1)
            strRunnableFile = Replace(strRunnableFile, ":", ": ")
            Call SBU(strRunnableFile)
        Else
            Call SBU(m_strTextMainArchive)
        End If

        'check main archive is re-packable
        If Not Unpacker.IsFileARepackableArchiveFileType(g_udtArchiveBrowse.strArchivePathFile) Then
            imgTick.Visible = False
            lblTick.Visible = False
        Else
            'check all sub archives are re-packable
            a_strKeys = Split(node.Key, ":")
            lngNumKeys = UBound(a_strKeys) + 1
            blnOK = True
            For lngIdx = 1 To lngNumKeys - 2
                If Not Unpacker.IsFileARepackableArchiveFileType(a_strKeys(lngIdx)) Then
                    blnOK = False
                    Exit Sub
                End If
            Next lngIdx
        End If
    End If
    imgTick.Visible = blnOK
    lblTick.Visible = blnOK

End Sub


