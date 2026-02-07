VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameSelector 
   Caption         =   "Game Selector"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmGameSelector.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange"
      Height          =   285
      Left            =   9000
      TabIndex        =   23
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtVer 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   7
      Top             =   6000
      Width           =   375
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Choose &GEMUS script parameters"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "&Apply to all games in current view"
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   5760
      Width           =   1575
   End
   Begin VB.ComboBox lstFilter 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Cl&ear"
      Height          =   285
      Left            =   9000
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3975
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   3975
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Game Files"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.DirListBox DirBox 
      Height          =   2565
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.ComboBox lstGamePath 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvContents 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "&Ver:"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label lblDefaultFile 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblIndex 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblImage 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected File:"
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
      Index           =   1
      Left            =   5040
      TabIndex        =   18
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label lblFile 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   4800
      UseMnemonic     =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblFileInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Info 2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblFileInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Info 1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      UseMnemonic     =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Runnable file in archive:"
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
      Index           =   0
      Left            =   5040
      TabIndex        =   14
      Top             =   5160
      UseMnemonic     =   0   'False
      Width           =   4815
   End
   Begin VB.Menu pmnuEmulators 
      Caption         =   "Popup Emulators"
      Visible         =   0   'False
      Begin VB.Menu pmnuEmulator 
         Caption         =   "Default Emulator"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmGameSelector"
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

Private m_strLastGamePath As String
Private m_blnByCode As Boolean
Private m_lngGPIdx As Long
Private m_strDefaultFile As String
Private m_strInvalidFile As String
Private m_strSize As String
Private m_strImages As String
Private m_strBlocks As String
Private m_strImageName As String
Private m_strType As String
Private m_strFile As String
Private m_strBytes As String
Private m_strLoadAddress As String
Private m_strVersion As String
Private m_strEntries As String
Private m_strStatusOK As String
Private m_strStatusNoOpenFile As String
Private m_strStatusInvalidFile As String
Private ma_strCartTypes(13) As String
Private m_strUnrecCart As String
Private m_strInactive As String
Private m_strActive As String
Private m_strExromLine As String
Private m_strGameLine As String
Private m_strTestGameFile As String
Private m_blnLoadBusy As Boolean
Private m_blnFirstTimeItemClick As Boolean

Private m_strTablePETSCII As String

Private Sub MakeTablePETSCII()
    
    Dim Char As String
    Dim i As Integer

    m_strTablePETSCII = " "
  
    For i = 0 To 25
        Char = Chr(Asc("A") + i)
        m_strTablePETSCII = m_strTablePETSCII & Char
    Next i
    m_strTablePETSCII = m_strTablePETSCII & "[£]^-"
    For i = 0 To 58
        Char = Chr(Asc(" ") + i)
        m_strTablePETSCII = m_strTablePETSCII & Char
    Next i
    m_strTablePETSCII = m_strTablePETSCII & "[£]^-"
    For i = 0 To 96
        Char = Chr(Asc("-"))
        m_strTablePETSCII = m_strTablePETSCII & Char
    Next i
    For i = 0 To 25
        Char = Chr(Asc("A") + i)
        m_strTablePETSCII = m_strTablePETSCII & Char
    Next i
    For i = 0 To 39
        Char = Chr(Asc("+"))
        m_strTablePETSCII = m_strTablePETSCII & Char
    Next i
    
End Sub

Private Function ConvertPETSCII(strInput As String) As String
    
    Dim i, Length As Integer
    Dim Char As String
    
    ConvertPETSCII = ""
    Length = Len(strInput)
    If Length > 0 Then
        Length = Length - 1
    End If

    For i = 0 To Length
        Char = Mid$(strInput, i + 1, 1)
        Char = Mid$(m_strTablePETSCII, Asc(Char) + 1, 1)
        ConvertPETSCII = ConvertPETSCII & Char
    Next i

End Function


Private Sub LoadEmulatorPopupMenu()

    '------------------------------------------------
    'loads the popup menu for the emulators selection
    '------------------------------------------------

    Dim lngMenuIdx As Long      'loop index

    'delete all current menus (if any)
    For lngMenuIdx = 1 To pmnuEmulator.Count - 1
        Unload pmnuEmulator(lngMenuIdx)
    Next lngMenuIdx

    If g_xobjGEMUS.TotalEmulators > 0 Then
        'load any new ones
        For lngMenuIdx = 1 To (g_xobjGEMUS.TotalEmulators - 1)
            Load pmnuEmulator(lngMenuIdx)
        Next lngMenuIdx
        'set their properties
        For lngMenuIdx = 0 To (g_xobjGEMUS.TotalEmulators - 1)
            With pmnuEmulator(lngMenuIdx)
                .Caption = g_xobjGEMUS.EmulatorName(lngMenuIdx)
                .Enabled = True
                .Visible = True
            End With
        Next lngMenuIdx
    Else
        'no emulators available
        pmnuEmulator(0).Enabled = False
        pmnuEmulator(0).Caption = frmMain.cmdPlay.Caption
    End If

End Sub
Private Function DisplayFile(strFile_in As String, strExt_in As String) As String

    MousePointer = vbHourglass
    DoEvents
    Select Case strExt_in
    
        Case "T64"
            DisplayFile = DisplayT64(strFile_in, strExt_in)

        Case "D64"
            DisplayFile = DisplayD64(strFile_in, strExt_in)
        
        Case "TAP"
            DisplayFile = DisplayTAP(strFile_in, strExt_in)
        
        Case "PRG"
            DisplayFile = DisplayPRG(strFile_in, strExt_in)
        
        Case "P00", "P01"
            DisplayFile = DisplayP00(strFile_in, strExt_in)
        
        Case "G64", "G41"
            DisplayFile = DisplayG64(strFile_in, strExt_in)

        Case "CRT"
            DisplayFile = DisplayCRT(strFile_in, strExt_in)

        Case Else
            DisplayFile = DisplayOtherSupported(strFile_in, strExt_in)
    
    End Select
    MousePointer = vbDefault

End Function
Private Function DisplayOtherSupported(strFile_in As String, strExt_in As String) As String

    '---------------------------------------------------
    'display the supported file in the contents listview
    '
    'Returns:   Empty string if ok, error string if not
    '---------------------------------------------------

    Dim lngSize&
    Dim bigSize As Currency

    On Error GoTo Err_Handler

    'clear contents listview
    lvContents.ListItems.Clear
    lvContents.ColumnHeaders.Clear
    lvContents.Enabled = False

    'get file size
    bigSize = g_xobjFunc.GetbigFileSize(strFile_in)

    'set labels and exit
    lblFileInfo(0) = strExt_in & " " & m_strFile
    lngSize& = bigSize / 1024
    lblFileInfo(1) = Format$(lngSize&, "0.00") & " KB"
    DisplayOtherSupported = ""
    lblIndex = ""
    Exit Function

Err_Handler:
    
    DisplayOtherSupported = m_strStatusNoOpenFile & " " & Err.Description

End Function
Private Function DisplayT64(strFile_in As String, strExt_in As String) As String

    Dim T64 As Gemus.clsT64
    Dim T64Entry As Gemus.clsT64Entry
    Dim strErr As String

    'create and load the t64 file
    Set T64 = New Gemus.clsT64
    strErr = T64.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .Visible = False
            .ColumnHeaders.Clear
            Call .ColumnHeaders.Add(, , m_strImages)
            .ColumnHeaders(1).Width = .Width - 500
            For Each T64Entry In T64.Entries
                Call .ListItems.Add(, , ConvertPETSCII(T64Entry.C64Filename))
            Next T64Entry
            If .ListItems.Count > 0 Then
                .Enabled = True
            Else
                .ColumnHeaders.Clear
            End If
            .Visible = True
            lblFileInfo(0) = "T64 " & m_strFile
            lblFileInfo(1) = """" & T64.TapeImageName & """   " & T64.EntriesUsed & "/" & T64.EntriesCount & " " & m_strEntries
        End With
        DisplayT64 = ""
    Else
        DisplayT64 = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayCRT(strFile_in As String, strExt_in As String) As String

    Dim CRT As Gemus.clsCRT
    Dim strType As String
    Dim strErr As String

    'create and load the crt file
    Set CRT = New Gemus.clsCRT
    strErr = CRT.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .Visible = False
            .ColumnHeaders.Clear
            If CRT.HardwareType < 0 Or CRT.HardwareType > 13 Then
                strType = m_strUnrecCart
            Else
                strType = ma_strCartTypes(CRT.HardwareType)
            End If
            Call .ColumnHeaders.Add(, , strType)
            .ColumnHeaders(1).Width = .Width - 500
            Call .ListItems.Add(, , """" & CRT.CartName & """")
            Call .ListItems.Add(, , m_strExromLine & ": " & IIf((CRT.ExromLine = True), m_strActive, m_strInactive))
            Call .ListItems.Add(, , m_strGameLine & ": " & IIf((CRT.GameLine = True), m_strActive, m_strInactive))
            Set .SelectedItem = Nothing
            .ListItems(1).Selected = False
            .ListItems(2).Selected = False
            .ListItems(3).Selected = False
            .Visible = True
            lblFileInfo(0) = "CRT " & m_strFile
            lblFileInfo(1) = CRT.FileSize & " KB  """ & CRT.Signature & """  " & CRT.Version
        End With
        DisplayCRT = ""
    Else
        DisplayCRT = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayPRG(strFile_in As String, strExt_in As String) As String

    Dim PRG As Gemus.clsPRG
    Dim strErr As String

    'create and load the t64 file
    Set PRG = New Gemus.clsPRG
    strErr = PRG.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .ColumnHeaders.Clear
            lblFileInfo(0) = "PRG " & m_strFile
            lblFileInfo(1) = PRG.FileSize & " KB   " & m_strLoadAddress & ": " & PRG.LoadAddress
        End With
        DisplayPRG = ""
    Else
        DisplayPRG = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayTAP(strFile_in As String, strExt_in As String) As String

    Dim TAP As Gemus.clsTAP
    Dim strErr As String

    'create and load the t64 file
    Set TAP = New Gemus.clsTAP
    strErr = TAP.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .ColumnHeaders.Clear
            lblFileInfo(0) = "TAP " & m_strFile
            lblFileInfo(1) = TAP.FileSize & " KB   """ & TAP.Signature & """   " & TAP.Version
        End With
        DisplayTAP = ""
    Else
        DisplayTAP = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayG64(strFile_in As String, strExt_in As String) As String

    Dim G64 As Gemus.clsG64
    Dim strErr As String

    'create and load the t64 file
    Set G64 = New Gemus.clsG64
    strErr = G64.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .ColumnHeaders.Clear
            lblFileInfo(0) = strExt_in & " " & m_strFile
            lblFileInfo(1) = G64.Version & "  """ & G64.Signature & """  (" & G64.NumTracks & " * " & G64.TrackSize & " " & m_strBytes & ")"
        End With
        DisplayG64 = ""
    Else
        DisplayG64 = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayP00(strFile_in As String, strExt_in As String) As String

    Dim P00 As Gemus.clsP00
    Dim strErr As String

    'create and load the t64 file
    Set P00 = New Gemus.clsP00
    strErr = P00.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .ColumnHeaders.Clear
            lblFileInfo(0) = strExt_in & " " & m_strFile
            lblFileInfo(1) = P00.FileSize & " KB   """ & P00.Filename & """"
        End With
        DisplayP00 = ""
    Else
        DisplayP00 = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function
Private Function DisplayD64(strFile_in As String, strExt_in As String) As String

    Dim D64 As Gemus.clsD64
    Dim D64Entry As Gemus.clsD64Entry
    Dim LItem As ListItem
    Dim strErr As String

    'create and load the t64 file
    Set D64 = New Gemus.clsD64
    strErr = D64.Load(strFile_in)
    If strErr = "" Then
        'setup the contents listview
        With lvContents
            .ListItems.Clear
            .Enabled = False
            .Visible = False
            .ColumnHeaders.Clear
            Call .ColumnHeaders.Add(, , m_strBlocks)
            Call .ColumnHeaders.Add(, , m_strImageName)
            Call .ColumnHeaders.Add(, , m_strType)
            .ColumnHeaders(1).Width = 1000
            .ColumnHeaders(2).Width = 2400
            .ColumnHeaders(3).Width = 940
            For Each D64Entry In D64.Entries
                Set LItem = .ListItems.Add(, , D64Entry.FileSizeInBlocks)
                With LItem
                    .SubItems(1) = """" & ConvertPETSCII(D64Entry.Filename) & """"
                    .SubItems(2) = D64Entry.FileType
                End With
                Set LItem = Nothing
            Next D64Entry
            If .ListItems.Count > 0 Then
                .Enabled = True
            End If
            .Visible = True
            lblFileInfo(0) = "D64 " & m_strFile
            lblFileInfo(1) = """" & D64.DiskName & """   """ & D64.DiskID & """"
        End With
        DisplayD64 = ""
    Else
        DisplayD64 = DisplayOtherSupported(strFile_in, strExt_in)
        lblIndex = ""
    End If

End Function


Private Sub FillFiles(strPath_in As String)

    Dim strFile As String
    Dim strPathWBS As String
    Dim strFilter As String

    With lvFiles
        'setup
        MousePointer = vbHourglass
        DoEvents
        .Visible = False
        .Enabled = False
        Call SendMessageAny(lvFiles.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
        .Sorted = False
        strPathWBS = strPath_in
        If Right$(strPathWBS, 1) <> "\" Then
            strPathWBS = strPathWBS & "\"
        End If
        strFilter = lstFilter.List(lstFilter.ListIndex)
        'fill
        With .ListItems
            If g_xobjFunc.DirExists(strPath_in) Then
                strFile = Dir(strPathWBS & strFilter, vbNormal Or vbReadOnly)
                Do Until strFile = ""
                    .Add , , strFile
                    strFile = Dir()
                Loop
            End If
        End With
        'finish
        .Sorted = True
        .Visible = True
        .Enabled = True
        MousePointer = vbDefault
    End With

End Sub

Private Function ParentGamePath(strPath_in As String) As Long

    '-----------------------------------------------
    'gets the parent game path of the passed in path
    'returns the index of the game path found
    'or -1 if not found
    '-----------------------------------------------

    Dim lngPathIdx As Long
    Dim strUpper As String

    ParentGamePath = -1
    strUpper = UCase$(strPath_in)
    For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
        If Left$(strUpper, Len(ga_udtRegPathGame(lngPathIdx).strRealPath)) = UCase$(ga_udtRegPathGame(lngPathIdx).strRealPath) Then
            ParentGamePath = lngPathIdx
            Exit Function
        End If
    Next lngPathIdx

End Function

Private Sub RemoveContentsColumns()

    'removes the columns in lvContents

    Dim lngIdx As Long

    For lngIdx = lvContents.ColumnHeaders.Count To 1 Step -1
        lvContents.ColumnHeaders.Remove lngIdx
    Next lngIdx

End Sub

Private Sub SelectFile(strSearch_in As String)

    Dim strOldFile As String
    Dim blnFound As Boolean
    Dim LItem As ListItem
    
    strOldFile = UCase$(strSearch_in)
    If strOldFile <> "" Then
        blnFound = False
        For Each LItem In lvFiles.ListItems
            If UCase$(LItem.Text) = strOldFile Then
                lvFiles.SelectedItem = LItem
                lvFiles.SelectedItem.Selected = True
                lvFiles.SelectedItem.EnsureVisible
                blnFound = True
                Exit For
            End If
        Next LItem
        If Not blnFound Then
            If lvFiles.ListItems.Count > 0 Then
                lvFiles.SelectedItem = lvFiles.ListItems(1)
                lvFiles.SelectedItem.Selected = True
                lvFiles.SelectedItem.EnsureVisible
                lvFiles_ItemClick lvFiles.SelectedItem
            Else
                Call SetNoFile
            End If
        End If
    Else
        If lvFiles.ListItems.Count > 0 Then
            lvFiles.SelectedItem = lvFiles.ListItems(1)
            lvFiles.SelectedItem.Selected = True
            lvFiles.SelectedItem.EnsureVisible
            lvFiles_ItemClick lvFiles.SelectedItem
        Else
            Call SetNoFile
        End If
    End If

End Sub

Private Sub SelectFirstPRG()

    Dim LItem As ListItem
    Dim blnSet As Boolean

    With lvContents
        If .Enabled = True Then
            If .ColumnHeaders.Count = 3 Then
                'd64
                blnSet = False
                For Each LItem In .ListItems
                    If LItem.SubItems(2) = "PRG" Then
                        .SelectedItem = LItem
                        .SelectedItem.Selected = True
                        .SelectedItem.EnsureVisible
                        lvContents_ItemClick .SelectedItem
                        blnSet = True
                        Exit For
                    End If
                Next LItem
                If blnSet = False Then
                    If .ListItems.Count > 0 Then
                        .SelectedItem = .ListItems(1)
                        .SelectedItem.Selected = True
                        .SelectedItem.EnsureVisible
                        lvContents_ItemClick .SelectedItem
                    Else
                        lblImage = ""
                        lblIndex = ""
                    End If
                End If
            Else
                'other
                If .ListItems.Count > 0 Then
                    .SelectedItem = .ListItems(1)
                    .SelectedItem.Selected = True
                    .SelectedItem.EnsureVisible
                    lvContents_ItemClick .SelectedItem
                Else
                    lblImage = ""
                    lblIndex = ""
                End If
            End If
        Else
            lblImage = ""
            lblIndex = ""
        End If
    End With

End Sub

Private Sub SetNoFile()

    lblFile = g_strNone
    lblDefaultFile = ""
    txtVer.Enabled = False
    txtVer.BackColor = BackColor
    lblImage = ""
    lblIndex = ""
    Set lvFiles.SelectedItem = Nothing
    lvContents.ListItems.Clear
    lvContents.ColumnHeaders.Clear
    lvContents.Enabled = False
    lblFileInfo(1).Caption = g_strNoneSelected
    lblFileInfo(0).Caption = ""
    txtStatus = g_strNone
    cmdTest.Enabled = False
    cmdOK.Enabled = True
    cmdChange.Enabled = False
    m_strDefaultFile = ""

End Sub

Private Sub SetInvalidFile(strStatus_in As String)

    lblFile = g_strNone
    txtVer.Enabled = False
    txtVer.BackColor = BackColor
    lblImage = ""
    lblDefaultFile = ""
    lblIndex = ""
    lblFileInfo(1).Caption = m_strInvalidFile
    lblFileInfo(0).Caption = ""
    lvContents.ListItems.Clear
    lvContents.ColumnHeaders.Clear
    lvContents.Enabled = False
    txtStatus = strStatus_in
    cmdTest.Enabled = False
    cmdOK.Enabled = False
    cmdChange.Enabled = False
    m_strDefaultFile = ""

End Sub
Private Sub SetOKFile(strFile_in As String)

    'strip gamepath from file
    lblFile = Right$(strFile_in, Len(strFile_in) - Len(ga_udtRegPathGame(lstGamePath.ListIndex).strWithBackSlash))
    cmdTest.Enabled = (g_xobjGEMUS.TotalEmulators > 0)
    txtVer.Enabled = True
    txtVer.BackColor = lvFiles.BackColor
    cmdOK.Enabled = True

End Sub


Private Sub TestGame(ByVal lngEmuIdx_in As Long)

    Dim strError As String
    Dim lngGPIdx As Long
    Dim lngImageIdx As Long
    Dim strDefaultFile As String
    Dim udtFields As ManualGEMUSFields

    If g_xobjFunc.FileExists(m_strTestGameFile, False, "") = True Then
        If lblIndex = "" Then
            lngImageIdx = 0
        Else
            lngImageIdx = lblIndex
        End If
        Call ReleaseSidPlayers(True)
        
        'run the game
        Call g_xobjGEMUS.SetRezipNotify(g_blnZipNotify)
        Call g_xobjGEMUS.SetRezip(False, False, eRezipAdditionalIgnore)
                
        If chkFields.Value = vbChecked Then
            Call g_xobjGEMUS.GetManualFieldsDefaults(udtFields, g_strGameBaseName)
            If g_xobjGEMUS.GetManualFieldsFromUser(udtFields, False, g_strGameBaseName) = False Then
                Exit Sub
            End If
            With udtFields
                strError = g_xobjGEMUS.RunEmulator(lngEmuIdx_in, m_strTestGameFile _
                    , lngImageIdx, strDefaultFile, g_udtPathGBGame.strRealPath, True, 1, .lngPlayers, .lngControl, .strComment, .strVComment, .lngPalNTSC, .blnTDE, .strNotes)
            End With
        Else
            strError = g_xobjGEMUS.RunEmulator(lngEmuIdx_in, m_strTestGameFile _
                , lngImageIdx, strDefaultFile, g_udtPathGBGame.strRealPath, False)
        End If

        If strError <> "" Then
            MsgBox strError, vbExclamation, gc_strAPPNAME
        End If
    Else
        MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!"), vbExclamation, gc_strAPPNAME
    End If

End Sub

Private Function WaitForPath(strPath_in As String) As String

    MousePointer = vbHourglass
    DoEvents

    Do
        If MsgBox(g_xobjFunc.GetINIMessage("Game Selector", "PathNotFound", "Path not found!") & vbCrLf & vbCrLf & "'" & strPath_in & "'", vbExclamation Or vbRetryCancel) = vbCancel Then
            'aborted
            WaitForPath = "~~ABORTED~~"
            MousePointer = vbDefault
            Exit Function
        End If
        If g_xobjFunc.DirExists(strPath_in) = True Then
            'found
            WaitForPath = ""
            MousePointer = vbDefault
            Exit Function
        End If
    Loop

End Function

Private Sub chkFields_Click()

    Call ShowTip(gc_lngTIPS_TESTGAMEFLAG)

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

    'get the proper file to display
    strFile = lvFiles.SelectedItem.Text
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
        m_strDefaultFile = strRunnable
        m_blnFirstTimeItemClick = True
        lvFiles.SelectedItem.Selected = True
        lvFiles.SelectedItem.EnsureVisible
        lvFiles_ItemClick lvFiles.SelectedItem
        lvFiles.SetFocus
    End If

End Sub

Private Sub cmdClear_Click()

    Call SetNoFile
    cmdOK.SetFocus

End Sub




Private Sub cmdOK_Click()

    Dim mbRes As VbMsgBoxResult

    'apply to all confirmation
    If chkAll.Value = vbChecked Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSure", "Are you sure you want to update ALL GAMES in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        Else
            'plus version info?
            mbRes = MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSureVersion", "Update ALL GAMES in the current view with the chosen gamefile-version data also?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME)
            If mbRes = vbYes Then
                g_blnGameSelectorApplyAllVersion = True
            ElseIf mbRes = vbNo Then
                g_blnGameSelectorApplyAllVersion = False
            Else
                Exit Sub
            End If
        End If
    End If

    'save values
    g_blnGameSelectorOKPressed = True
    g_strGameSelectorFilename = IIf((lblFile = g_strNone), "", lblFile)
    g_strGameSelectorImageName = IIf((lblImage = g_strNone), "", lblImage)
    g_strGameSelectorDefaultFile = lblDefaultFile
    g_strGameSelectorIndex = IIf((lblIndex = ""), 0, Val(lblIndex))
    g_lngGameSelectorVersion = IIf((txtVer.Text = ""), 0, Val(txtVer.Text))
    g_blnApplyGameFileToAll = (chkAll.Value = vbChecked)

    'save settings and close
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "GameSelectorFilter", lstFilter.ListIndex, False
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastGameSelectorPath", m_strLastGamePath, True
    Unload Me

End Sub


Private Sub cmdRefresh_Click()

    DirBox_Change

End Sub

Private Sub cmdTest_Click()

    Call TestGame(-1)

End Sub

Private Sub cmdTest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu pmnuEmulators
    End If

End Sub


Private Sub DirBox_Change()

    Dim lngPathIdx As Long
    Dim lngCurrIdx As Long
    Dim strError As String
    Dim strStoreFile As String

    lngPathIdx = ParentGamePath(DirBox.Path)
    lngCurrIdx = lstGamePath.ListIndex

    'store selected file
    If lvFiles.SelectedItem Is Nothing Then
        strStoreFile = ""
    Else
        strStoreFile = lvFiles.SelectedItem.Text
    End If

    If lngPathIdx < 0 Then
    
        'invalid outside game paths
        '==========================
        
        If g_xobjFunc.DirExists(m_strLastGamePath) = True Then
            DirBox.Path = m_strLastGamePath
            Call SelectFile(strStoreFile)
        Else
            strError = WaitForPath(m_strLastGamePath)
            If strError <> "" Then
                'can't carry on
                Unload Me
            Else
                DirBox.Path = m_strLastGamePath
                Call SelectFile(strStoreFile)
            End If
        End If
        
    Else
        
        'inside a game path
        '==================
        
        If g_xobjFunc.DirExists(DirBox.Path) = True Then
            If lngPathIdx <> lngCurrIdx Then
                'game path changed
                m_blnByCode = True
                lstGamePath.ListIndex = lngPathIdx
                m_blnByCode = False
            End If
            m_strLastGamePath = DirBox.Path
            Call FillFiles(m_strLastGamePath)
            Call SelectFile(strStoreFile)
        Else
            strError = WaitForPath(DirBox.Path)
            If strError <> "" Then
                'try last path
                strError = WaitForPath(m_strLastGamePath)
                If strError <> "" Then
                    'can't carry on
                    Unload Me
                Else
                    DirBox.Path = m_strLastGamePath
                    Call SelectFile(strStoreFile)
                End If
            Else
                'found new path (user inserted media)
                If lngPathIdx <> lngCurrIdx Then
                    'game path changed
                    m_blnByCode = True
                    lstGamePath.ListIndex = lngPathIdx
                    m_blnByCode = False
                End If
                m_strLastGamePath = DirBox.Path
                Call FillFiles(m_strLastGamePath)
                Call SelectFile(strStoreFile)
            End If
        End If
        
    End If

End Sub

Private Sub DirBox_Click()

    DirBox.Path = DirBox.List(DirBox.ListIndex)

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 116 Then
        'F5 - refresh
        If cmdRefresh.Enabled = True Then
            cmdRefresh_Click
        End If
    End If

End Sub

Private Sub Form_Load()

    Dim lngPathIdx As Long                      'loop index
    Dim lngFilterIdx As Long                    'filter index for file types
    Dim strPath As String                       'for splitting
    Dim strFile As String                       'for splitting
    Dim strFontFile As String                   'for getting the C64 font name
    Dim LItem As ListItem                       'for finding file in filebox
    Dim lngImageIdx As Long                     'loop index
    Const DesignScaleHeight As Single = 6375    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 9975     'The required scalewidth for the form
    Dim a_strFileTypes() As String
    Dim lngTotalValidFileTypes As Long
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim blnRegMaximized As Boolean              'last set window maximized state
    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips

    m_blnByCode = True
    m_lngGPIdx = -1
    m_blnFirstTimeItemClick = True

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngGameSelectorWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngGameSelectorWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngGameSelectorWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngGameSelectorWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "GameSelectorWidth", sngRegWidth, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "GameSelectorHeight", sngRegHeight, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "GameSelectorTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "GameSelectorLeft", sngRegLeft, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "GameSelectorMaximized", blnRegMaximized, False, False
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
        Call GameSelectorSubClass(Me.hwnd)
    End If

    'load the default text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    cmdClear.Caption = g_strButtonClearText
    With g_xobjFunc
        cmdRefresh.Caption = .GetINIMessage("Game Selector", "RefreshButton", "&Refresh")
        cmdTest.Caption = .GetINIMessage("Game Selector", "TestButton", "&Test")
        cmdTest.ToolTipText = frmMain.cmdPlay.ToolTipText
        chkAll.Caption = .GetINIMessage("Game Selector", "ApplyToAll", "&Apply to all games in current view")
        Caption = .GetINIMessage("Game Selector", "Caption", "Game Selector")
        lblPrompt(0) = .GetINIMessage("Game Selector", "RunnableFile", "Runnable File") & ":"
        lblPrompt(1) = .GetINIMessage("Game Selector", "SelectedFile", "Selected File") & ":"
        lvFiles.ColumnHeaders(1).Text = .GetINIMessage("Game Selector", "FileboxHeading", "Game Files")
        m_strSize = .GetINIMessage("Game Selector", "Size", "Size")
        m_strImages = .GetINIMessage("Game Selector", "Images", "Images")
        m_strBlocks = .GetINIMessage("Game Selector", "Blocks", "Blocks")
        m_strImageName = .GetINIMessage("Game Selector", "ImageName", "Image Name")
        m_strType = .GetINIMessage("Game Selector", "Type", "Type")
        m_strVersion = .GetINIMessage("Game Selector", "Version", "Version")
        m_strFile = .GetINIMessage("Game Selector", "File", "File")
        m_strEntries = .GetINIMessage("Game Selector", "Entries", "Entries")
        m_strBytes = .GetINIMessage("Game Selector", "Bytes", "bytes")
        m_strInvalidFile = .GetINIMessage("Game Selector", "InvalidFiletype", "Invalid Emulator File")
        m_strStatusOK = .GetINIMessage("Game Selector", "OK", "OK.")
        m_strLoadAddress = .GetINIMessage("Game Selector", "LoadAddress", "Load Address")
        m_strStatusNoOpenFile = .GetINIMessage("Game Selector", "Status_OpenFileError", "Could not open file.")
        m_strStatusInvalidFile = .GetINIMessage("Game Selector", "InvalidFiletype", "Invalid Emulator File")
        m_strInactive = .GetINIMessage("Game Selector", "Inactive", "Inactive")
        m_strActive = .GetINIMessage("Game Selector", "Active", "Active")
        m_strExromLine = .GetINIMessage("Game Selector", "ExromLine", "EXROM Line")
        m_strGameLine = .GetINIMessage("Game Selector", "GameLine", "GAME Line")
        m_strUnrecCart = .GetINIMessage("Game Selector", "UnrecognisedCartridge", "Unrecognised Cartridge")
        ma_strCartTypes(0) = .GetINIMessage("Game Selector", "NormalCartridge", "Normal Cartridge")
        ma_strCartTypes(1) = "Action Replay"
        ma_strCartTypes(2) = "KCS Power Cartridge"
        ma_strCartTypes(3) = "Final Cartridge III"
        ma_strCartTypes(4) = "Simons Basic"
        ma_strCartTypes(5) = "Ocean type 1 (128K and 256K)"
        ma_strCartTypes(6) = "Expert Cartridge"
        ma_strCartTypes(7) = "Fun Play"
        ma_strCartTypes(8) = "Super Games"
        ma_strCartTypes(9) = "Atomic Power"
        ma_strCartTypes(10) = "Epyx Fastload"
        ma_strCartTypes(11) = "Westermann"
        ma_strCartTypes(12) = "Rex"
        ma_strCartTypes(13) = "Final Cartridge I"
    End With
    
    Call MakeTablePETSCII
    
    'are we adding?
    If g_blnGameSelectorNewGame = True Then
        chkAll.Enabled = False
    End If
    
    'load emulator selector list
    Call LoadEmulatorPopupMenu

    'load the game paths into the gamepath listbox
    For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
        lstGamePath.AddItem ga_udtRegPathGame(lngPathIdx).strRealPath
    Next lngPathIdx

    'load the file filters
    lngTotalValidFileTypes = g_xobjGEMUS.GetAllValidFileTypesArray(a_strFileTypes())
    With lstFilter
        .AddItem "*.*"
        For lngPathIdx = 0 To lngTotalValidFileTypes - 1
            .AddItem "*." & a_strFileTypes(lngPathIdx)
        Next lngPathIdx
        .AddItem "*.7Z"
        .AddItem "*.ARJ"
        .AddItem "*.BZ2"
        .AddItem "*.BZIP2"
        .AddItem "*.CAB"
        .AddItem "*.GZ"
        .AddItem "*.GZIP"
        .AddItem "*.ISO"
        .AddItem "*.LHA"
        .AddItem "*.LZH"
        .AddItem "*.LZMA"
        .AddItem "*.RAR"
        .AddItem "*.TAR"
        .AddItem "*.TAZ"
        .AddItem "*.TBZ"
        .AddItem "*.TBZ2"
        .AddItem "*.TGZ"
        .AddItem "*.Z"
        .AddItem "*.ZIP"
    End With
    lvFiles.ColumnHeaders(1).Width = lvFiles.Width - 380

    'get last selected filter
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "GameSelectorFilter", lngFilterIdx, False, 0
    If lngFilterIdx < 0 Or (lngFilterIdx > lstFilter.ListCount - 1) Then
        lngFilterIdx = 0
    End If

    'set the version
    txtVer.Text = g_lngGameSelectorVersion
    
    g_strGameSelectorImageName = ""

    'we know there is a valid game path whenever this form is shown
    'so we just need to select it depending on if the gamefile is found or not
    'if g_strGameSelectorFilename is not "" then the file exists

    If g_strGameSelectorFilename = "" Then
        
        'no gamefile or gamefile not found
        '=================================

        'get last dir traversed
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastGameSelectorPath", m_strLastGamePath, True, ""
        'check if it exists and set the gamepath selector index
        If g_xobjFunc.DirExists(m_strLastGamePath) = True Then
            'is it within a game path?
            lngPathIdx = ParentGamePath(m_strLastGamePath)
            If lngPathIdx > -1 Then
                'ok!
                lstGamePath.ListIndex = lngPathIdx
            Else
                'not in any game paths - set it to first existing game path
                lngPathIdx = GetExistingGamePath()
                m_strLastGamePath = ga_udtRegPathGame(lngPathIdx).strRealPath
                lstGamePath.ListIndex = lngPathIdx
            End If
        Else
            'nope - set it to first existing game path
            lngPathIdx = GetExistingGamePath()
            m_strLastGamePath = ga_udtRegPathGame(lngPathIdx).strRealPath
            lstGamePath.ListIndex = lngPathIdx
        End If
        'set the dirbox path and let the filebox fill
        m_blnLoadBusy = True
        m_strDefaultFile = ""
        lstFilter.ListIndex = lngFilterIdx
        If DirBox.Path <> m_strLastGamePath Then
            DirBox.Path = m_strLastGamePath
        Else
            DirBox_Change
        End If
        DoEvents
        m_blnLoadBusy = False
        'and set no file
        Call SetNoFile
    Else

        'gamefile found
        '==============
        'set the game path
        lstGamePath.ListIndex = ParentGamePath(g_strGameSelectorFilename)
        'set the dirbox path and let the filebox fill
        lstFilter.ListIndex = 0 'all files
        m_strDefaultFile = g_strGameSelectorDefaultFile
        Call g_xobjFunc.GetFileAndPath(g_strGameSelectorFilename, strPath, strFile)
        m_blnLoadBusy = True
        If DirBox.Path <> strPath Then
            DirBox.Path = strPath
        Else
            DirBox_Change
        End If
        DoEvents
        m_blnLoadBusy = False
        'find and select the file
        strFile = UCase$(strFile)
        For Each LItem In lvFiles.ListItems
            If strFile = UCase$(LItem.Text) Then
                Set lvFiles.SelectedItem = LItem
                lvFiles.SelectedItem.Selected = True
                lvFiles.SelectedItem.EnsureVisible
                lvFiles_ItemClick LItem
                'select the correct index
                For lngImageIdx = 1 To lvContents.ListItems.Count
                    If lngImageIdx = (g_strGameSelectorIndex + 1) Then
                        lvContents.SelectedItem = lvContents.ListItems(lngImageIdx)
                        lvContents.SelectedItem.Selected = True
                        lvContents_ItemClick lvContents.SelectedItem
                        Exit For
                    End If
                Next lngImageIdx
                Exit For
            End If
        Next LItem
        
    End If

    m_blnByCode = False
    g_blnGameSelectorOKPressed = False
    
End Sub

Private Sub Form_Resize()

    Dim sngHalfHeight As Single

    If WindowState <> vbMinimized Then
    
        'width
        With cmdRefresh
            .Left = (ScaleWidth - 120) - .Width
            lstFilter.Width = (.Left - lstFilter.Left) - 90
        End With
        With lvFiles
            .Width = (ScaleWidth - .Left) - 120
            .ColumnHeaders(1).Width = .Width - 380
        End With
        With lblFile
            .Width = lvFiles.Width - cmdClear.Width
            cmdClear.Left = .Left + .Width
            txtStatus.Width = .Width
        End With
        cmdChange.Left = cmdClear.Left
        With cmdCancel
            .Left = ScaleWidth - 120 - .Width
            cmdOK.Left = .Left - 120 - cmdOK.Width
        End With
        chkAll.Left = (cmdOK.Left - chkAll.Width) - 120
        
        'height
        With cmdCancel
            .Top = ScaleHeight - 120 - .Height
            cmdOK.Top = .Top
            chkAll.Top = .Top - 120
            lblVer.Top = chkAll.Top
            txtVer.Top = lblVer.Top + 240
            cmdTest.Top = .Top
            chkFields.Top = .Top + 120
        End With
        With txtStatus
            .Top = cmdCancel.Top - 480
            cmdChange.Top = .Top
            lblPrompt(0).Top = .Top - 240
            lblFile.Top = .Top - 600
            cmdClear.Top = lblFile.Top
            lblPrompt(1).Top = lblFile.Top - 240
        End With
        lvFiles.Height = (lblPrompt(1).Top - 120) - lvFiles.Top
    
        sngHalfHeight = ((cmdTest.Top - 120) - DirBox.Top) / 2
        DirBox.Height = sngHalfHeight
        lblFileInfo(0).Top = DirBox.Top + sngHalfHeight + 90
        lblFileInfo(1).Top = lblFileInfo(0).Top + 240
        lvContents.Top = lblFileInfo(1).Top + 360
        lvContents.Height = (cmdTest.Top - 120) - lvContents.Top
        
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SendMessageAny(lvFiles.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnGameSelectorSubclassed = True Then
            Call GameSelectorUnSubClass(Me.hwnd)
        End If
    End If

    'save reg settings
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "GameSelectorMaximized", True, False
        End If
    End With

End Sub


Private Sub lstFilter_Click()

    Dim strOldFile As String

    If m_blnByCode = False Then
        'store selected file
        If lvFiles.SelectedItem Is Nothing Then
            strOldFile = ""
        Else
            strOldFile = UCase$(lvFiles.SelectedItem.Text)
        End If
        'show the filtered file list
        Call FillFiles(DirBox.Path)
        're-select the file if poss
        Call SelectFile(strOldFile)
    End If

End Sub


Private Sub lstGamePath_Click()

    Dim strError As String
    Dim strStoreFile As String

    If m_blnByCode = False Then
        'store selected file
        If lvFiles.SelectedItem Is Nothing Then
            strStoreFile = ""
        Else
            strStoreFile = lvFiles.SelectedItem.Text
        End If
        'change game path
        With ga_udtRegPathGame(lstGamePath.ListIndex)
            If g_xobjFunc.DirExists(.strRealPath) = False Then
                strError = WaitForPath(.strRealPath)
                If strError <> "" Then
                    'back to what it was (don't change game path)
                    lstGamePath.ListIndex = m_lngGPIdx
                Else
                    'store for next time
                    m_lngGPIdx = lstGamePath.ListIndex
                    'move to the path
                    DirBox.Path = .strRealPath
                    Call SelectFile(strStoreFile)
                End If
            Else
                'store for next time
                m_lngGPIdx = lstGamePath.ListIndex
                'move to the path
                DirBox.Path = .strRealPath
                Call SelectFile(strStoreFile)
            End If
        End With
    Else
        'store for next time
        m_lngGPIdx = lstGamePath.ListIndex
    End If

End Sub

Private Sub lvContents_ItemClick(ByVal Item As MSComctlLib.ListItem)

    With lvContents
        If .ColumnHeaders.Count = 3 Then
            lblImage = .SelectedItem.SubItems(1)
        Else
            lblImage = .SelectedItem.Text
        End If
        lblIndex = .SelectedItem.Index - 1
    End With

End Sub



Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

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
    Dim strRunnablePathFile As String
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
    strFile = Item.Text
    strPathFileToTest = DirBox.Path
    If Right$(strPathFileToTest, 1) <> "\" Then
        strPathFileToTest = strPathFileToTest & "\"
    End If
    strPathFileToTest = strPathFileToTest & strFile
    MousePointer = vbHourglass
    DoEvents
    lngTotalValidFileTypes = g_xobjGEMUS.GetAllValidFileTypesArray(a_strFileTypes())
    
    If m_blnFirstTimeItemClick Then
        strRunnable = m_strDefaultFile
    Else
        strRunnable = ""
    End If
    m_blnFirstTimeItemClick = False

    '********************************* try to unpack/display the chosen file *********************************
    blnSuccess = False
    blnUserCancelled = False
    strUnpackToWBS = g_udtPathGBGame.strWithBackSlash
    If Unpacker.Unpack(strPathFileToTest, strRunnable, a_strFileTypes, lngTotalValidFileTypes, _
                                    strUnpackToWBS, blnIsRepackable, strRunnablePathFile, _
                                    blnUserMustBrowseArchive, strError, strNotUsed) _
    Then
        If blnUserMustBrowseArchive Then
            'more than one compatible game file in archive, ask user to pick one
            If Unpacker.GetRunnableFileInArchiveFromUser(strPathFileToTest, strUnpackToWBS, a_strFileTypes, strRunnable) Then
                'try again now the user picked a runnable file
                If Unpacker.Unpack(strPathFileToTest, strRunnable, a_strFileTypes, lngTotalValidFileTypes, _
                                                strUnpackToWBS, blnIsRepackable, strRunnablePathFile, _
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
        strExt = UCase$(Right$(strRunnablePathFile, Len(strRunnablePathFile) - InStrRev(strRunnablePathFile, ".")))
        strError = DisplayFile(strRunnablePathFile, strExt)
        If strError = "" Then
            'ok
            m_strTestGameFile = strRunnablePathFile
            m_strDefaultFile = strRunnable
            lblDefaultFile = strRunnable
            txtStatus.Text = strRunnable
            Call SelectFirstPRG
            Call SetOKFile(strPathFileToTest)
            cmdChange.Enabled = Unpacker.IsFileASupportedArchiveFileType(strPathFileToTest)
        Else
            'error displaying file
            Call SetInvalidFile(strError)
        End If
    ElseIf blnUserCancelled = True Then
        Call SetNoFile
    Else
        Call SetInvalidFile(strError)
    End If
    
    MousePointer = vbDefault
    Set Unpacker = Nothing

End Sub

Private Sub pmnuEmulator_Click(Index As Integer)
    
    Call TestGame(Index)
    
End Sub


Private Sub txtVer_GotFocus()

    With txtVer
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtVer_KeyPress(KeyAscii As Integer)

    'allow delete and backspace keypresses
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    
    'only numbers
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub


