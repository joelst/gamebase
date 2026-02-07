VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlFileSelector 
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   KeyPreview      =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   4140
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4200
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "ctlFileSelector.ctx":0000
      Left            =   0
      List            =   "ctlFileSelector.ctx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin MSComctlLib.ImageList iLstToolbar 
      Left            =   4080
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFileSelector.ctx":0004
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFileSelector.ctx":00A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFileSelector.ctx":0134
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFileSelector.ctx":01C6
            Key             =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbButtons 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "iLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "up"
            ImageKey        =   "up"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "down"
            ImageKey        =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlFileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private m_blnOKToResize As Boolean      'resize recursion flag

'properties and private variables
Private m_strBrowsePrompt As String         'prompt title when browsing for folders
Private m_strPathRegKey As String           'registry key to save last traversed path
Private m_strLastTraversedPath As String    'last traversed path in the browse api window
Private m_lngMaxEntries As Long             'maximum number of paths allowed
Private m_strGBName As String

Public Event KeyPress(KeyAscii As Integer)

Public Sub AddItem(strPath As String)

    lstPaths.AddItem (lstPaths.ListCount + 1) & ". " & strPath

End Sub

Public Property Let BrowsePrompt(strNewPrompt As String)

    m_strBrowsePrompt = strNewPrompt
    PropertyChanged "BrowsePrompt"

End Property
Public Property Let BrowseRegKey(strNewKey As String)

    m_strPathRegKey = strNewKey
    PropertyChanged "BrowseRegKey"

End Property

Public Property Get BrowsePrompt() As String

    BrowsePrompt = m_strBrowsePrompt

End Property

Public Property Get BrowseRegKey() As String
    
    BrowseRegKey = m_strPathRegKey
    
End Property

Public Property Let GameBaseName(strName_in As String)

    m_strGBName = strName_in

End Property

Public Function GetLastTraversedPath() As String

    GetLastTraversedPath = m_strLastTraversedPath

End Function

Public Property Get List(ByVal intIndex As Integer) As String

    With lstPaths
        List = Right$(.List(intIndex), Len(.List(intIndex)) - InStr(.List(intIndex), ". ") - 1)
    End With

End Property
Public Property Get ListCount()

    ListCount = lstPaths.ListCount

End Property

Public Property Get ListIndex()

    ListIndex = lstPaths.ListIndex

End Property

Public Property Let MaxPathEntries(ByVal lngNewVal As Long)

    m_lngMaxEntries = lngNewVal
    PropertyChanged "MaxPathEntries"

End Property

Public Property Get MaxPathEntries() As Long

    MaxPathEntries = m_lngMaxEntries

End Property

Private Sub RenumberAssociatedFiles()

    Dim lngFileIdx As Long  'loop index
    Dim strFile As String   'without number

    With lstPaths
        For lngFileIdx = 0 To (.ListCount - 1)
            strFile = Right$(.List(lngFileIdx), Len(.List(lngFileIdx)) - InStr(.List(lngFileIdx), ".") + 1)
            .List(lngFileIdx) = Trim$(lngFileIdx + 1) & strFile
        Next lngFileIdx
    End With

End Sub

Private Sub SetButtonsState()

    '----------------------------------------------------
    'sets the enabled state of the buttons in the toolbar
    '----------------------------------------------------

    With lstPaths
        tbButtons.Buttons("new").Enabled = (.ListCount <= m_lngMaxEntries)
        tbButtons.Buttons("delete").Enabled = (.ListIndex > -1)
        tbButtons.Buttons("up").Enabled = (.ListIndex > 0)
        tbButtons.Buttons("down").Enabled = (.ListIndex < .ListCount - 1)
    End With

End Sub

Public Sub SetLastTraversedPath(strPath As String)

    m_strLastTraversedPath = strPath

End Sub

Private Sub lstPaths_Click()

    'set the state of the buttons
    Call SetButtonsState
        
End Sub

Private Sub lstPaths_DblClick()

    Dim strNewFile As String

    'get the  path\filename
    With lstPaths
        strNewFile = GetAssociatedFile(Right$(.List(.ListIndex), Len(.List(.ListIndex)) - InStr(.List(.ListIndex), ".") - 1), m_strLastTraversedPath)
        'add to listbox if path\filename is selected
        If strNewFile <> "" Then
            .List(.ListIndex) = Trim$(.ListIndex + 1) & ". " & strNewFile
        End If
    End With

End Sub


Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngOldIdx As Long       'for the "delete" case
    Dim strTempPath As String   'for swapping on the "up" and "down" cases
    Dim strNewFile As String    'the new filename when added
    

    Select Case Button.Key

    Case "new"

        'get the new path\filename
        strNewFile = GetAssociatedFile("", m_strLastTraversedPath)
        'add to listbox if path\filename is selected
        If strNewFile <> "" Then
            With lstPaths
                AddItem strNewFile
                .ListIndex = .NewIndex
            End With
        End If

    Case "delete"

        'delete the current selected path from the listbox
        With lstPaths
            lngOldIdx = .ListIndex
            .RemoveItem .ListIndex
            If .ListCount = 0 Then
                Call SetButtonsState
            Else
                If lngOldIdx = .ListCount Then
                    'last path deleted
                    .ListIndex = lngOldIdx - 1
                Else
                    'force click on same path
                    .ListIndex = -1
                    .ListIndex = lngOldIdx
                End If
                Call RenumberAssociatedFiles
            End If
        End With

    Case "up"

        'swap path with the one above it
        With lstPaths
            strTempPath = .List(.ListIndex - 1)
            .List(.ListIndex - 1) = .List(.ListIndex)
            .List(.ListIndex) = strTempPath
            .ListIndex = .ListIndex - 1
        End With
        Call RenumberAssociatedFiles

    Case "down"

        'swap path with the one below it
        With lstPaths
            strTempPath = .List(.ListIndex + 1)
            .List(.ListIndex + 1) = .List(.ListIndex)
            .List(.ListIndex) = strTempPath
            .ListIndex = .ListIndex + 1
        End With
        Call RenumberAssociatedFiles
    
    End Select

End Sub
Private Sub UserControl_Initialize()

    m_blnOKToResize = True

End Sub


Private Sub UserControl_InitProperties()

    'set the default values for the properties
    BrowsePrompt = "Select Folder..."
    BrowseRegKey = ""
    MaxPathEntries = 20

End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'read the properties from the control's runtime container
    BrowsePrompt = PropBag.ReadProperty("BrowsePrompt", "Select Folder...")
    BrowseRegKey = PropBag.ReadProperty("BrowseRegKey", "")
    MaxPathEntries = PropBag.ReadProperty("MaxPathEntries", 20)

End Sub

Private Sub UserControl_Resize()

    If m_blnOKToResize Then

        'width
        lstPaths.Width = Width

        'height
        lstPaths.Top = tbButtons.Height
        lstPaths.Height = Height - tbButtons.Height

        'resize to snapped height because of listbox
        m_blnOKToResize = False
        Height = lstPaths.Top + lstPaths.Height
        m_blnOKToResize = True
    End If

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'save the properties to the control's runtime container
    PropBag.WriteProperty "BrowsePrompt", BrowsePrompt, "Select Folder..."
    PropBag.WriteProperty "BrowseRegKey", BrowseRegKey, ""
    PropBag.WriteProperty "MaxPathEntries", MaxPathEntries, 20

End Sub

Private Function GetAssociatedFile(strFilename_in As String, strPath_in As String) As String

    Dim strFile As String       'for splitting into path\file vars
    Dim strPath As String       'for splitting into path\file vars
    Dim strINIFile As String
    Dim xobjFunc As GBFuncs.clsFunctions
    Dim xobjRegistry As GBFuncs.clsRegistry

    On Error GoTo Cancel_Pressed

    Set xobjFunc = New GBFuncs.clsFunctions
    Set xobjRegistry = New GBFuncs.clsRegistry
    strINIFile = App.Path
    If Right$(strINIFile, 1) <> "\" Then strINIFile = strINIFile & "\"
    strINIFile = strINIFile & "GemusText.ini"

    'set up the open dialog
    With cDialog
        .CancelError = True
        .DefaultExt = ".exe"
        .DialogTitle = GetINIMessageFromFile("Emulator Property Window", "AssociatedFilePrompt", "Select Associated File", strINIFile)
        .Filter = "All Files (*.*)|*.*|INI Files (*.ini)|*.ini|Config Files (*.cfg)|*.cfg|Text Files (*.txt)|*.txt|Executables (*.exe)|*.exe"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        If xobjFunc.FileExists(strFilename_in, False, "") = True Then
            .Filename = strFilename_in
        ElseIf xobjFunc.DirExists(strPath_in) = True Then
            .InitDir = strPath_in
            .Filename = ""
        Else
            .InitDir = App.Path
            .Filename = ""
        End If
    End With

    'get the file
    cDialog.ShowOpen
    GetAssociatedFile = cDialog.Filename

    'save the last path
    xobjFunc.GetFileAndPath cDialog.Filename, strPath, strFile
    xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & m_strGBName & "\BrowsePaths", "LastEmuFilePath", strPath, True
    m_strLastTraversedPath = strPath
    Set xobjRegistry = Nothing
    Set xobjFunc = Nothing
    Exit Function

Cancel_Pressed:
    GetAssociatedFile = ""
    
    Set xobjRegistry = Nothing
    Set xobjFunc = Nothing
    
End Function

