VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlPathSelector 
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ScaleHeight     =   2265
   ScaleWidth      =   4770
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "ctlPathSelector.ctx":0000
      Left            =   0
      List            =   "ctlPathSelector.ctx":0002
      TabIndex        =   1
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
            Picture         =   "ctlPathSelector.ctx":0004
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlPathSelector.ctx":00A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlPathSelector.ctx":0134
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlPathSelector.ctx":01C6
            Key             =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbButtons 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   582
      ButtonWidth     =   609
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
Attribute VB_Name = "ctlPathSelector"
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

Public Sub AddItem(strPath As String)

    lstPaths.AddItem strPath

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

Public Function GetLastTraversedPath() As String

    GetLastTraversedPath = m_strLastTraversedPath

End Function

Public Property Get List(ByVal intIndex As Integer) As String

    List = lstPaths.List(intIndex)

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

Private Sub SetButtonsState()

    '----------------------------------------------------
    'sets the enabled state of the buttons in the toolbar
    '----------------------------------------------------

    With lstPaths
        tbButtons.Buttons("new").Enabled = (.ListCount <= m_lngMaxEntries)
        tbButtons.Buttons("delete").Enabled = (.ListIndex > gc_lngNOLISTINDEX)
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

    If lstPaths.ListIndex > gc_lngNOLISTINDEX Then

        'get an initial valid path to select
        If g_xobjFunc.DirExists(lstPaths.List(lstPaths.ListIndex)) Then
            m_strLastTraversedPath = lstPaths.List(lstPaths.ListIndex)
        ElseIf Not g_xobjFunc.DirExists(m_strLastTraversedPath) Then
            m_strLastTraversedPath = App.Path
        End If
        'get the new path
        g_strBrowseForFolder = m_strLastTraversedPath
        frmBrowseFolder.Show vbModal
        
        'change the folder in the listbox if path is selected and save the last traversed path
        If g_strBrowseForFolder <> "" Then
            lstPaths.List(lstPaths.ListIndex) = g_strBrowseForFolder
            m_strLastTraversedPath = g_strBrowseForFolder
        End If
        
    End If

End Sub


Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngOldIdx As Long       'for the "delete" case
    Dim strTempPath As String   'for swapping on the "up" and "down" cases

    Select Case Button.Key

    Case "new"

        'get an initial valid path to select
        If Not g_xobjFunc.DirExists(m_strLastTraversedPath) Then
            m_strLastTraversedPath = App.Path
        End If
        'get the new path
        g_strBrowseForFolder = m_strLastTraversedPath
        frmBrowseFolder.Show vbModal
        
        'add to listbox if path is selected and save the last traversed path
        If g_strBrowseForFolder <> "" Then
            With lstPaths
                .AddItem g_strBrowseForFolder
                .ListIndex = .NewIndex
            End With
            m_strLastTraversedPath = g_strBrowseForFolder
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
                    .ListIndex = gc_lngNOLISTINDEX
                    .ListIndex = lngOldIdx
                End If
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

    Case "down"

        'swap path with the one below it
        With lstPaths
            strTempPath = .List(.ListIndex + 1)
            .List(.ListIndex + 1) = .List(.ListIndex)
            .List(.ListIndex) = strTempPath
            .ListIndex = .ListIndex + 1
        End With
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


