VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlEmuSelector 
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   ScaleHeight     =   4740
   ScaleWidth      =   7305
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "ctlEmuSelector.ctx":0000
      Left            =   0
      List            =   "ctlEmuSelector.ctx":0002
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
            Picture         =   "ctlEmuSelector.ctx":0004
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlEmuSelector.ctx":00A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlEmuSelector.ctx":0134
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlEmuSelector.ctx":01C6
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
      Width           =   7305
      _ExtentX        =   12885
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
Attribute VB_Name = "ctlEmuSelector"
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
Private m_lngMaxEntries As Long                 'maximum number of paths allowed
Private Type EmulatorRegType
    strName As String
    strExe As String
    strScript As String
    blnUseShort As Boolean
    strFiles() As String
    lngNumFiles As Long
    blnUseDefaultTypes As Boolean
    strTypes As String
    lngWait As Long
End Type
Private ma_udtEmulators() As EmulatorRegType     'the emulator data
Private m_strGBName As String
Private m_strEmuFile As String
Public Sub AddItem(strName_in As String, strExe_in As String, strScript_in As String, ByVal blnUseShort_in As Boolean, strFiles_in() As String, ByVal lngNumFiles_in As Long, ByVal blnUseDefaultTypes_in As Boolean, strTypes_in As String, ByVal lngWait_in As Long)

    'add an emulator to the list

    ReDim Preserve ma_udtEmulators(lstPaths.ListCount)

    With ma_udtEmulators(lstPaths.ListCount)
        .blnUseShort = blnUseShort_in
        .strScript = strScript_in
        .strExe = strExe_in
        .strName = strName_in
        .strFiles() = strFiles_in()
        .lngNumFiles = lngNumFiles_in
        .blnUseDefaultTypes = blnUseDefaultTypes_in
        .strTypes = strTypes_in
        .lngWait = lngWait_in
    End With

    lstPaths.AddItem ma_udtEmulators(lstPaths.ListCount).strName

End Sub

Public Sub GetItem(ByVal intIndex As Integer, strName_out As String, strExe_out As String, strScript_out As String, blnUseShort_out As Boolean, strFiles_out() As String, lngNumFiles_out As Long, blnUseDefaultTypes_out As Boolean, strTypes_out As String, lngWait_out As Long)

    '------------------------------------------------------------------------------
    'gets the info for the emulator at intIndex in the listbox
    '
    'Inputs:    intIndex:           The index of the emulator in the listbox
    '
    'Outputs:   strName_out:        Var to receive the emulator name
    '           strExe_out:         Var to receive the emulator exe
    '           strScript_out:      Var to receive the emulator command line params
    '           blnUseShort_out:    Var to receive the Use Short Filenames flag
    '           strFiles_out:       Array to receive the emulator files
    '           lngNumFiles_out:    Var to receive the number of emulator files
    '------------------------------------------------------------------------------

    'copy across the values
    With ma_udtEmulators(intIndex)
        strName_out = .strName
        strExe_out = .strExe
        strScript_out = .strScript
        blnUseShort_out = .blnUseShort
        strFiles_out() = .strFiles()
        lngNumFiles_out = .lngNumFiles
        blnUseDefaultTypes_out = .blnUseDefaultTypes
        strTypes_out = .strTypes
        lngWait_out = .lngWait
    End With

End Sub


Public Property Let ItemData(ByVal intIndex As Integer, ByVal lngData As Long)

    lstPaths.ItemData(intIndex) = lngData

End Property

Public Property Get List(ByVal intIndex As Integer) As String

    List = lstPaths.List(intIndex)

End Property

Public Property Get ItemData(ByVal intIndex As Integer) As Long

    ItemData = lstPaths.ItemData(intIndex)

End Property
Public Property Get ListCount()

    ListCount = lstPaths.ListCount

End Property

Public Property Get ListIndex()

    ListIndex = lstPaths.ListIndex

End Property

Public Property Let MaxEntries(ByVal lngNewVal As Long)

    m_lngMaxEntries = lngNewVal
    PropertyChanged "MaxEntries"

End Property

Public Property Get MaxEntries() As Long

    MaxEntries = m_lngMaxEntries

End Property

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

Public Sub SetGameBaseNameInfo(strGBName_in As String, strEmuFile_in As String)

    m_strGBName = strGBName_in
    m_strEmuFile = strEmuFile_in

End Sub

Private Sub lstPaths_Click()

    'set the state of the buttons
    Call SetButtonsState
        
End Sub

Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngOldIdx As Long                   'for the "delete" case
    Dim strTempPath As String               'for swapping on the "up" and "down" cases
    Dim lngEmulatorIdx As Long              'for deleting/swapping
    Dim udtTempEmulator As EmulatorRegType  'for swapping on the "up" and "down" cases

    Select Case Button.Key

    Case "new"

        'get the new sid player properties
        With g_udtEmulatorProps
            .strName = ""
            .strFullPath = ""
            .strScript = ""
            .blnUseShort = False
            .lngNumFiles = 0
            .blnUseDefaultTypes = True
            .strTypes = ""
            .lngWait = 1
            ReDim .strFiles(0)
            .strGBName = m_strGBName
            .strEmuFile = m_strEmuFile
        End With
        frmEmulatorProps.Show vbModal
        'did we get a new player?
        If g_udtEmulatorProps.strName <> "" Then
            With g_udtEmulatorProps
                Call AddItem(.strName, .strFullPath, .strScript, .blnUseShort, .strFiles(), .lngNumFiles, .blnUseDefaultTypes, .strTypes, .lngWait)
            End With
            With lstPaths
                .ListIndex = .NewIndex
            End With
        End If

    Case "delete"

        'delete the current selected path from the listbox
        With lstPaths
            'delete from array
            For lngEmulatorIdx = .ListIndex + 1 To UBound(ma_udtEmulators)
                ma_udtEmulators(lngEmulatorIdx - 1) = ma_udtEmulators(lngEmulatorIdx)
            Next lngEmulatorIdx
            If UBound(ma_udtEmulators) > 0 Then
                ReDim Preserve ma_udtEmulators(UBound(ma_udtEmulators) - 1)
            End If
            'delete from listbox
            lngOldIdx = .ListIndex
            .RemoveItem .ListIndex
            If .ListCount = 0 Then
                Call SetButtonsState
            Else
                If lngOldIdx = .ListCount Then
                    'last player deleted
                    .ListIndex = lngOldIdx - 1
                Else
                    'force click on same player
                    .ListIndex = -1
                    .ListIndex = lngOldIdx
                End If
            End If
        End With

    Case "up"

        'swap player with the one above it
        With lstPaths
            'array
            udtTempEmulator = ma_udtEmulators(.ListIndex - 1)
            ma_udtEmulators(.ListIndex - 1) = ma_udtEmulators(.ListIndex)
            ma_udtEmulators(.ListIndex) = udtTempEmulator
            'listbox
            strTempPath = .List(.ListIndex - 1)
            .List(.ListIndex - 1) = .List(.ListIndex)
            .List(.ListIndex) = strTempPath
            .ListIndex = .ListIndex - 1
        End With

    Case "down"

        'swap player with the one below it
        With lstPaths
            'array
            udtTempEmulator = ma_udtEmulators(.ListIndex + 1)
            ma_udtEmulators(.ListIndex + 1) = ma_udtEmulators(.ListIndex)
            ma_udtEmulators(.ListIndex) = udtTempEmulator
            'listbox
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
    MaxEntries = 100

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'read the properties from the control's runtime container
    MaxEntries = PropBag.ReadProperty("MaxEntries", 200)

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
    PropBag.WriteProperty "MaxEntries", MaxEntries, 200

End Sub



