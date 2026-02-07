VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlExtrasSelector 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "ctlExtrasSelector.ctx":0000
      Left            =   0
      List            =   "ctlExtrasSelector.ctx":0002
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
            Picture         =   "ctlExtrasSelector.ctx":0004
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExtrasSelector.ctx":00A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExtrasSelector.ctx":0134
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExtrasSelector.ctx":01C6
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
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "iLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "up"
            ImageKey        =   "up"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "down"
            ImageKey        =   "down"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "spacer"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delete"
            ImageKey        =   "delete"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlExtrasSelector"
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

'resize recursion flag
Private m_blnOKToResize As Boolean

'properties and private variables
Private m_lngMaxEntries As Long
Private Type ExtrasStoreType
    lngEX_Id As Long
    strName As String
    strPath As String
    lngType As Long
    strData As String
    strRunnableFile As String
End Type
Private ma_udtExtrasStore() As ExtrasStoreType

Public Event EnableAllCheckbox()
Public Event DisableAllCheckbox()
Public Sub AddItem(ByVal lngEXID_in As Long, strName_in As String, strPath_in As String, ByVal lngType_in As Long, strData_in As String, strRunnableFile_in As String)

    'add an extra to the list

    ReDim Preserve ma_udtExtrasStore(lstPaths.ListCount)
    With ma_udtExtrasStore(lstPaths.ListCount)
        .lngEX_Id = lngEXID_in
        .lngType = lngType_in
        .strData = strData_in
        .strName = strName_in
        .strPath = strPath_in
        .strRunnableFile = strRunnableFile_in
    End With
    lstPaths.AddItem ma_udtExtrasStore(lstPaths.ListCount).strName

End Sub


Public Sub GetItem(ByVal intIndex As Integer, lngEXID_out As Long, strName_out As String, strPath_out As String, lngType_out As Long, strData_out As String, strRunnableFile_out As String)
  
    'gets the info for the extra at intIndex in the listbox

    'copy across the values
    With ma_udtExtrasStore(intIndex)
        lngEXID_out = .lngEX_Id
        lngType_out = .lngType
        strData_out = .strData
        strName_out = .strName
        strPath_out = .strPath
        strRunnableFile_out = .strRunnableFile
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
        tbButtons.Buttons("delete").Enabled = (.ListIndex > gc_lngNOLISTINDEX)
        tbButtons.Buttons("up").Enabled = (.ListIndex > 0)
        tbButtons.Buttons("down").Enabled = (.ListIndex < .ListCount - 1)
    End With

End Sub

Public Sub SetListIndex(ByVal NewIndex As Long)

    lstPaths.ListIndex = NewIndex

End Sub

Private Sub lstPaths_Click()

    'set the state of the buttons
    Call SetButtonsState
        
End Sub

Private Sub lstPaths_DblClick()

    Dim strPath As String
    Dim strFile As String
    Dim strError As String

    If lstPaths.ListIndex > gc_lngNOLISTINDEX Then
        'get the starting path and file
        If ma_udtExtrasStore(lstPaths.ListIndex).strPath <> "" And ma_udtExtrasStore(lstPaths.ListIndex).lngType < 3 Then
            strError = FindExtrasFile(g_udtExtrasProps.lngPathIdx, ma_udtExtrasStore(lstPaths.ListIndex).strPath)
            If strError <> "" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
                strError = GetValidExtrasPath(g_udtExtrasProps.lngPathIdx)
                If strError <> "" Then
                    MsgBox strError, vbExclamation, gc_strAPPNAME
                Else
                    Call g_xobjFunc.GetFileAndPath(ga_udtRegPathExtras(g_udtExtrasProps.lngPathIdx).strWithBackSlash & ma_udtExtrasStore(lstPaths.ListIndex).strPath, strPath, strFile)
                    If g_xobjFunc.DirExists(strPath) = False Then
                        strPath = ga_udtRegPathExtras(g_udtExtrasProps.lngPathIdx).strRealPath
                    End If
                End If
            Else
                Call g_xobjFunc.GetFileAndPath(ga_udtRegPathExtras(g_udtExtrasProps.lngPathIdx).strWithBackSlash & ma_udtExtrasStore(lstPaths.ListIndex).strPath, strPath, strFile)
            End If
        Else
            'no file
            strError = GetValidExtrasPath(g_udtExtrasProps.lngPathIdx)
            If strError <> "" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
            Else
                strPath = ga_udtRegPathExtras(g_udtExtrasProps.lngPathIdx).strRealPath
                If ma_udtExtrasStore(lstPaths.ListIndex).lngType = 3 Then
                    strFile = ma_udtExtrasStore(lstPaths.ListIndex).strPath
                Else
                    strFile = ""
                End If
            End If
        End If
    
        'set the new extra's default properties
        With g_udtExtrasProps
            .strName = lstPaths.List(lstPaths.ListIndex)
            .lngGemusOption = GetExtrasPropsFromDataString(ma_udtExtrasStore(lstPaths.ListIndex).strData)
            .lngType = ma_udtExtrasStore(lstPaths.ListIndex).lngType
            .strFile = strFile
            .strInitPath = strPath
            .strRunnableFile = ma_udtExtrasStore(lstPaths.ListIndex).strRunnableFile
        End With
        frmExtraProps.Show vbModal
        'did the user press cancel?
        If g_udtExtrasProps.strName <> "" Then
            With ma_udtExtrasStore(lstPaths.ListIndex)
                .lngType = g_udtExtrasProps.lngType
                .strData = GetExtrasDataString()
                .strName = g_udtExtrasProps.strName
                .strPath = g_udtExtrasProps.strFile
                .strRunnableFile = g_udtExtrasProps.strRunnableFile
                lstPaths.List(lstPaths.ListIndex) = .strName
            End With
        End If
    End If

End Sub


Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngOldIdx As Long                   'for the "delete" case
    Dim strTempPath As String               'for swapping on the "up" and "down" cases
    Dim lngExtraIdx As Long                 'for deleting/swapping
    Dim udtTempExtra As ExtrasStoreType     'for swapping on the "up" and "down" cases
    Dim a_strTypes() As String
    Dim strError As String
    Dim strInitPath As String
    Dim strLastPath As String
    Dim udtFields As ManualGEMUSFields

    Select Case Button.Key

    Case "new"

        'get the starting path
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "ExtrasPropsPath", strLastPath, True, ""
        If GetMatchingExtrasPath(strLastPath, g_udtExtrasProps.lngPathIdx) <> "" Then
            strError = GetValidExtrasPath(g_udtExtrasProps.lngPathIdx)
            If strError <> "" Then
                MsgBox strError, vbExclamation, gc_strAPPNAME
                Exit Sub
            End If
            strInitPath = ga_udtRegPathExtras(g_udtExtrasProps.lngPathIdx).strRealPath
        Else
            strInitPath = strLastPath
        End If

        'set the new extra's default properties
        With g_udtExtrasProps
            .strName = g_xobjFunc.GetINIMessage("Extras Properties Window", "NewExtra", "New Extra")
            .lngType = 0
            .lngGemusOption = 1
            .strFile = ""
            .strInitPath = strInitPath
            .strRunnableFile = ""
            Call g_xobjGEMUS.GetManualFieldsDefaults(udtFields, g_strGameBaseName)
            .blnGF_TDE = udtFields.blnTDE
            .lngGF_Control = udtFields.lngControl
            .lngGF_PalNTSC = udtFields.lngPalNTSC
            .lngGF_Players = udtFields.lngPlayers
            .strGF_Comment = udtFields.strComment
            .strGF_Notes = udtFields.strNotes
            .strGF_VComment = udtFields.strVComment
            .lngGF_ImageIndex = 0
        End With
        frmExtraProps.Show vbModal
        'did we get a new extra?
        If g_udtExtrasProps.strName <> "" Then
            With g_udtExtrasProps
                Call AddItem(0, .strName, .strFile, .lngType, GetExtrasDataString(), .strRunnableFile)
            End With
            With lstPaths
                .ListIndex = .NewIndex
            End With
        End If
        RaiseEvent DisableAllCheckbox

    Case "delete"

        'delete the current selected extra from the listbox
        With lstPaths
            'delete from array
            For lngExtraIdx = .ListIndex + 1 To UBound(ma_udtExtrasStore)
                ma_udtExtrasStore(lngExtraIdx - 1) = ma_udtExtrasStore(lngExtraIdx)
            Next lngExtraIdx
            If UBound(ma_udtExtrasStore) > 0 Then
                ReDim Preserve ma_udtExtrasStore(UBound(ma_udtExtrasStore) - 1)
            End If
            'delete from listbox
            lngOldIdx = .ListIndex
            .RemoveItem .ListIndex
            If .ListCount = 0 Then
                Call SetButtonsState
                RaiseEvent EnableAllCheckbox
            Else
                If lngOldIdx = .ListCount Then
                    'last extra deleted
                    .ListIndex = lngOldIdx - 1
                Else
                    'force click on same extra
                    .ListIndex = gc_lngNOLISTINDEX
                    .ListIndex = lngOldIdx
                End If
            End If
        End With

    Case "up"

        'swap extra with the one above it
        With lstPaths
            'array
            udtTempExtra = ma_udtExtrasStore(.ListIndex - 1)
            ma_udtExtrasStore(.ListIndex - 1) = ma_udtExtrasStore(.ListIndex)
            ma_udtExtrasStore(.ListIndex) = udtTempExtra
            'listbox
            strTempPath = .List(.ListIndex - 1)
            .List(.ListIndex - 1) = .List(.ListIndex)
            .List(.ListIndex) = strTempPath
            .ListIndex = .ListIndex - 1
        End With

    Case "down"

        'swap extra with the one below it
        With lstPaths
            'array
            udtTempExtra = ma_udtExtrasStore(.ListIndex + 1)
            ma_udtExtrasStore(.ListIndex + 1) = ma_udtExtrasStore(.ListIndex)
            ma_udtExtrasStore(.ListIndex) = udtTempExtra
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
    MaxEntries = 500

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'read the properties from the control's runtime container
    MaxEntries = PropBag.ReadProperty("MaxEntries", 500)

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
    PropBag.WriteProperty "MaxEntries", MaxEntries, 500

End Sub


