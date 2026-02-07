VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlExeSelector 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox lstPaths 
      Height          =   1425
      ItemData        =   "ctlExeSelector.ctx":0000
      Left            =   0
      List            =   "ctlExeSelector.ctx":0002
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
            Picture         =   "ctlExeSelector.ctx":0004
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExeSelector.ctx":00A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExeSelector.ctx":0134
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlExeSelector.ctx":01C6
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
Attribute VB_Name = "ctlExeSelector"
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
Private m_lngMaxEntries As Long                 'maximum number of paths allowed
Private Type SidPlayerRegType
    strName As String
    strExe As String
    strCLP As String
    blnUseShort As Boolean
    a_strTypes() As String
    lngTotalTypes As Long
End Type
Private ma_udtSidPlayers() As SidPlayerRegType     'the sid player data

Public Sub AddItem(strName_in As String, strExe_in As String, strCLP_in As String, ByVal blnUseShort_in As Boolean, a_strTypes_in() As String, ByVal lngTotalTypes_in As Long)

    'add a sidplayer to the list

    ReDim Preserve ma_udtSidPlayers(lstPaths.ListCount)

    With ma_udtSidPlayers(lstPaths.ListCount)
        .blnUseShort = blnUseShort_in
        .strCLP = strCLP_in
        .strExe = strExe_in
        .strName = strName_in
        .a_strTypes() = a_strTypes_in()
        .lngTotalTypes = lngTotalTypes_in
    End With

    lstPaths.AddItem ma_udtSidPlayers(lstPaths.ListCount).strName

End Sub

Public Sub GetItem(ByVal intIndex As Integer, strName_out As String, strExe_out As String, strCLP_out As String, blnUseShort_out As Boolean, a_strTypes_out() As String, lngTotalTypes_out As Long)

    '----------------------------------------------------------------------------
    'gets the info for the sid player at intIndex in the listbox
    '
    'Inputs:    intIndex:           The index of the sidplayer in the listbox
    '
    'Outputs:   strName_out:        Var to receive the player name
    '           strExe_out:         Var to receive the player exe
    '           strCLP_out:         Var to receive the player command line params
    '           blnUseShort_out:    Var to receive the Use Short Filenames flag
    '----------------------------------------------------------------------------

    'copy across the values
    With ma_udtSidPlayers(intIndex)
        strName_out = .strName
        strExe_out = .strExe
        strCLP_out = .strCLP
        blnUseShort_out = .blnUseShort
        a_strTypes_out() = .a_strTypes()
        lngTotalTypes_out = .lngTotalTypes
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

Private Sub lstPaths_Click()

    'set the state of the buttons
    Call SetButtonsState
        
End Sub

Private Sub lstPaths_DblClick()

    If lstPaths.ListIndex > gc_lngNOLISTINDEX Then
        'set the new sid player properties to edit
        With ma_udtSidPlayers(lstPaths.ListIndex)
            g_udtSidPlayerProps.strName = .strName
            g_udtSidPlayerProps.strFullPath = .strExe
            g_udtSidPlayerProps.strCommandLine = .strCLP
            g_udtSidPlayerProps.blnUseShort = .blnUseShort
            g_udtSidPlayerProps.a_strTypes() = .a_strTypes()
            g_udtSidPlayerProps.lngTotalTypes = .lngTotalTypes
        End With
        frmSidPlayerProps.Show vbModal
        'did the user press cancel?
        If g_udtSidPlayerProps.strName <> "" Then
            With ma_udtSidPlayers(lstPaths.ListIndex)
                .blnUseShort = g_udtSidPlayerProps.blnUseShort
                .strCLP = g_udtSidPlayerProps.strCommandLine
                .strExe = g_udtSidPlayerProps.strFullPath
                .strName = g_udtSidPlayerProps.strName
                .a_strTypes() = g_udtSidPlayerProps.a_strTypes()
                .lngTotalTypes = g_udtSidPlayerProps.lngTotalTypes
                lstPaths.List(lstPaths.ListIndex) = .strName
            End With
        End If
    End If

End Sub


Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngOldIdx As Long                   'for the "delete" case
    Dim strTempPath As String               'for swapping on the "up" and "down" cases
    Dim lngPlayerIdx As Long                'for deleting/swapping
    Dim udtTempPlayer As SidPlayerRegType   'for swapping on the "up" and "down" cases
    Dim a_strTypes() As String
    

    Select Case Button.Key

    Case "new"

        'get the new sid player properties
        With g_udtSidPlayerProps
            .strName = ""
            .strFullPath = ""
            .strCommandLine = "%musicfile%"
            .blnUseShort = False
            .lngTotalTypes = 0
            ReDim a_strTypes(0)
            .a_strTypes = a_strTypes()
        End With
        frmSidPlayerProps.Show vbModal
        'did we get a new player?
        If g_udtSidPlayerProps.strName <> "" Then
            With g_udtSidPlayerProps
                Call AddItem(.strName, .strFullPath, .strCommandLine, .blnUseShort, .a_strTypes(), .lngTotalTypes)
            End With
            With lstPaths
                .ListIndex = .NewIndex
            End With
        End If

    Case "delete"

        'delete the current selected path from the listbox
        With lstPaths
            'delete from array
            For lngPlayerIdx = .ListIndex + 1 To UBound(ma_udtSidPlayers)
                ma_udtSidPlayers(lngPlayerIdx - 1) = ma_udtSidPlayers(lngPlayerIdx)
            Next lngPlayerIdx
            If UBound(ma_udtSidPlayers) > 0 Then
                ReDim Preserve ma_udtSidPlayers(UBound(ma_udtSidPlayers) - 1)
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
                    .ListIndex = gc_lngNOLISTINDEX
                    .ListIndex = lngOldIdx
                End If
            End If
        End With

    Case "up"

        'swap player with the one above it
        With lstPaths
            'array
            udtTempPlayer = ma_udtSidPlayers(.ListIndex - 1)
            ma_udtSidPlayers(.ListIndex - 1) = ma_udtSidPlayers(.ListIndex)
            ma_udtSidPlayers(.ListIndex) = udtTempPlayer
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
            udtTempPlayer = ma_udtSidPlayers(.ListIndex + 1)
            ma_udtSidPlayers(.ListIndex + 1) = ma_udtSidPlayers(.ListIndex)
            ma_udtSidPlayers(.ListIndex) = udtTempPlayer
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
    MaxEntries = PropBag.ReadProperty("MaxEntries", 100)

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
    PropBag.WriteProperty "MaxEntries", MaxEntries, 100

End Sub


