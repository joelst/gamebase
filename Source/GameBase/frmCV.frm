VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Views"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmCV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTables 
      Caption         =   "Include Tables"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   2415
      Begin VB.OptionButton optTable 
         Caption         =   "&Both Tables"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Sids Table Only"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Games Table Only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame fraFilters 
      Caption         =   "View Filters"
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   6735
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvFilters 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Field"
            Object.Width           =   2831
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operator"
            Object.Width           =   2725
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Filter"
            Object.Width           =   2990
         EndProperty
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Filter Mode"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   6735
      Begin VB.OptionButton optMode 
         Caption         =   "Any filters can &match"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton optMode 
         Caption         =   "All &filters must match"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   2160
      MaxLength       =   128
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View &Title:"
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCV"
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

Private Const mc_lngMAXFILTERS As Long = 255
Private m_blnFiltersChanged As Boolean
Private m_lngSelectedTableIdx As Long
Private m_lngSelectedModeIdx As Long
Private m_blnfrmCVFilterLoaded As Boolean
Private Function GetOperatorText(ByVal eOperator_in As FilterOperatorEnum, ByVal eClauseType_in As FilterClauseTypeEnum) As String

    '------------------------------------------------
    'gets the OPERATOR text to use in lvFilters
    'depending on the clause and operator type
    '
    'Inputs:    eOperator_in:       The Operator Type
    '           eClauseType_in:     The Clause Type
    '
    'Returns:   The text to use for the operator
    '------------------------------------------------
    
    If eClauseType_in = e_ClauseType_FieldContainsText Then
    
        'FIELD CONTAINS TEXT
        
        If eOperator_in = e_Op_Equal Then
            GetOperatorText = g_strCV_Contains
        
        ElseIf eOperator_in = e_Op_DoesNotEqual Then
            GetOperatorText = g_strCV_NotContains
            
        ElseIf eOperator_in = e_Op_StartsWith Then
            GetOperatorText = g_strCV_StartsWith
        
        Else    'e_Op_EndsWith
            GetOperatorText = g_strCV_EndsWith
            
        End If
      
    ElseIf eClauseType_in = e_ClauseType_FieldIsFilled Then
    
        'FIELD IS FILLED
        
        If eOperator_in = e_Op_Equal Then
            GetOperatorText = g_strCV_Filled
        
        Else    ' e_Op_DoesNotEqual
            GetOperatorText = g_strCV_Empty
 
        End If
    
    ElseIf eClauseType_in = e_ClauseType_FileExists Then
    
        'FILE EXISTS
        
        If eOperator_in = e_Op_Equal Then
            GetOperatorText = g_strCV_Exists
        
        Else    ' e_Op_DoesNotEqual
            GetOperatorText = g_strCV_Missing
 
        End If

    ElseIf eClauseType_in = e_ClauseType_IsDatabaseField Then
    
        'IS DATABASE FIELD
    
        If eOperator_in = e_Op_Equal Then
            GetOperatorText = g_strCV_Equal
        
        ElseIf eOperator_in = e_Op_DoesNotEqual Then
            GetOperatorText = g_strCV_NotEqual
            
        ElseIf eOperator_in = e_Op_IsAfter Then
            GetOperatorText = g_strCV_IsAfter
            
        ElseIf eOperator_in = e_Op_IsBefore Then
            GetOperatorText = g_strCV_IsBefore

        ElseIf eOperator_in = e_Op_IsLessThan Then
            GetOperatorText = g_strCV_LessThan
            
        Else    ' e_Op_IsMoreThan
            GetOperatorText = g_strCV_MoreThan

        End If
        
    Else    'e_ClauseType_NoClause
    
        'NO CLAUSE! - Should NEVER happen
    
        GetOperatorText = "(None)"
        
    End If

End Function
Private Function GetFilterDataText(ByVal eClauseType_in As FilterClauseTypeEnum, strClauseText_in As String, strTable_in As String, strField_in As String) As String

    '------------------------------------------------
    'gets the FILTER DATA text to use in lvFilters
    'depending on the clause and field type
    '
    'Inputs:    eClauseType_in:     The Clause type
    '           strClauseText_in:   The Clause text
    '           strTable_in:        The Clause table
    '           strField_in:        The Clause field
    '
    'Returns:   The text to use for the filter data
    '------------------------------------------------

    Dim strQuery As String      'for getting the linked field data

    If eClauseType_in = e_ClauseType_FieldContainsText Then

        'FIELD CONTAINS TEXT
        GetFilterDataText = """" & strClauseText_in & """"

    ElseIf eClauseType_in = e_ClauseType_FieldIsFilled Then

        'FIELD IS FILLED
        GetFilterDataText = ""

    ElseIf eClauseType_in = e_ClauseType_FileExists Then
    
        'FILE EXISTS
        GetFilterDataText = ""

    ElseIf eClauseType_in = e_ClauseType_IsDatabaseField Then

        'IS DATABASE FIELD
        Select Case strTable_in
            Case "Games"
                
                'Standard Field Data
                Select Case strField_in
                    
                    Case "Rating"
                        GetFilterDataText = ga_strRatingLookupTable(Int(strClauseText_in))
                    
                    Case "Control"
                        GetFilterDataText = ga_strControlLookupTable(Int(strClauseText_in))
                    
                    Case "V_Length"
                        GetFilterDataText = strClauseText_in
                    
                    Case "V_LengthType"
                        GetFilterDataText = ga_strLengthTypeLookupTable(CInt(strClauseText_in))
                    
                    Case "V_Trainers", "PlayersTo", "PlayersFrom"
                        If Int(strClauseText_in) < 0 Then
                            GetFilterDataText = g_strUnknown
                        Else
                            GetFilterDataText = strClauseText_in
                        End If
                    
                    Case "V_PalNTSC"
                        GetFilterDataText = ga_strPalNTSCLookupTable(Int(strClauseText_in))
                    
                    Case "Version"
                        GetFilterDataText = strClauseText_in

                    Case "ReviewRating"
                        GetFilterDataText = strClauseText_in

                    Case Else
                        GetFilterDataText = ga_strYesNoUnknownLookupTable(Abs(Int(strClauseText_in)))
                        
                End Select
                
            Case "Years"
                
                GetFilterDataText = strClauseText_in

            Case Else
                
                'Linked Field Data
                Select Case strField_in
                    Case "PU_Id"
                        strQuery = "Publisher AS FieldText FROM Publishers"
                    Case "PR_Id"
                        strQuery = "Programmer AS FieldText FROM Programmers"
                    Case "MU_Id"
                        strQuery = "Musician AS FieldText FROM Musicians"
                    Case "LA_Id"
                        strQuery = "Language AS FieldText FROM Languages"
                    Case "CR_Id"
                        strQuery = "Cracker AS FieldText FROM Crackers"
                    Case "GE_Id"
                        strQuery = "Genre AS FieldText FROM Genres"
                    Case "PG_Id"
                        strQuery = "ParentGenre AS FieldText FROM PGenres"
                    Case "DI_Id"
                        strQuery = "Difficulty AS FieldText FROM Difficulty"
                    Case "YE_Id"
                        strQuery = "Year AS FieldText FROM Years"
                    Case "AR_Id"
                        strQuery = "Artist AS FieldText FROM Artists"
                    Case "DE_Id"
                        strQuery = "Developer AS FieldText FROM Developers"
                    Case "LI_Id"
                        strQuery = "License AS FieldText FROM Licenses"
                    Case "RA_Id"
                        strQuery = "Rarity AS FieldText FROM Rarities"
                End Select
                strQuery = "SELECT " & strQuery & " WHERE " & strField_in & " = " & strClauseText_in
                'get the field data
                Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
                With g_rsGames
                    .MoveFirst
                    GetFilterDataText = !FieldText
                    .Close
                End With
                Set g_rsGames = Nothing

        End Select

    Else    'e_ClauseType_NoClause

        'NO CLAUSE! - Should NEVER happen

        GetFilterDataText = "(None)"

    End If

End Function

Private Function GetFieldText(strField_in As String, strTable_in As String) As String

    '------------------------------------------------
    'gets the FIELD NAME text to use in lvFilters
    '
    'Inputs:    strField_in:        The Clause field
    '
    'Returns:   The text to use for the filter data
    '------------------------------------------------

    Select Case strField_in
        Case "Name"
            If strTable_in = "Extras" Then
                GetFieldText = g_strField_ExtrasName
            Else
                GetFieldText = g_strField_Name
            End If
        Case "Comment"
            GetFieldText = g_strField_Comment
        Case "Filename", "FA"
            GetFieldText = g_strField_GameFilename
        Case "ScrnShotFilename"
            GetFieldText = g_strField_ScreenShotFilename
        Case "SidFilename", "SA"
            GetFieldText = g_strField_SidFilename
        Case "Year", "YE_Id"
            GetFieldText = g_strField_Year
        Case "Genre", "GE_Id"
            GetFieldText = g_strField_Genre
        Case "ParentGenre", "PG_Id"
            GetFieldText = g_strField_ParentGenre
        Case "Publisher", "PU_Id"
            GetFieldText = g_strField_Publisher
        Case "Artist", "AR_Id"
            GetFieldText = g_strField_Artist
        Case "Developer", "DE_Id"
            GetFieldText = g_strField_Developer
        Case "License", "LI_Id"
            GetFieldText = g_strField_License
        Case "Rarity", "RA_Id"
            GetFieldText = g_strField_Rarity
        Case "Programmer", "PR_Id"
            GetFieldText = g_strField_Programmer
        Case "Musician", "MU_Id"
            GetFieldText = g_strField_Musician
        Case "Language", "LA_Id"
            GetFieldText = g_strField_Language
        Case "Cracker", "CR_Id"
            GetFieldText = g_strField_Cracker
        Case "Difficulty", "DI_Id"
            GetFieldText = g_strField_Difficulty
        Case "HighScore"
            GetFieldText = g_strField_HighScore
        Case "Fav"
            GetFieldText = g_strField_Favourite
        Case "Extras"
            GetFieldText = g_strField_Extras
        Case "Classic"
            GetFieldText = g_strField_Classic
        Case "Rating"
            GetFieldText = g_strField_Rating
        Case "ReviewRating"
            GetFieldText = g_strField_ReviewRating
        Case "V_LoadingScreen"
            GetFieldText = g_strField_LoadingScreen
        Case "V_Playable"
            GetFieldText = g_strField_Playable
        Case "V_Original"
            GetFieldText = g_strField_Original
        Case "V_TitleScreen"
            GetFieldText = g_strField_TitleScreen
        Case "V_HighScoreSaver"
            GetFieldText = g_strField_HighScoreSaver
        Case "V_IncludedDocs"
            GetFieldText = g_strField_IncludedDocs
        Case "V_PalNTSC"
            GetFieldText = g_strField_PALNTSC
        Case "V_TrueDriveEmu"
            GetFieldText = g_strField_TrueDriveEmulation
        Case "V_Length"
            GetFieldText = g_strField_GameLength
        Case "V_LengthType"
            GetFieldText = g_strField_GameLengthType
        Case "V_Trainers"
            GetFieldText = g_strField_Trainers
        Case "PlayersTo", "PlayersFrom"
            GetFieldText = g_strField_NumPlayers
        Case "PlayersSim"
            GetFieldText = g_strField_Players_Simultaneous
        Case "V_Comment"
            GetFieldText = g_strField_VersionComment
        Case "Prequel"
            GetFieldText = g_strField_Prequel
        Case "Sequel"
            GetFieldText = g_strField_Sequel
        Case "Related"
            GetFieldText = g_strField_Related
        Case "CloneOf"
            GetFieldText = g_strField_CloneOf
        Case "MemoText"
            GetFieldText = g_strField_Notes
        Case "Grp"
            GetFieldText = g_strField_MusicianGroup
        Case "Nick"
            GetFieldText = g_strField_MusicianNick
        Case "Photo"
            GetFieldText = g_strField_MusicianPhoto
        Case "Control"
            GetFieldText = g_strField_Control
        Case "Version"
            GetFieldText = g_strField_GameVer
        Case "WebLink_Name"
            GetFieldText = g_strField_WebLinkName
        Case "WebLink_URL"
            GetFieldText = g_strField_WebLinkURL
        Case "V_WebLink_Name"
            GetFieldText = g_strField_VWebLinkName
        Case "V_WebLink_URL"
            GetFieldText = g_strField_VWebLinkURL
    End Select

End Function
Private Sub ListFilters()

    '--------------------------------------------
    'lists the custom view's filters in lvFilters
    '--------------------------------------------

    Dim lngFilterIdx As Long            'loop index
    Dim itmFilter As ListItem           'listitem for adding to lvFilters
    Dim blnAllowSidTable As Boolean     'flag for allowing the sid table in the custom view

    With g_udtCustomView
        If .lngFilterCount > 0 Then
            'show we're busy
            MousePointer = vbHourglass
            DoEvents
            
            'hide and clear listview
            lvFilters.Visible = False
            lvFilters.ListItems.Clear

            'loop through filters adding them to the listview
            For lngFilterIdx = 1 To .lngFilterCount
                With .a_udtFilters(lngFilterIdx)
                    'create listitem object
                    Set itmFilter = lvFilters.ListItems.Add()
                    'fill properties
                    itmFilter.Text = GetFieldText(.strFieldName, .strFieldTable)
                    itmFilter.SubItems(1) = GetOperatorText(.eOperator, .eClauseType)
                    itmFilter.SubItems(2) = GetFilterDataText(.eClauseType, .strClauseData, .strFieldTable, .strFieldName)
                    'release listitem object reference
                    Set itmFilter = Nothing
                End With
            Next lngFilterIdx
            
            'disable the sid/both table includes if needed
            blnAllowSidTable = True
            For lngFilterIdx = 1 To .lngFilterCount
                If .a_udtFilters(lngFilterIdx).strMusicFieldTable = "" Then
                    'no music table defined in one of the filters, so filter can only be for games table
                    blnAllowSidTable = False
                    Exit For
                End If
            Next lngFilterIdx
            If blnAllowSidTable Then
                'enable all option buttons
                optTable(1).Enabled = True
                optTable(2).Enabled = True
            Else
                'disable the sid/both option buttons
                optTable(1).Enabled = False
                optTable(2).Enabled = False
                'move selected table to games if not already selected
                optTable(0).Value = True
            End If
            
            'show listview and resume icon
            lvFilters.Visible = True
            MousePointer = vbDefault
        End If
    End With

End Sub
Private Sub SetFilterButtons()

    '----------------------------------------------------------------------------
    'sets the filter buttons depending on how many filters and if one is selected
    '----------------------------------------------------------------------------
    
    With lvFilters
        cmdAdd.Enabled = (.ListItems.Count < mc_lngMAXFILTERS)
        cmdRemove.Enabled = Not (.SelectedItem Is Nothing)
        cmdEdit.Enabled = cmdRemove.Enabled
        cmdClear.Enabled = (.ListItems.Count > 0)
    End With

End Sub

Private Sub cmdAdd_Click()

    '------------------------------------------
    'shows the filters form to add a new filter
    '------------------------------------------
    
    Dim lngSelectedIdx As Long      'the new filter index

    'initialise the global filter structure for the filter form
    With g_udtCVFilter
        .eClauseType = e_ClauseType_FieldContainsText
        .eOperator = e_Op_Equal
        .strClauseData = ""
        .strFieldName = ""
        .strFieldTable = ""
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
    End With
    
    'show the filter form
    g_strCVFilterWindowTitle = g_xobjFunc.GetINIMessage("CustomViewEditor", "NewFilter", "New Filter")
    m_blnfrmCVFilterLoaded = True
    g_blnEditingCVFilter = False
    frmCVFilter.Show vbModal, Me
    
    'if filter has been returned copy to filters array and signal to save
    If g_udtCVFilter.strFieldName <> "" Then
        With g_udtCustomView
            'add space in filters array
            .lngFilterCount = .lngFilterCount + 1
            ReDim Preserve .a_udtFilters(.lngFilterCount)
            lngSelectedIdx = .lngFilterCount
            'copy across filter
            With g_udtCustomView.a_udtFilters(lngSelectedIdx)
                .eClauseType = g_udtCVFilter.eClauseType
                .eOperator = g_udtCVFilter.eOperator
                .strClauseData = g_udtCVFilter.strClauseData
                .strFieldName = g_udtCVFilter.strFieldName
                .strFieldTable = g_udtCVFilter.strFieldTable
                .strMusicFieldName = g_udtCVFilter.strMusicFieldName
                .strMusicFieldTable = g_udtCVFilter.strMusicFieldTable
            End With
        End With
        'signal to save
        m_blnFiltersChanged = True
        'show the amended list
        Call ListFilters
        'set the filter buttons
        Call SetFilterButtons
        'select the new item
        With lvFilters
            .SelectedItem = lvFilters.ListItems(lngSelectedIdx)
            .SelectedItem.Selected = True
            .SelectedItem.EnsureVisible
        End With
        'enable the ok button
        cmdOK.Enabled = ((txtTitle <> "") And (lvFilters.ListItems.Count > 0))
    End If

    'set focus to the add button
    If cmdAdd.Enabled = True Then
        cmdAdd.SetFocus
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClear_Click()

    '----------------------------------------------
    'clears the list of filters for the custom view
    '----------------------------------------------

    'clear from listview
    lvFilters.ListItems.Clear
    
    'enable all tables
    optTable(1).Enabled = True
    optTable(2).Enabled = True

    'clear from custom view structure
    ReDim g_udtCustomView.a_udtFilters(0)
    g_udtCustomView.lngFilterCount = 0

    'set filter buttons
    Call SetFilterButtons
    
    'disable ok button
    cmdOK.Enabled = False
    
    'set focus to the add button
    If cmdAdd.Enabled Then
        cmdAdd.SetFocus
    End If

    'set save flag
    m_blnFiltersChanged = True

End Sub

Private Sub cmdEdit_Click()

    '------------------------------------------------------------
    'shows the filters form to edit the currently selected filter
    '------------------------------------------------------------
    
    Dim lngSelectedIdx As Long      'the selected filter
    
    'get the filter to edit
    If Not (lvFilters.SelectedItem Is Nothing) Then
        lngSelectedIdx = lvFilters.SelectedItem.Index
    Else
        'no item to edit!
        Call SetFilterButtons
        lvFilters.SetFocus
        Exit Sub
    End If

    'copy the filter into the global filter structure for the filter form
    With g_udtCustomView.a_udtFilters(lngSelectedIdx)
        g_udtCVFilter.eClauseType = .eClauseType
        g_udtCVFilter.eOperator = .eOperator
        g_udtCVFilter.strClauseData = .strClauseData
        g_udtCVFilter.strFieldName = .strFieldName
        g_udtCVFilter.strFieldTable = .strFieldTable
        g_udtCVFilter.strMusicFieldName = .strMusicFieldName
        g_udtCVFilter.strMusicFieldTable = .strMusicFieldTable
    End With
    
    'show the filter form
    g_strCVFilterWindowTitle = g_xobjFunc.GetINIMessage("CustomViewEditor", "EditFilter", "Edit Filter")
    m_blnfrmCVFilterLoaded = True
    g_blnEditingCVFilter = True
    frmCVFilter.Show vbModal, Me
    
    'if filter has changed copy back to filters array and signal to save
    If g_udtCVFilter.strFieldName <> "" Then
        'only need to copy across and save if the view has changed
        With g_udtCustomView.a_udtFilters(lngSelectedIdx)
            If .eClauseType <> g_udtCVFilter.eClauseType Or _
            .eOperator <> g_udtCVFilter.eOperator Or _
            .strClauseData <> g_udtCVFilter.strClauseData Or _
            .strFieldName <> g_udtCVFilter.strFieldName Or _
            .strFieldTable <> g_udtCVFilter.strFieldTable Or _
            .strMusicFieldName <> g_udtCVFilter.strMusicFieldName Or _
            .strMusicFieldTable <> g_udtCVFilter.strMusicFieldTable Then
                'view changed - copy across and signal to save!
                .eClauseType = g_udtCVFilter.eClauseType
                .eOperator = g_udtCVFilter.eOperator
                .strClauseData = g_udtCVFilter.strClauseData
                .strFieldName = g_udtCVFilter.strFieldName
                .strFieldTable = g_udtCVFilter.strFieldTable
                .strMusicFieldName = g_udtCVFilter.strMusicFieldName
                .strMusicFieldTable = g_udtCVFilter.strMusicFieldTable
                m_blnFiltersChanged = True
                'show the amended list
                Call ListFilters
                'select the changed item
                With lvFilters
                    .SelectedItem = lvFilters.ListItems(lngSelectedIdx)
                    .SelectedItem.Selected = True
                    .SelectedItem.EnsureVisible
                End With
            End If
        End With
    End If

    'set the ok button
    cmdOK.Enabled = ((txtTitle <> "") And (lvFilters.ListItems.Count > 0))

    'set focus to the listview again
    lvFilters.SetFocus

End Sub

Private Sub cmdOK_Click()

    'do we want to save the custom view? (has it changed?)
    With g_udtCustomView
        If (txtTitle <> .strTitle) Or (m_lngSelectedTableIdx <> .eIncludeMusicTable) Or (m_lngSelectedModeIdx <> .eFilterMode) Or (m_blnFiltersChanged) Then
            'copy details to view structure
            .strTitle = txtTitle
            .eIncludeMusicTable = m_lngSelectedTableIdx
            .eFilterMode = m_lngSelectedModeIdx
            'set flag to save the changed view
            g_blnSaveCustomView = True
        End If
    End With

    'return to frmMain
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    
    '-----------------------------------------
    'removes the selected filter from the list
    '-----------------------------------------

    Dim lngSelectedIdx As Long      'the filter to delete
    Dim lngFilterIdx As Long        'loop index

    'get the filter to remove
    If Not (lvFilters.SelectedItem Is Nothing) Then
        lngSelectedIdx = lvFilters.SelectedItem.Index
    Else
        'no item to delete!
        Call SetFilterButtons
        lvFilters.SetFocus
        Exit Sub
    End If

    'remove from custom view structure
    With g_udtCustomView
        'move filters after the selected one back one index
        For lngFilterIdx = lngSelectedIdx + 1 To .lngFilterCount
            .a_udtFilters(lngFilterIdx - 1) = .a_udtFilters(lngFilterIdx)
        Next lngFilterIdx
        'redim minus the removed filter
        .lngFilterCount = .lngFilterCount - 1
        ReDim Preserve .a_udtFilters(.lngFilterCount)
    End With

    If g_udtCustomView.lngFilterCount = 0 Then
        'no filters
        lvFilters.ListItems.Clear
        optTable(1).Enabled = True
        optTable(2).Enabled = True
        Call SetFilterButtons
        cmdOK.Enabled = False
        If cmdAdd.Enabled Then
            cmdAdd.SetFocus
        End If
    Else
        'list filters again
        Call ListFilters
        Call SetFilterButtons
        cmdOK.Enabled = ((txtTitle <> "") And (lvFilters.ListItems.Count > 0))
        'set focus to the filter that took place of the one removed
        With lvFilters
            If lngSelectedIdx > .ListItems.Count Then
                Set .SelectedItem = .ListItems(lngSelectedIdx - 1)
            Else
                Set .SelectedItem = .ListItems(lngSelectedIdx)
            End If
            .SelectedItem.Selected = True
            .SelectedItem.EnsureVisible
            .SetFocus
        End With
    End If
    

    'set save flag
    m_blnFiltersChanged = True

End Sub

Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 4680    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6975     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    'read in the ini file text
    With g_xobjFunc
        cmdAdd.Caption = .GetINIMessage("CustomViewEditor", "Button_Add", "&Add")
        cmdEdit.Caption = .GetINIMessage("CustomViewEditor", "Button_Edit", "&Edit")
        cmdRemove.Caption = .GetINIMessage("CustomViewEditor", "Button_Remove", "&Remove")
        cmdClear.Caption = .GetINIMessage("CustomViewEditor", "Button_Clear", "C&lear")
        lblTitle.Caption = .GetINIMessage("CustomViewEditor", "Title", "View &Title") & ":"
        fraMode.Caption = .GetINIMessage("CustomViewEditor", "FilterMode", "Filter Mode")
        fraFilters.Caption = .GetINIMessage("CustomViewEditor", "ViewFilters", "View Filters")
        optMode(0).Caption = .GetINIMessage("CustomViewEditor", "AllFilters", "All &filters must match")
        optMode(1).Caption = .GetINIMessage("CustomViewEditor", "AnyFilters", "A&ny filter can match")
        fraTables.Caption = .GetINIMessage("CustomViewEditor", "IncludeTables", "Include Tables")
        optTable(0).Caption = .GetINIMessage("CustomViewEditor", "GamesTable", "&Games Table Only")
        optTable(1).Caption = .GetINIMessage("CustomViewEditor", "MusicTable", "&Music Table Only")
        optTable(2).Caption = .GetINIMessage("CustomViewEditor", "BothTables", "&Both Tables")
        g_strCVFilter_Field = .GetINIMessage("CustomViewEditor", "Field", "Field")
        g_strCVFilter_Operator = .GetINIMessage("CustomViewEditor", "Operator", "Operator")
        g_strCVFilter_FilterData = .GetINIMessage("CustomViewEditor", "Filter", "Filter Data")
        'filter operators
        g_strCV_Contains = .GetINIMessage("CustomViewEditor", "Contains", "contains text")
        g_strCV_NotContains = .GetINIMessage("CustomViewEditor", "NotContains", "excludes text")
        g_strCV_StartsWith = .GetINIMessage("CustomViewEditor", "StartsWith", "starts with text")
        g_strCV_EndsWith = .GetINIMessage("CustomViewEditor", "EndsWith", "ends with text")
        g_strCV_Filled = .GetINIMessage("CustomViewEditor", "FieldIsFilled", "field is filled")
        g_strCV_Empty = .GetINIMessage("CustomViewEditor", "FieldIsEmpty", "field is empty")
        g_strCV_Exists = .GetINIMessage("CustomViewEditor", "FileExists", "file exists")
        g_strCV_Missing = .GetINIMessage("CustomViewEditor", "FileMissing", "file does not exist")
        g_strCV_Equal = .GetINIMessage("CustomViewEditor", "Equals", "is")
        g_strCV_NotEqual = .GetINIMessage("CustomViewEditor", "DoesNotEqual", "is not")
        g_strCV_IsBefore = .GetINIMessage("CustomViewEditor", "IsBefore", "is before")
        g_strCV_IsAfter = .GetINIMessage("CustomViewEditor", "IsAfter", "is after")
        g_strCV_LessThan = .GetINIMessage("CustomViewEditor", "LessThan", "is Less than")
        g_strCV_MoreThan = .GetINIMessage("CustomViewEditor", "MoreThan", "is more than")
        'filter data
        g_strCVFD_GameFilename = .GetINIMessage("CustomViewEditor", "GameFilename", "Game Filename")
        g_strCVFD_ScrnShotFilename = .GetINIMessage("CustomViewEditor", "PictureFilename", "Picture Filename")
        g_strCVFD_SidFilename = .GetINIMessage("CustomViewEditor", "MusicFilename", "Music Filename")
    End With

    'set the button text and window title
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    Caption = g_strCustomViewWindowTitle
    
    'set the filter columnheader text
    lvFilters.ColumnHeaders(1).Text = g_strCVFilter_Field
    lvFilters.ColumnHeaders(2).Text = g_strCVFilter_Operator
    lvFilters.ColumnHeaders(3).Text = g_strCVFilter_FilterData

    'assume user will press cancel button
    g_blnSaveCustomView = False

    'fill in the view properties
    With g_udtCustomView
        txtTitle = .strTitle
        optMode(.eFilterMode).Value = True
        optTable(.eIncludeMusicTable).Value = True
        'disable the ok button if there are no filters
        cmdOK.Enabled = (.lngFilterCount > 0)
        'list the filters and set the filter buttons
        If .lngFilterCount > 0 Then
            Call ListFilters
            Call SetFilterButtons
        End If

    End With
    
    'nothing changed yet
    m_blnFiltersChanged = False
    
    'haven't loaded the filter form yet
    m_blnfrmCVFilterLoaded = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If m_blnfrmCVFilterLoaded Then
        Unload frmCVFilter
    End If

End Sub


Private Sub lvFilters_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    'enable/disable the filter buttons
    Call SetFilterButtons

End Sub


Private Sub optMode_Click(Index As Integer)

    'store the selected mode
    m_lngSelectedModeIdx = Index
    
End Sub

Private Sub optTable_Click(Index As Integer)

    'store the selected table include
    m_lngSelectedTableIdx = Index

End Sub

Private Sub txtTitle_Change()

    cmdOK.Enabled = ((txtTitle <> "") And (lvFilters.ListItems.Count > 0))

End Sub

Private Sub txtTitle_GotFocus()

    With txtTitle
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


