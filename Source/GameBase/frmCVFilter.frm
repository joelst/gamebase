VERSION 5.00
Begin VB.Form frmCVFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom View Filter"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmCVFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstFilterData 
      Height          =   3180
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   4695
   End
   Begin VB.ListBox lstField 
      Height          =   3180
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstOperator 
      Height          =   3180
      ItemData        =   "frmCVFilter.frx":0442
      Left            =   3120
      List            =   "frmCVFilter.frx":0444
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtFilterData 
      Height          =   285
      Left            =   6120
      MaxLength       =   64
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblFilterData 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Data:"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblOperator 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Caption         =   "Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmCVFilter"
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

Private Enum FilterSelectorEnum
    e_FilterSelector_YesNo
    e_FilterSelector_Text
    e_FilterSelector_DBTable
    e_FilterSelector_None
    e_FilterSelector_Other
    e_FilterSelector_YesNoUnknown
End Enum
Private Type FieldOpType
    strDisplayText As String
    eOperator As FilterOperatorEnum
    eClauseType As FilterClauseTypeEnum
    strFieldTable As String
    strFieldName As String
    strMusicFieldTable As String
    strMusicFieldName As String
    eFilterSelector As FilterSelectorEnum
End Type
Private Type FieldRulesType
    strFieldName As String
    udtAllowedOps() As FieldOpType
    lngMaxOps As Long
End Type
Private ma_udtFields() As FieldRulesType

Private Const mc_strILLEGALCHARS As String = """*[]"
Private Const mc_strILLEGALCHARSMSG As String = """   [   ]   *"

Private Enum CVFilterFieldPositionEnum
    e_CVFilterFieldPos_Name
    e_CVFilterFieldPos_Publisher
    e_CVFilterFieldPos_Developer
    e_CVFilterFieldPos_Year
    e_CVFilterFieldPos_Programmer
    e_CVFilterFieldPos_Musician
    e_CVFilterFieldPos_Artist
    e_CVFilterFieldPos_License
    e_CVFilterFieldPos_ReviewRating
    e_CVFilterFieldPos_Genre
    e_CVFilterFieldPos_ParentGenre
    e_CVFilterFieldPos_Rarity
    e_CVFilterFieldPos_NumPlayers
    e_CVFilterFieldPos_SimulPlay
    e_CVFilterFieldPos_Language
    e_CVFilterFieldPos_WebLinkName
    e_CVFilterFieldPos_WebLinkURL
    e_CVFilterFieldPos_Prequel
    e_CVFilterFieldPos_Sequel
    e_CVFilterFieldPos_Related
    e_CVFilterFieldPos_CloneOf
    e_CVFilterFieldPos_Comment
    e_CVFilterFieldPos_Filename
    e_CVFilterFieldPos_SidFilename
    e_CVFilterFieldPos_ScrnShotFilename
    e_CVFilterFieldPos_Cracker
    e_CVFilterFieldPos_NumTrainers
    e_CVFilterFieldPos_Playable
    e_CVFilterFieldPos_Original
    e_CVFilterFieldPos_TitleScreen
    e_CVFilterFieldPos_LoadingScreen
    e_CVFilterFieldPos_HighScoreSaver
    e_CVFilterFieldPos_IncludedDocs
    e_CVFilterFieldPos_PalNTSC
    e_CVFilterFieldPos_TrueDriveEmulation
    e_CVFilterFieldPos_GameLength
    e_CVFilterFieldPos_GameLengthType
    e_CVFilterFieldPos_VWebLinkName
    e_CVFilterFieldPos_VWebLinkURL
    e_CVFilterFieldPos_VersionComment
    e_CVFilterFieldPos_GameVer
    e_CVFilterFieldPos_Difficulty
    e_CVFilterFieldPos_HighScore
    e_CVFilterFieldPos_Favourite
    e_CVFilterFieldPos_Extras
    e_CVFilterFieldPos_Extras2
    e_CVFilterFieldPos_Classic
    e_CVFilterFieldPos_Rating
    e_CVFilterFieldPos_Control
    e_CVFilterFieldPos_MemoText
    e_CVFilterFieldPos_MusicianGroup
    e_CVFilterFieldPos_MusicianNick
    e_CVFilterFieldPos_MusicianPhoto
    e_CVFilterFieldPos_Adult
End Enum

Private Function GetLinkFieldName(strLinkField_in As String)

    '----------------------------------------------------
    'gets the actual field name from the link field name
    'so the filterdata listbox can be filled correctly
    'wen retrieving tables from the database
    '
    'Inputs:    strLinkField_in:    The link field name
    '
    'Returns:   The proper field name from the linked one
    '----------------------------------------------------
    
    Select Case strLinkField_in
    
        Case "PU_Id"
            GetLinkFieldName = "Publisher"
        
        Case "PR_Id"
            GetLinkFieldName = "Programmer"
        
        Case "MU_Id"
            GetLinkFieldName = "Musician"
        
        Case "LA_Id"
            GetLinkFieldName = "Language"
        
        Case "GE_Id"
            GetLinkFieldName = "Genre"
        
        Case "PG_Id"
            GetLinkFieldName = "ParentGenre"
        
        Case "CR_Id"
            GetLinkFieldName = "Cracker"
        
        Case "DI_Id"
            GetLinkFieldName = "Difficulty"

        Case "AR_Id"
            GetLinkFieldName = "Artist"

        Case "DE_Id"
            GetLinkFieldName = "Developer"

        Case "LI_Id"
            GetLinkFieldName = "License"

        Case "RA_Id"
            GetLinkFieldName = "Rarity"

    End Select

End Function

Private Sub ListFilterData()

    '------------------------------------------------------------------------
    'lists the filter selections depending on the field and operator selected
    '------------------------------------------------------------------------
    
    Dim strQuery As String      'for fetching recordsets from the database
    Dim strLinkField As String  'proper fieldname from link fieldname
    Dim fldFilter As Field      'for speed filling the listbox
    Dim fldFilter2 As Field     'for speed filling the listbox
    Dim fldID As Field          'for storing in itemdata property of listbox
    Dim lngIdx As Long          'loop index
    
    'show we're busy
    MousePointer = vbHourglass
    DoEvents
    
    'hide text box, listbox and label, and clear the listbox
    lblFilterData.Visible = False
    txtFilterData.Visible = False
    lstFilterData.Visible = False
    lstFilterData.Clear
    
    With ma_udtFields(lstField.ListIndex).udtAllowedOps(lstOperator.ListIndex)
    
        If .eFilterSelector = e_FilterSelector_DBTable Then
            
            '''''''''''''''
            'DATABASE TABLE
            '''''''''''''''
        
            'get the query and fill the listbox
            strLinkField = GetLinkFieldName(.strFieldName)
            strQuery = "SELECT " & strLinkField & " AS Fld, " & .strFieldName & " AS ID FROM " & .strFieldTable & " WHERE " & strLinkField & " <> """" ORDER BY " & strLinkField
            Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
            Set fldFilter = g_rsGames!fld
            Set fldID = g_rsGames!ID
            With lstFilterData
                g_rsGames.MoveFirst
                Do Until g_rsGames.EOF
                    .AddItem fldFilter
                    .ItemData(.NewIndex) = fldID
                    g_rsGames.MoveNext
                Loop
                g_rsGames.Close
            End With
            Set fldFilter = Nothing
            Set fldID = Nothing
            Set g_rsGames = Nothing
            
            'show the label and listbox
            lblFilterData.Visible = True
            lstFilterData.Visible = True
        
        ElseIf .eFilterSelector = e_FilterSelector_Text Then
        
            ''''''''''''
            'FILTER TEXT
            ''''''''''''
            
            'show the label and textbox
            lblFilterData.Visible = True
            txtFilterData.Visible = True
        
        ElseIf .eFilterSelector = e_FilterSelector_None Then
        
            '''''''''''''''
            'NO FILTER DATA
            '''''''''''''''
            
            'don't show anything!
        
        ElseIf .eFilterSelector = e_FilterSelector_YesNo Then
        
            '''''''''''''
            'YES NO FIELD
            '''''''''''''
            
            'fill the listbox with "Yes" and "No" entries
            lstFilterData.AddItem ga_strYesNoLookupTable(1) 'Yes
            lstFilterData.ItemData(0) = -1
            lstFilterData.AddItem ga_strYesNoLookupTable(0) 'No
            lstFilterData.ItemData(1) = 0
            
            'show the label and listbox
            lblFilterData.Visible = True
            lstFilterData.Visible = True
    
        ElseIf .eFilterSelector = e_FilterSelector_YesNoUnknown Then
        
            '''''''''''''''''''''
            'YES NO UNKNOWN FIELD
            '''''''''''''''''''''
            
            'fill the listbox with "Yes" and "No" entries
            lstFilterData.AddItem ga_strYesNoUnknownLookupTable(1) 'Yes
            lstFilterData.ItemData(0) = 1
            lstFilterData.AddItem ga_strYesNoUnknownLookupTable(0) 'No
            lstFilterData.ItemData(1) = 0
            lstFilterData.AddItem ga_strYesNoUnknownLookupTable(2) 'Unknown
            lstFilterData.ItemData(2) = 2

            'show the label and listbox
            lblFilterData.Visible = True
            lstFilterData.Visible = True
    
        Else
        
            ''''''
            'OTHER
            ''''''
            
            If .strFieldName = "YE_Id" Then
            
                '''''''''''''''''''''''''
                'YEARS EQUAL OR NOT EQUAL
                '''''''''''''''''''''''''
                
                With lstFilterData
                    For lngIdx = 1 To gc_lngYEARLOOKUPMAX
                        .AddItem ga_strYearLookupTable(lngIdx)
                        .ItemData(.NewIndex) = lngIdx
                    Next lngIdx
                End With
                
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            ElseIf .strFieldName = "Year" Then
                
                If .eOperator = e_Op_IsAfter Then
                
                    '''''''''''''''
                    'YEARS IS AFTER
                    '''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 10 To gc_lngYEARLOOKUPMAX - 1
                            .AddItem ga_strYearLookupTable(lngIdx)
                            .ItemData(.NewIndex) = ga_strYearLookupTable(lngIdx)
                        Next lngIdx
                    End With
                
                Else
                
                    ''''''''''''''''
                    'YEARS IS BEFORE
                    ''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 11 To gc_lngYEARLOOKUPMAX
                            .AddItem ga_strYearLookupTable(lngIdx)
                            .ItemData(.NewIndex) = ga_strYearLookupTable(lngIdx)
                        Next lngIdx
                    End With
                
                End If
                
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True

            ElseIf .strFieldName = "ReviewRating" Then
            
                If .eOperator = e_Op_Equal Or .eOperator = e_Op_DoesNotEqual Then
                
                    ''''''''''''''''''''''''''''
                    'TRAINERS EQUAL OR NOT EQUAL
                    ''''''''''''''''''''''''''''
                    
                    With lstFilterData
                        .AddItem "(" & g_strUnknown & ")"
                        .ItemData(.NewIndex) = -1
                        For lngIdx = 0 To 100
                            .AddItem lngIdx & "%"
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With

                ElseIf .eOperator = e_Op_IsMoreThan Then
                
                    '''''''''''''''''''
                    'TRAINERS MORE THAN
                    '''''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 0 To 99
                            .AddItem lngIdx & "%"
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                
                Else    'e_Op_IsLessThan
                
                    '''''''''''''''''''
                    'TRAINERS LESS THAN
                    '''''''''''''''''''
                    With lstFilterData
                        For lngIdx = 1 To 100
                            .AddItem lngIdx & "%"
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                            
                End If

                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True

            ElseIf .strFieldName = "GE_Id" Then
            
                '''''''''''''''''''''''''
                'GENRE EQUAL OR NOT EQUAL
                '''''''''''''''''''''''''
                
                'get the query and fill the listbox
                strQuery = "SELECT Genre, ParentGenre, GE_ID FROM PGenres INNER JOIN Genres ON PGenres.PG_Id = Genres.PG_Id ORDER BY PGenres.ParentGenre, Genres.Genre"
                Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
                Set fldFilter = g_rsGames!Genre
                Set fldFilter2 = g_rsGames!ParentGenre
                Set fldID = g_rsGames!GE_Id
                With lstFilterData
                    g_rsGames.MoveFirst
                    Do Until g_rsGames.EOF
                        If Len(fldFilter2) > 0 Then
                            .AddItem fldFilter2 & " - " & fldFilter
                        Else
                            .AddItem fldFilter
                        End If
                        .ItemData(.NewIndex) = fldID
                        g_rsGames.MoveNext
                    Loop
                    g_rsGames.Close
                End With
                Set fldFilter = Nothing
                Set fldFilter2 = Nothing
                Set fldID = Nothing
                Set g_rsGames = Nothing
                
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            ElseIf .strFieldName = "PlayersFrom" Then
                
                If .eOperator = e_Op_Equal Or .eOperator = e_Op_DoesNotEqual Then
                
                    '''''''''''''''''''''''''''
                    'PLAYERS EQUAL OR NOT EQUAL
                    '''''''''''''''''''''''''''
                    
                    With lstFilterData
                        .AddItem "(" & g_strUnknown & ")"
                        .ItemData(.NewIndex) = -1
                        For lngIdx = 0 To 10
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                
                Else    'e_Op_IsMoreThan
                
                    ''''''''''''''''''
                    'PLAYERS MORE THAN
                    ''''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 0 To 10
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With

                End If
                
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            ElseIf .strFieldName = "PlayersTo" Then
            
                ''''''''''''''''''
                'PLAYERS LESS THAN
                ''''''''''''''''''

                With lstFilterData
                    For lngIdx = 1 To 10
                        .AddItem lngIdx
                        .ItemData(.NewIndex) = lngIdx
                    Next lngIdx
                End With
                
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
                
            ElseIf .strFieldName = "V_Trainers" Then
            
                If .eOperator = e_Op_Equal Or .eOperator = e_Op_DoesNotEqual Then
                
                    ''''''''''''''''''''''''''''
                    'TRAINERS EQUAL OR NOT EQUAL
                    ''''''''''''''''''''''''''''
                    
                    With lstFilterData
                        .AddItem "(" & g_strUnknown & ")"
                        .ItemData(.NewIndex) = -1
                        For lngIdx = 0 To 20
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With

                ElseIf .eOperator = e_Op_IsMoreThan Then
                
                    '''''''''''''''''''
                    'TRAINERS MORE THAN
                    '''''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 0 To 20
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                
                Else    'e_Op_IsLessThan
                
                    '''''''''''''''''''
                    'TRAINERS LESS THAN
                    '''''''''''''''''''
                    With lstFilterData
                        For lngIdx = 1 To 20
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                            
                End If
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
                
            ElseIf .strFieldName = "V_Length" Then
            
                With lstFilterData
                    For lngIdx = 0 To 1000
                        .AddItem lngIdx
                        .ItemData(.NewIndex) = lngIdx
                    Next lngIdx
                End With
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            ElseIf .strFieldName = "V_LengthType" Then
            
                With lstFilterData
                    For lngIdx = 0 To 19
                        .AddItem ga_strLengthTypeLookupTable(lngIdx)
                        .ItemData(.NewIndex) = lngIdx
                    Next lngIdx
                End With
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            ElseIf .strFieldName = "Rating" Then
            
                If .eOperator = e_Op_Equal Or .eOperator = e_Op_DoesNotEqual Then
                
                    ''''''''''''''''''''''''''
                    'RATING EQUAL OR NOT EQUAL
                    ''''''''''''''''''''''''''
                    
                    With lstFilterData
                        .AddItem ga_strRatingLookupTable(0)
                        .ItemData(.NewIndex) = 0
                        For lngIdx = 1 To 5
                            .AddItem ga_strRatingLookupTable(lngIdx)
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With

                ElseIf .eOperator = e_Op_IsMoreThan Then
                
                    '''''''''''''''''
                    'RATING MORE THAN
                    '''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 1 To 4
                            .AddItem ga_strRatingLookupTable(lngIdx)
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                
                Else    'e_Op_IsLessThan
                
                    '''''''''''''''''
                    'RATING LESS THAN
                    '''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 2 To 5
                            .AddItem ga_strRatingLookupTable(lngIdx)
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                            
                End If
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
     
            ElseIf .strFieldName = "Control" Then

                '''''''''''''''''''''''''''
                'CONTROL EQUAL OR NOT EQUAL
                '''''''''''''''''''''''''''

                With lstFilterData
                    .AddItem ga_strControlLookupTable(0)
                    .ItemData(.NewIndex) = 0
                    For lngIdx = 1 To 8
                        .AddItem ga_strControlLookupTable(lngIdx)
                        .ItemData(.NewIndex) = lngIdx
                    Next lngIdx
                End With

                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True

            ElseIf .strFieldName = "V_PalNTSC" Then
                    
                ''''''''''''''''''''''''''''''''''''
                'VERSION PAL/NTSC EQUAL OR NOT EQUAL
                ''''''''''''''''''''''''''''''''''''
                
                With lstFilterData
                    .AddItem ga_strPalNTSCLookupTable(0)
                    .AddItem ga_strPalNTSCLookupTable(1)
                    .AddItem ga_strPalNTSCLookupTable(2)
                    .AddItem ga_strPalNTSCLookupTable(3)
                    .ItemData(0) = 0
                    .ItemData(1) = 1
                    .ItemData(2) = 2
                    .ItemData(3) = 3
                End With
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True

            ElseIf .strFieldName = "Version" Then
            
                If .eOperator = e_Op_Equal Or .eOperator = e_Op_DoesNotEqual Then
                
                    '''''''''''''''''''''''''''
                    'VERSION EQUAL OR NOT EQUAL
                    '''''''''''''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 0 To 999
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With

                ElseIf .eOperator = e_Op_IsMoreThan Then
                
                    ''''''''''''''''''
                    'VERSION MORE THAN
                    ''''''''''''''''''
                    
                    With lstFilterData
                        For lngIdx = 0 To 998
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                
                Else    'e_Op_IsLessThan
                
                    ''''''''''''''''''
                    'VERSION LESS THAN
                    ''''''''''''''''''
                    With lstFilterData
                        For lngIdx = 1 To 999
                            .AddItem lngIdx
                            .ItemData(.NewIndex) = lngIdx
                        Next lngIdx
                    End With
                            
                End If
            
                'show the label and listbox
                lblFilterData.Visible = True
                lstFilterData.Visible = True
            
            Else
            
            End If
            
        End If
    
    End With
    
    'finished
    MousePointer = vbDefault

End Sub
Private Sub LoadFields()

    '------------------------------------------
    'Loads the fields into the Fields listbox
    'and fills the Fields array for convenience
    '------------------------------------------
    
    Dim lngOpIdx As Long        'loop index
    Dim lngFieldIdx As Long     'loop index
    
    '''''''''''''''''''''''''''''''''''''
    'initialise and fill the fields array
    '''''''''''''''''''''''''''''''''''''
    
    'redim array to hold all fields
    ReDim ma_udtFields(53)
    
    'GAME OR SID NAME
    With ma_udtFields(e_CVFilterFieldPos_Name)
        .strFieldName = g_strField_Name
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Name"
            .strFieldTable = "Games"
            .strMusicFieldName = "Name"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Name"
            .strFieldTable = "Games"
            .strMusicFieldName = "Name"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Name"
            .strFieldTable = "Games"
            .strMusicFieldName = "Name"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Name"
            .strFieldTable = "Games"
            .strMusicFieldName = "Name"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With
    
    'PUBLISHER
    With ma_udtFields(e_CVFilterFieldPos_Publisher)
        .strFieldName = g_strField_Publisher
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Publisher"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Publisher"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Publisher"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Publisher"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "PU_Id"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "PU_Id"
            .strFieldTable = "Publishers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'DEVELOPER
    With ma_udtFields(e_CVFilterFieldPos_Developer)
        .strFieldName = g_strField_Developer
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Developer"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Developer"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Developer"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Developer"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "DE_Id"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "DE_Id"
            .strFieldTable = "Developers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'YEAR PUBLISHED
    With ma_udtFields(e_CVFilterFieldPos_Year)
        .strFieldName = g_strField_Year
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "YE_Id"
            .strFieldTable = "Years"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "YE_Id"
            .strFieldTable = "Years"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsAfter
            .strDisplayText = g_strCV_IsAfter
            .strFieldName = "Year"
            .strFieldTable = "Years"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsBefore
            .strDisplayText = g_strCV_IsBefore
            .strFieldName = "Year"
            .strFieldTable = "Years"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With
    
    'PROGRAMMER
    With ma_udtFields(e_CVFilterFieldPos_Programmer)
        .strFieldName = g_strField_Programmer
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Programmer"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Programmer"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Programmer"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Programmer"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "PR_Id"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "PR_Id"
            .strFieldTable = "Programmers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'MUSICIAN
    With ma_udtFields(e_CVFilterFieldPos_Musician)
        .strFieldName = g_strField_Musician
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Musician"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Musician"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Musician"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Musician"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Musician"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Musician"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Musician"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Musician"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "MU_Id"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "MU_Id"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "MU_Id"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "MU_Id"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'ARTIST
    With ma_udtFields(e_CVFilterFieldPos_Artist)
        .strFieldName = g_strField_Artist
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Artist"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Artist"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Artist"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Artist"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "AR_Id"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "AR_Id"
            .strFieldTable = "Artists"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'LICENSE
    With ma_udtFields(e_CVFilterFieldPos_License)
        .strFieldName = g_strField_License
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "License"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "License"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "License"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "License"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "LI_Id"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "LI_Id"
            .strFieldTable = "Licenses"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'REVIEW RATING
    With ma_udtFields(e_CVFilterFieldPos_ReviewRating)
        .strFieldName = g_strField_ReviewRating
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "ReviewRating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "ReviewRating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "ReviewRating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "ReviewRating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With

    End With

    'GENRE
    With ma_udtFields(e_CVFilterFieldPos_Genre)
        .strFieldName = g_strField_Genre
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Genre"
            .strFieldTable = "Genres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Genre"
            .strFieldTable = "Genres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "GE_Id"
            .strFieldTable = "Genres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "GE_Id"
            .strFieldTable = "Genres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'PARENT GENRE
    With ma_udtFields(e_CVFilterFieldPos_ParentGenre)
        .strFieldName = g_strField_ParentGenre
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "ParentGenre"
            .strFieldTable = "PGenres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "ParentGenre"
            .strFieldTable = "PGenres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "PG_Id"
            .strFieldTable = "PGenres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "PG_Id"
            .strFieldTable = "PGenres"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With
    
    'RARITY
    With ma_udtFields(e_CVFilterFieldPos_Rarity)
        .strFieldName = g_strField_Rarity
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Rarity"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Rarity"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Rarity"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Rarity"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "RA_Id"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "RA_Id"
            .strFieldTable = "Rarities"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With
    
    'NUMBER OF PLAYERS
    With ma_udtFields(e_CVFilterFieldPos_NumPlayers)
        .strFieldName = g_strField_NumPlayers
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "PlayersFrom"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "PlayersFrom"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "PlayersTo"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "PlayersFrom"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'SIMULTANEOUS PLAY
    With ma_udtFields(e_CVFilterFieldPos_SimulPlay)
        .strFieldName = g_strField_Players_Simultaneous
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "PlayersSim"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'LANGUAGE
    With ma_udtFields(e_CVFilterFieldPos_Language)
        .strFieldName = g_strField_Language
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Language"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Language"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Language"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Language"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "LA_Id"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "LA_Id"
            .strFieldTable = "Languages"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'WEBLINK NAME
    With ma_udtFields(e_CVFilterFieldPos_WebLinkName)
        .strFieldName = g_strField_WebLinkName
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'WEBLINK URL
    With ma_udtFields(e_CVFilterFieldPos_WebLinkURL)
        .strFieldName = g_strField_WebLinkURL
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'GAME HAS PREQUEL
    With ma_udtFields(e_CVFilterFieldPos_Prequel)
        .strFieldName = g_strField_Prequel
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Prequel"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'GAME HAS SEQUEL
    With ma_udtFields(e_CVFilterFieldPos_Sequel)
        .strFieldName = g_strField_Sequel
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Sequel"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'GAME HAS RELATED GAME
    With ma_udtFields(e_CVFilterFieldPos_Related)
        .strFieldName = g_strField_Related
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Related"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'GAME IS CLONE
    With ma_udtFields(e_CVFilterFieldPos_CloneOf)
        .strFieldName = g_strField_CloneOf
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "CloneOf"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'PERSONAL COMMENT
    With ma_udtFields(e_CVFilterFieldPos_Comment)
        .strFieldName = g_strField_Comment
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'GAME FILENAME
    With ma_udtFields(e_CVFilterFieldPos_Filename)
        .strFieldName = g_strField_GameFilename
        .lngMaxOps = 7
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FileExists
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Exists
            .strFieldName = "FA"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FileExists
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Missing
            .strFieldName = "FA"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(6)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(7)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Filename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With
    
    'SID FILENAME
    With ma_udtFields(e_CVFilterFieldPos_SidFilename)
        .strFieldName = g_strField_SidFilename
        .lngMaxOps = 7
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FileExists
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Exists
            .strFieldName = "SA"
            .strFieldTable = "Games"
            .strMusicFieldName = "SA"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FileExists
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Missing
            .strFieldName = "SA"
            .strFieldTable = "Games"
            .strMusicFieldName = "SA"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(6)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(7)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "SidFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = "Filename"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With
    
    'SCREENSHOT FILENAME
    With ma_udtFields(e_CVFilterFieldPos_ScrnShotFilename)
        .strFieldName = g_strField_ScreenShotFilename
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "ScrnShotFilename"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'VERSION CRACKER
    With ma_udtFields(e_CVFilterFieldPos_Cracker)
        .strFieldName = g_strField_Cracker
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Cracker"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Cracker"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Cracker"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Cracker"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "CR_Id"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "CR_Id"
            .strFieldTable = "Crackers"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'VERSION NUMBER OF TRAINERS
    With ma_udtFields(e_CVFilterFieldPos_NumTrainers)
        .strFieldName = g_strField_Trainers
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_Trainers"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_Trainers"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "V_Trainers"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "V_Trainers"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With

    End With



    'PLAYABLE
    With ma_udtFields(e_CVFilterFieldPos_Playable)
        .strFieldName = g_strField_Playable
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_Playable"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_Playable"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With
    
    'ORIGINAL
    With ma_udtFields(e_CVFilterFieldPos_Original)
        .strFieldName = g_strField_Original
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_Original"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_Original"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With
    
    'VERSION TITLE SCREEN
    With ma_udtFields(e_CVFilterFieldPos_TitleScreen)
        .strFieldName = g_strField_TitleScreen
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_TitleScreen"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_TitleScreen"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With


    'VERSION LOADING SCREEN
    With ma_udtFields(e_CVFilterFieldPos_LoadingScreen)
        .strFieldName = g_strField_LoadingScreen
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_LoadingScreen"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_LoadingScreen"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With

    'VERSION HIGH SCORE SAVER
    With ma_udtFields(e_CVFilterFieldPos_HighScoreSaver)
        .strFieldName = g_strField_HighScoreSaver
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_HighScoreSaver"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_HighScoreSaver"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With

    'VERSION INCLUDED DOCS
    With ma_udtFields(e_CVFilterFieldPos_IncludedDocs)
        .strFieldName = g_strField_IncludedDocs
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_IncludedDocs"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_IncludedDocs"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With

    'VERSION PAL/NTSC
    With ma_udtFields(e_CVFilterFieldPos_PalNTSC)
        .strFieldName = g_strField_PALNTSC
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_PalNTSC"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_PalNTSC"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'VERSION TRUE DRIVE EMULATION
    With ma_udtFields(e_CVFilterFieldPos_TrueDriveEmulation)
        .strFieldName = g_strField_TrueDriveEmulation
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_TrueDriveEmu"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_TrueDriveEmu"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNoUnknown
        End With
    End With

    'VERSION GAME LENGTH
    With ma_udtFields(e_CVFilterFieldPos_GameLength)
        .strFieldName = g_strField_GameLength
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_Length"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_Length"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "V_Length"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "V_Length"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'VERSION GAME LENGTH TYPE
    With ma_udtFields(e_CVFilterFieldPos_GameLengthType)
        .strFieldName = g_strField_GameLengthType
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "V_LengthType"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "V_LengthType"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'VERSION WEBLINK NAME
    With ma_udtFields(e_CVFilterFieldPos_VWebLinkName)
        .strFieldName = g_strField_VWebLinkName
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "V_WebLink_Name"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'VERSION WEBLINK URL
    With ma_udtFields(e_CVFilterFieldPos_VWebLinkURL)
        .strFieldName = g_strField_VWebLinkURL
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "V_WebLink_URL"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'VERSION COMMENT
    With ma_udtFields(e_CVFilterFieldPos_VersionComment)
        .strFieldName = g_strField_VersionComment
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "V_Comment"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'GAMEFILE-VERSION NUMBER
    With ma_udtFields(e_CVFilterFieldPos_GameVer)
        .strFieldName = g_strField_GameVer
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Version"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "Version"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "Version"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "Version"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With

    End With


    'DIFFICULTY
    With ma_udtFields(e_CVFilterFieldPos_Difficulty)
        .strFieldName = g_strField_Difficulty
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "DI_Id"
            .strFieldTable = "Difficulty"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "DI_Id"
            .strFieldTable = "Difficulty"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_DBTable
        End With
    End With

    'HIGH SCORE
    With ma_udtFields(e_CVFilterFieldPos_HighScore)
        .strFieldName = g_strField_HighScore
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "HighScore"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'FAVOURITE
    With ma_udtFields(e_CVFilterFieldPos_Favourite)
        .strFieldName = g_strField_Favourite
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Fav"
            .strFieldTable = "Games"
            .strMusicFieldName = "SFav"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'HAS EXTRAS
    With ma_udtFields(e_CVFilterFieldPos_Extras)
        .strFieldName = g_strField_Extras
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Extras"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'EXTRAS NAME
    With ma_udtFields(e_CVFilterFieldPos_Extras2)
        .strFieldName = g_strField_ExtrasName
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Name"
            .strFieldTable = "Extras"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Name"
            .strFieldTable = "Extras"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Name"
            .strFieldTable = "Extras"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Name"
            .strFieldTable = "Extras"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'CLASSIC GAME
    With ma_udtFields(e_CVFilterFieldPos_Classic)
        .strFieldName = g_strField_Classic
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Classic"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    'GAME RATING
    With ma_udtFields(e_CVFilterFieldPos_Rating)
        .strFieldName = g_strField_Rating
        .lngMaxOps = 3
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Rating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "Rating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsMoreThan
            .strDisplayText = g_strCV_MoreThan
            .strFieldName = "Rating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_IsLessThan
            .strDisplayText = g_strCV_LessThan
            .strFieldName = "Rating"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'GAME CONTROL
    With ma_udtFields(e_CVFilterFieldPos_Control)
        .strFieldName = g_strField_Control
        .lngMaxOps = 1
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Control"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotEqual
            .strFieldName = "Control"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Other
        End With
    End With

    'MEMO TEXT
    With ma_udtFields(e_CVFilterFieldPos_MemoText)
        .strFieldName = g_strField_Notes
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "MemoText"
            .strFieldTable = "Games"
            .strMusicFieldName = ""
            .strMusicFieldTable = ""
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'MUSICIAN GROUP
    With ma_udtFields(e_CVFilterFieldPos_MusicianGroup)
        .strFieldName = g_strField_MusicianGroup
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Grp"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Grp"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'MUSICIAN NICK NAME
    With ma_udtFields(e_CVFilterFieldPos_MusicianNick)
        .strFieldName = g_strField_MusicianNick
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Nick"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Nick"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'MUSICIAN PHOTO
    With ma_udtFields(e_CVFilterFieldPos_MusicianPhoto)
        .strFieldName = g_strField_MusicianPhoto
        .lngMaxOps = 5
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Filled
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(1)
            .eClauseType = e_ClauseType_FieldIsFilled
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_Empty
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_None
        End With
        With .udtAllowedOps(2)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Contains
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(3)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_DoesNotEqual
            .strDisplayText = g_strCV_NotContains
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(4)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_StartsWith
            .strDisplayText = g_strCV_StartsWith
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
        With .udtAllowedOps(5)
            .eClauseType = e_ClauseType_FieldContainsText
            .eOperator = e_Op_EndsWith
            .strDisplayText = g_strCV_EndsWith
            .strFieldName = "Photo"
            .strFieldTable = "Musicians"
            .strMusicFieldName = "Photo"
            .strMusicFieldTable = "Musicians"
            .eFilterSelector = e_FilterSelector_Text
        End With
    End With

    'ADULT FLAG
    With ma_udtFields(e_CVFilterFieldPos_Adult)
        .strFieldName = g_strField_Adult
        .lngMaxOps = 0
        ReDim .udtAllowedOps(.lngMaxOps)
        With .udtAllowedOps(0)
            .eClauseType = e_ClauseType_IsDatabaseField
            .eOperator = e_Op_Equal
            .strDisplayText = g_strCV_Equal
            .strFieldName = "Adult"
            .strFieldTable = "Games"
            .strMusicFieldName = "Adult"
            .strMusicFieldTable = "Music"
            .eFilterSelector = e_FilterSelector_YesNo
        End With
    End With

    ''''''''''''''''''''''''
    'fill the fields listbox
    ''''''''''''''''''''''''
    
    'copy text from array into listbox
    With lstField
        For lngFieldIdx = 0 To UBound(ma_udtFields)
            .AddItem ma_udtFields(lngFieldIdx).strFieldName
        Next lngFieldIdx
    End With

End Sub

Private Sub SetFilterDetails()

    '------------------------------------------------------
    'this will set the initial filter details if the user
    'is editing a filter.  It sets the controls so that
    'the filter is visible and selected and ready to change
    '------------------------------------------------------

    Dim lngIdxToSelect As CVFilterFieldPositionEnum     'the field index to select

    'find the correct field in the lstFields listbox to select
    lngIdxToSelect = gc_lngNOLISTINDEX
    With g_udtCVFilter
        Select Case .strFieldName
            Case "Name"
                If .strFieldTable = "Extras" Then
                    lngIdxToSelect = e_CVFilterFieldPos_Extras2
                Else
                    lngIdxToSelect = e_CVFilterFieldPos_Name
                End If
            Case "Publisher", "PU_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Publisher
            Case "Publisher", "AR_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Artist
            Case "Publisher", "DE_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Developer
            Case "Publisher", "LI_Id"
                lngIdxToSelect = e_CVFilterFieldPos_License
            Case "Publisher", "RA_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Rarity
            Case "Year", "YE_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Year
            Case "Programmer", "PR_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Programmer
            Case "Musician", "MU_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Musician
            Case "Genre", "GE_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Genre
            Case "ParentGenre", "PG_Id"
                lngIdxToSelect = e_CVFilterFieldPos_ParentGenre
            Case "PlayersFrom", "PlayersTo"
                lngIdxToSelect = e_CVFilterFieldPos_NumPlayers
            Case "PlayersSim"
                lngIdxToSelect = e_CVFilterFieldPos_SimulPlay
            Case "Language", "LA_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Language
            Case "Prequel"
                lngIdxToSelect = e_CVFilterFieldPos_Prequel
            Case "Sequel"
                lngIdxToSelect = e_CVFilterFieldPos_Sequel
            Case "Related"
                lngIdxToSelect = e_CVFilterFieldPos_CloneOf
            Case "CloneOf"
                lngIdxToSelect = e_CVFilterFieldPos_Related
            Case "Comment"
                lngIdxToSelect = e_CVFilterFieldPos_Comment
            Case "WebLink_Name"
                lngIdxToSelect = e_CVFilterFieldPos_WebLinkName
            Case "WebLink_URL"
                lngIdxToSelect = e_CVFilterFieldPos_WebLinkURL
            Case "V_WebLink_Name"
                lngIdxToSelect = e_CVFilterFieldPos_VWebLinkName
            Case "V_WebLink_URL"
                lngIdxToSelect = e_CVFilterFieldPos_VWebLinkURL
            Case "Filename", "FA"
                lngIdxToSelect = e_CVFilterFieldPos_Filename
            Case "SidFilename", "SA"
                lngIdxToSelect = e_CVFilterFieldPos_SidFilename
            Case "ScrnShotFilename"
                lngIdxToSelect = e_CVFilterFieldPos_ScrnShotFilename
            Case "Cracker", "CR_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Cracker
            Case "V_Trainers"
                lngIdxToSelect = e_CVFilterFieldPos_NumTrainers
            Case "V_LoadingScreen"
                lngIdxToSelect = e_CVFilterFieldPos_LoadingScreen
            Case "V_HighScoreSaver"
                lngIdxToSelect = e_CVFilterFieldPos_HighScoreSaver
            Case "V_Playable"
                lngIdxToSelect = e_CVFilterFieldPos_Playable
            Case "V_Original"
                lngIdxToSelect = e_CVFilterFieldPos_Original
            Case "V_TitleScreen"
                lngIdxToSelect = e_CVFilterFieldPos_TitleScreen
            Case "V_IncludedDocs"
                lngIdxToSelect = e_CVFilterFieldPos_IncludedDocs
            Case "V_PalNTSC"
                lngIdxToSelect = e_CVFilterFieldPos_PalNTSC
            Case "V_TrueDriveEmu"
                lngIdxToSelect = e_CVFilterFieldPos_TrueDriveEmulation
            Case "V_Length"
                lngIdxToSelect = e_CVFilterFieldPos_GameLength
            Case "V_LengthType"
                lngIdxToSelect = e_CVFilterFieldPos_GameLengthType
            Case "V_Comment"
                lngIdxToSelect = e_CVFilterFieldPos_VersionComment
            Case "Version"
                lngIdxToSelect = e_CVFilterFieldPos_GameVer
            Case "Difficulty", "DI_Id"
                lngIdxToSelect = e_CVFilterFieldPos_Difficulty
            Case "HighScore"
                lngIdxToSelect = e_CVFilterFieldPos_HighScore
            Case "Fav"
                lngIdxToSelect = e_CVFilterFieldPos_Favourite
            Case "Extras"
                lngIdxToSelect = e_CVFilterFieldPos_Extras
            Case "Classic"
                lngIdxToSelect = e_CVFilterFieldPos_Classic
            Case "Rating"
                lngIdxToSelect = e_CVFilterFieldPos_Rating
            Case "Control"
                lngIdxToSelect = e_CVFilterFieldPos_Control
            Case "MemoText"
                lngIdxToSelect = e_CVFilterFieldPos_MemoText
            Case "Grp"
                lngIdxToSelect = e_CVFilterFieldPos_MusicianGroup
            Case "Nick"
                lngIdxToSelect = e_CVFilterFieldPos_MusicianNick
            Case "Photo"
                lngIdxToSelect = e_CVFilterFieldPos_MusicianPhoto
            Case "Adult"
                lngIdxToSelect = e_CVFilterFieldPos_Adult
            Case "ReviewRating"
                lngIdxToSelect = e_CVFilterFieldPos_ReviewRating
        End Select
    End With

    'select the field - this triggers a chain reaction to fill the lstOperator listbox
    'and select the operator which fills the lstFilterData listbox or txtFilterData textbox
    'and selects that!
    lstField.ListIndex = lngIdxToSelect

End Sub

Private Sub cmdCancel_Click()

    'signal not to save
    g_udtCVFilter.strFieldName = ""

    'hide window
    Hide

End Sub

Private Sub cmdOK_Click()

    Dim lngCharIdx As Long      'loop index

    If txtFilterData.Visible = True Then
        'test for illegal characters in textbox
        For lngCharIdx = 1 To Len(mc_strILLEGALCHARS)
            If InStr(txtFilterData, Mid$(mc_strILLEGALCHARS, lngCharIdx, 1)) <> 0 Then
                MsgBox "The text you typed contains illegal characters!" & vbCrLf & vbCrLf & "You cannot use any of the following characters in the filter: -" & vbCrLf & vbCrLf & mc_strILLEGALCHARSMSG, vbExclamation, gc_strAPPNAME
                txtFilterData.SetFocus
                Exit Sub
            End If
        Next lngCharIdx
    End If

    'copy data across
    With ma_udtFields(lstField.ListIndex).udtAllowedOps(lstOperator.ListIndex)
        g_udtCVFilter.eClauseType = .eClauseType
        g_udtCVFilter.eOperator = .eOperator
        g_udtCVFilter.strFieldName = .strFieldName
        g_udtCVFilter.strFieldTable = .strFieldTable
        g_udtCVFilter.strMusicFieldName = .strMusicFieldName
        g_udtCVFilter.strMusicFieldTable = .strMusicFieldTable
        If lstFilterData.Visible = True Then
            g_udtCVFilter.strClauseData = lstFilterData.ItemData(lstFilterData.ListIndex)
        ElseIf txtFilterData.Visible = True Then
            g_udtCVFilter.strClauseData = txtFilterData
        Else
            g_udtCVFilter.strClauseData = ""
        End If
    End With

    'hide window
    Hide

End Sub


Private Sub Form_Activate()

    '---------------------------------------
    'Called once each time the form is shown
    '---------------------------------------

    'set the window title
    Caption = g_strCVFilterWindowTitle
    
    If g_blnEditingCVFilter Then
        'set the filter details to edit
        Call SetFilterDetails
        'only process once
        g_blnEditingCVFilter = False
    End If

End Sub

Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 4215    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 10935     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    'set the button text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    
    'set the labels
    lblField = g_strCVFilter_Field & ":"
    lblOperator = g_strCVFilter_Operator & ":"
    lblFilterData = g_strCVFilter_FilterData & ":"
    
    'load the fields into the fields listbox
    Call LoadFields
    
End Sub

Private Sub lstField_Click()

    '--------------------------------------------
    'load the field's operators into lstOperators
    '--------------------------------------------

    Dim lngOpIdx As Long                        'loop index
    Dim eOperator As FilterOperatorEnum         'for setting the operator when editing a filter
    Dim eClauseType As FilterClauseTypeEnum     'for setting the operator when editing a filter
    Dim blnFound As Boolean                     'found flag for operator search
    
    'only if a valid item
    If lstField.ListIndex = gc_lngNOLISTINDEX Then
        Exit Sub
    End If
    
    'hide and clear listbox
    lstOperator.Visible = False
    lstOperator.Clear

    'fill with field's operators
    With ma_udtFields(lstField.ListIndex)
        For lngOpIdx = 0 To .lngMaxOps
            lstOperator.AddItem .udtAllowedOps(lngOpIdx).strDisplayText
            'also store the Clause type and Operator type (needed when editing a filter)
            lstOperator.ItemData(lstOperator.NewIndex) = .udtAllowedOps(lngOpIdx).eClauseType & "," & .udtAllowedOps(lngOpIdx).eOperator
        Next lngOpIdx
    End With

    'show the listbox
    lstOperator.Visible = True
    
    
    'set the listindex for the operator
    If g_blnEditingCVFilter Then
        'find the selected operator for the filter we're editing
        With ma_udtFields(lstField.ListIndex)
            blnFound = False
            For lngOpIdx = 0 To .lngMaxOps
                If (g_udtCVFilter.eClauseType = .udtAllowedOps(lngOpIdx).eClauseType) And (g_udtCVFilter.eOperator = .udtAllowedOps(lngOpIdx).eOperator) Then
                    'set the listindex
                    blnFound = True
                    Exit For
                End If
            Next lngOpIdx
            'set the listindex if found
            If blnFound Then
                lstOperator.ListIndex = lngOpIdx
            Else
                'oops
                lstOperator.ListIndex = 0
            End If
        End With
    Else
        'just select the first operator
        lstOperator.ListIndex = 0
    End If
    
End Sub



Private Sub lstOperator_Click()

    '--------------------------------------------
    'load the field's operators into lstOperators
    '--------------------------------------------
    
    Dim lngDataIdx As Long      'loop index
    Dim blnFound As Boolean     'found flag for clause data (when editing a filter)

    'only if a valid item
    If lstOperator.ListIndex = gc_lngNOLISTINDEX Then
        Exit Sub
    End If

    'Show Filter Data
    Call ListFilterData
    
    If lstFilterData.Visible = True Then
        If g_blnEditingCVFilter Then
            'find editing filterdata to select
            blnFound = False
            With lstFilterData
                For lngDataIdx = 0 To (.ListCount - 1)
                    If .ItemData(lngDataIdx) = Val(g_udtCVFilter.strClauseData) Then
                        blnFound = True
                        Exit For
                    End If
                Next lngDataIdx
                If blnFound Then
                    'found it
                    lstFilterData.ListIndex = lngDataIdx
                Else
                    'oops
                    lstFilterData.ListIndex = 0
                End If
            End With
            
        Else
            'if listbox is visible, click first item
            lstFilterData.ListIndex = 0
            
        End If
        'set focus and enable ok button
        lstFilterData.SetFocus
        cmdOK.Enabled = True
        
    ElseIf txtFilterData.Visible = True Then
        If g_blnEditingCVFilter Then
            txtFilterData = g_udtCVFilter.strClauseData
            txtFilterData.SetFocus
        End If
        'if textbox is visible, only allow OK if not blank
        cmdOK.Enabled = (txtFilterData <> "")
    Else
        'no filterdata needed, enable ok button
        cmdOK.Enabled = True
        
    End If
End Sub

Private Sub txtFilterData_Change()

    cmdOK.Enabled = (txtFilterData <> "")

End Sub


Private Sub txtFilterData_GotFocus()

    With txtFilterData
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtFilterData_KeyPress(KeyAscii As Integer)

    'test for invalid character pressed
    If InStr(mc_strILLEGALCHARS, Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub


