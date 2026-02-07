VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1999-2010 James Burrows
' Email: gamebase.frontend@leesti.com
' Homepage: http://www.bu22.com/
'----------------------------------------------
'   GBSql:  Database Functions used in GB
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

Private Enum FilterClauseTypeEnum
    e_ClauseType_FieldContainsText
    e_ClauseType_FileExists
    e_ClauseType_FieldIsFilled
    e_ClauseType_IsDatabaseField
    e_ClauseType_NoClause
End Enum
Private Enum FilterOperatorEnum
    e_Op_Equal
    e_Op_DoesNotEqual
    e_Op_IsBefore
    e_Op_IsAfter
    e_Op_IsLessThan
    e_Op_IsMoreThan
    e_Op_StartsWith
    e_Op_EndsWith
End Enum
Private Type CVFilterFixedType
    strFieldTable As String * 128
    strFieldName As String * 128
    strMusicFieldTable As String * 128
    strMusicFieldName As String * 128
    eOperator As FilterOperatorEnum
    eClauseType As FilterClauseTypeEnum
    strClauseData As String * 64
End Type
Private Type CVFilterFixedBufferType
    strBuffer As String * 580
End Type
Private ma_udtFilters() As CVFilterFixedType
Private m_lngFilterCount As Long
Private m_blnGamesTable As Boolean
Private m_strSecColFields As String
Private m_strDefaultFields As String
Private m_strDefaultTable As String
Private m_blnANDMode As Boolean
Private m_blnNoAdultItems As Boolean

Private Const mc_strDEF_GAMES_FIELDS = "Games.Name, Games.CCode, Games.GA_Id, Games.Fav AS Bold1, Games.Classic AS Bold2"
Private Const mc_strDEF_GAMES_TABLE = "Games"

Private Const mc_strDEF_MUSIC_FIELDS = "Music.Name, Music.GA_Id, Music.SFav AS Bold1"
Private Const mc_strDEF_MUSIC_TABLE = "Music"
Private Function GetFilter(udtFilter_in As CVFilterFixedType) As String

    '--------------------------------------------------------------------------------
    'gets the filter text for the SQL query from the filter info in the passed in UDT
    '
    'Inputs:    udtFilter_in:   The structure containing the filter info
    '
    'Returns:   The filter text for the query
    '--------------------------------------------------------------------------------

    Dim strTableDotField As String  'TableName.FieldName concatenated string
    Dim strFieldName As String      'The Field Name
    Dim strTableName As String      'The Table Name
    Dim strClauseData As String     'The Clause data

    With udtFilter_in
        
        'Store for simplicity when building the SQL string
        If m_blnGamesTable Then
            strFieldName = Trim$(.strFieldName)
            strTableName = Trim$(.strFieldTable)
        Else
            strFieldName = Trim$(.strMusicFieldName)
            strTableName = Trim$(.strMusicFieldTable)
        End If
        strTableDotField = strTableName & "." & strFieldName
        strClauseData = Trim$(.strClauseData)
            
        'test which type of filter
        If .eClauseType = e_ClauseType_FieldContainsText Then
        
            'FIELD CONTAINS TEXT
            If .eOperator = e_Op_Equal Then
                If strTableName = "Genres" Then
                    GetFilter = strTableDotField & " LIKE ""*" & WCR(strClauseData) & "*"" OR PGenres.ParentGenre LIKE ""*" & WCR(strClauseData) & "*"""
                Else
                    GetFilter = strTableDotField & " LIKE ""*" & WCR(strClauseData) & "*"""
                End If
            ElseIf .eOperator = e_Op_DoesNotEqual Then
                If strTableName = "Genres" Then
                    GetFilter = strTableDotField & " NOT LIKE ""*" & WCR(strClauseData) & "*"" OR PGenres.ParentGenre LIKE ""*" & WCR(strClauseData) & "*"""
                Else
                    GetFilter = strTableDotField & " NOT LIKE ""*" & WCR(strClauseData) & "*"""
                End If
            ElseIf .eOperator = e_Op_StartsWith Then
                If strTableName = "Genres" Then
                    GetFilter = strTableDotField & " LIKE """ & WCR(strClauseData) & "*"" OR PGenres.ParentGenre LIKE """ & WCR(strClauseData) & "*"""
                Else
                    GetFilter = strTableDotField & " LIKE """ & WCR(strClauseData) & "*"""
                End If
            Else    'e_Op_EndsWith
                If strTableName = "Genres" Then
                    GetFilter = strTableDotField & " LIKE ""*" & WCR(strClauseData) & """ OR PGenres.ParentGenre LIKE ""*" & WCR(strClauseData) & """"
                Else
                    GetFilter = strTableDotField & " LIKE ""*" & WCR(strClauseData) & """"
                End If
            End If

        ElseIf .eClauseType = e_ClauseType_FieldIsFilled Then
        
            'FIELD IS FILLED
            If .eOperator = e_Op_Equal Then
                GetFilter = strTableDotField & " <> """""
            Else
                GetFilter = strTableDotField & " = """""
            End If
        
        ElseIf .eClauseType = e_ClauseType_FileExists Then
            
            'FILE EXISTS
            If .eOperator = e_Op_Equal Then
                GetFilter = strTableDotField & " = True"
            Else
                GetFilter = strTableDotField & " = False"
            End If

        ElseIf .eClauseType = e_ClauseType_IsDatabaseField Then
            
            'DATABASE FIELD
            If .eOperator = e_Op_Equal Then
                If strFieldName = "PlayersFrom" Then
                    'number of players
                    GetFilter = "(" & strTableDotField & " = " & strClauseData & ") AND (Games.PlayersTo = " & strClauseData & ")"
                ElseIf strFieldName = "PlayersTo" Then
                    'number of players
                    GetFilter = "(" & strTableDotField & " = " & strClauseData & ")"
                ElseIf strFieldName = "Prequel" Or strFieldName = "Sequel" Or strFieldName = "Related" Or strFieldName = "CloneOf" Then
                    'prequel, sequel, related game and clones
                    If strClauseData = "-1" Then
                        GetFilter = strTableDotField & " > 0"   'has
                    Else
                        GetFilter = strTableDotField & " = 0"   'hasn't
                    End If
                Else
                    'all else
                    GetFilter = strTableDotField & " = " & strClauseData
                End If
            ElseIf .eOperator = e_Op_DoesNotEqual Then
                If strFieldName = "PlayersFrom" Then
                    'number of players
                    GetFilter = "NOT (" & strTableDotField & " = " & strClauseData & ") AND (Games.PlayersTo = " & strClauseData & ")"
                Else
                    'all else
                    GetFilter = "NOT (" & strTableDotField & " = " & strClauseData & ")"
                End If
            ElseIf .eOperator = e_Op_IsAfter Then
                'years only
                GetFilter = "(" & strTableDotField & " > " & strClauseData & ") AND (" & strTableDotField & " < 9991)"
            ElseIf .eOperator = e_Op_IsBefore Then
                'years only
                GetFilter = strTableDotField & " < " & strClauseData
            ElseIf .eOperator = e_Op_IsLessThan Then
                If (strFieldName = "PlayersTo") Or (strFieldName = "V_Trainers") Or (strFieldName = "ReviewRating") Then
                    'number of players, trainers, review rating
                    GetFilter = "(" & strTableDotField & " < " & strClauseData & ") AND (" & strTableDotField & " > -1)"
                ElseIf strFieldName = "V_Length" Or strFieldName = "V_LengthType" Then
                    'game length
                    GetFilter = "(" & strTableDotField & " < " & strClauseData & ")"
                Else
                    'rating
                    GetFilter = "(" & strTableDotField & " < " & strClauseData & ") AND (" & strTableDotField & " > 0)"
                End If
            ElseIf .eOperator = e_Op_IsMoreThan Then
                If (strFieldName = "PlayersFrom") Then
                    'number of players
                    GetFilter = "(" & strTableDotField & " > " & strClauseData & ") OR (PlayersTo > " & strClauseData & ")"
                ElseIf (strFieldName = "V_Trainers") Or (strFieldName = "Rating") Or (strFieldName = "ReviewRating") Then
                    'trainers, rating, review rating
                    GetFilter = strTableDotField & " > " & strClauseData
                ElseIf strFieldName = "V_Length" Or strFieldName = "V_LengthType" Then
                    'game length
                    GetFilter = strTableDotField & " > " & strClauseData
                End If
            End If
        
        Else 'NoClause
            
            'to avoid error
            GetFilter = m_strDefaultTable & ".Name <> """""
            
        End If
        
    End With
    

End Function

Private Function GetLinkField(strTableName_in As String) As String

    '------------------------------------------------------------------
    'gets the link field to use in the INNER JOIN part of a query
    '
    'Inputs:    strTableName_in:    The table to get the link field for
    '
    'Returns:   a string containing the link field
    '------------------------------------------------------------------
    
    Select Case strTableName_in
    
        Case "Publishers"
            GetLinkField = "PU_Id"
        
        Case "Years"
            GetLinkField = "YE_Id"
        
        Case "Programmers"
            GetLinkField = "PR_Id"
            
        Case "Musicians"
            GetLinkField = "MU_Id"
            
        Case "Languages"
            GetLinkField = "LA_Id"
            
        Case "Genres"
            GetLinkField = "GE_Id"
            
        Case "PGenres"
            GetLinkField = "PG_Id"
            
        Case "Difficulty"
            GetLinkField = "DI_Id"
            
        Case "Crackers"
            GetLinkField = "CR_Id"

        Case "Artists"
            GetLinkField = "AR_Id"

        Case "Developers"
            GetLinkField = "DE_Id"

        Case "Licenses"
            GetLinkField = "LI_Id"

        Case "Rarities"
            GetLinkField = "RA_Id"

        Case "Extras"
            GetLinkField = "GA_Id"
            
    End Select

End Function

Public Function GetViewQueryString(ByVal blnGamesTable_in As Boolean, strSecColFields_in As String, ByVal lngFilterCount_in As Long, a_strFilterBuffers_in() As String, ByVal blnANDMode_in As Boolean, ByVal blnNoAdultItems_in As Boolean) As String

    '-----------------------------------------------------------------------------------
    'Gets the SQL Query string for a View in Gamebase64 and Filter64
    '
    'Inputs:    blnGamesTable_in:       Whether we are operating on Games or Music table
    '           strSecColFields_in:     The Fields shown in the second column
    '           lngFilterCount_in:      The number of filters in the view
    '           a_strFilterBuffers_in:  The string buffer array containing the filters
    '           blnANDMode_in:          Flag for Filter Mode (AND/OR)
    '           blnNoAdultItems_in:     Flag for filtering out the Adult Games/Sids
    '
    'Returns:   The SQL string for the view
    '
    'Remarks:   If either lngFilterCount_in = 0 or strFilterBuffer_in = "" then
    '           it is taken that there are no filters for the view
    '-----------------------------------------------------------------------------------
    
    Dim lngFilterIdx As Long                    'loop index
    Dim strBuffer As CVFilterFixedBufferType    'for converting the filter data
    
    'set the filters array up
    If lngFilterCount_in = 0 Then
        'no filters
        m_lngFilterCount = 0
        ReDim ma_udtFilters(0)
    Else
        'at least one filter
        m_lngFilterCount = lngFilterCount_in
        ReDim ma_udtFilters(m_lngFilterCount)
        'copy into filters array
        For lngFilterIdx = 1 To m_lngFilterCount
            strBuffer.strBuffer = a_strFilterBuffers_in(lngFilterIdx)
            LSet ma_udtFilters(lngFilterIdx) = strBuffer
        Next lngFilterIdx
    End If

    'store the rest of the data
    m_blnANDMode = blnANDMode_in
    m_strSecColFields = strSecColFields_in
    m_blnGamesTable = blnGamesTable_in
    m_blnNoAdultItems = blnNoAdultItems_in
    If m_blnGamesTable Then
        m_strDefaultFields = mc_strDEF_GAMES_FIELDS
        m_strDefaultTable = mc_strDEF_GAMES_TABLE
    Else
        m_strDefaultFields = mc_strDEF_MUSIC_FIELDS
        m_strDefaultTable = mc_strDEF_MUSIC_TABLE
    End If

    'set the SELECT Tables/Fields and INNER JOIN part of the query
    GetViewQueryString = GetTablesAndFields()
    
    'set the WHERE clause of the query
    If m_lngFilterCount > 0 Then
        GetViewQueryString = GetViewQueryString & " " & GetWhereClause()
    Else
        'just adult filter?
        If m_blnNoAdultItems Then
            GetViewQueryString = GetViewQueryString & " WHERE (" & m_strDefaultTable & ".Adult = False)"
        End If
    End If
    
    'set the ORDER BY Clause
    GetViewQueryString = GetViewQueryString & " ORDER BY " & m_strDefaultTable & ".Name"
    
    'make it distinct if the extras stuff is in there
    If InStr(GetViewQueryString, "Extras.GA_Id") > 0 Then
        GetViewQueryString = Replace(GetViewQueryString, "SELECT ", "SELECT DISTINCT ", 1, 1)
        GetViewQueryString = Replace(GetViewQueryString, "FROM Extras INNER ", "FROM Extras RIGHT ", 1, 1)
    End If
    
End Function


Private Function GetWhereClause() As String

    '-----------------------------------------------------------------------
    'traverses all the filters and formulates the WHERE clause for the query
    '
    'Returns:  The WHERE Clause
    '-----------------------------------------------------------------------

    Dim lngFilterIdx As Long    'loop index
    Dim strANDOR As String      'filter mode
    
    'set the filter mode string
    strANDOR = IIf((m_blnANDMode), " AND ", " OR ")

    'start the where clause with an opening bracket
    GetWhereClause = GetWhereClause & "WHERE ("
    
    'add each filter to the where clause, separate with the AND/OR operator
    For lngFilterIdx = 1 To m_lngFilterCount
        
        'opening filter bracket
        GetWhereClause = GetWhereClause & "("
        
        'get the filter
        With ma_udtFilters(lngFilterIdx)
            GetWhereClause = GetWhereClause & GetFilter(ma_udtFilters(lngFilterIdx))
        End With
        
        'closing filter bracket
        GetWhereClause = GetWhereClause & ")"
        
        'add the filter mode string if more filters to come
        If lngFilterIdx < m_lngFilterCount Then
            GetWhereClause = GetWhereClause & strANDOR
        End If
        
    Next lngFilterIdx
    
    'add the adult filter restriction if needed
    If m_blnNoAdultItems Then
        GetWhereClause = GetWhereClause & " AND (" & m_strDefaultTable & ".Adult = False)"
    End If
    
    'finish the WHERE Clause with a closing bracket
    GetWhereClause = GetWhereClause & ")"

End Function


Private Function GetTablesAndFields() As String

    '----------------------------------------------------------------------------
    'takes the 3 sets of fields for a query and constructs a formatted
    'SQL SELECT string for the SQL Fields and Tables arguments
    '
    'Returns:       String containing the formatted fields and tables in the
    '               format: "SELECT <Fields> FROM <TABLE> INNER JOIN (..."
    '
    'Remarks:       Aliases are preserved if there are 2 instances of the same
    '               field, (one with an Alias and one without).
    '
    'Assumptions:   Each field in each string must be in the format: -
    '               "TableName.FieldName" or "TableName.FieldName AS AliasName".
    '               Each string of fields must not start with or end with the
    '               delimiting comma character.
    '---------------------------------------------------------------------------
    
    Dim a_strFields() As String         'array for fields
    Dim lngMaxFields As Long            'maximum number of fields in a_strFields() array
    Dim lngCharIdx As Long              'loop index
    Dim strChar As String               'character variable
    Dim lngFieldIdx As Long             'loop index
    Dim lngCorrectFieldIdx As Long      'loop index
    Dim lngFindIdx As Long              'loop index
    Dim lngTableIdx As Long             'loop index
    Dim lngFilterIdx As Long            'loop index
    Dim a_strCorrectFields() As String  'array for fields after filtering out duplicates
    Dim lngMaxCorrectFields As Long     'maximum number of fields in a_strCorrectFields() array
    Dim blnFound As Boolean             'flag for removing duplicates
    Dim a_strTables() As String         'array for unique table names
    Dim lngMaxTables As Long            'maximum number of fields in a_strTables() array
    Dim strTable As String              'storage for Table name
    Dim strFieldWithoutAlias As String  'for duplicate field comparisons
    Dim lngAliasStart As Long           'for duplicate field comparisons
    Dim strLinkField As String          'the link field for the INNER JOINs
    Dim strLinkTable As String          'the link table for the INNER Joins
    Dim blnIncludePGenre As Boolean     'genres exception for table includes
    Dim blnIncludeGenre As Boolean      'genres exception for table includes
    
    '''''''''''''''''''''''''''''''
    'set the beginning of the query
    '''''''''''''''''''''''''''''''
    
    GetTablesAndFields = "SELECT "
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'parse all the fields into an array for easy processing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'initialise the array
    lngMaxFields = 0
    ReDim a_strFields(0)
    
    'default fields
    For lngCharIdx = 1 To Len(m_strDefaultFields)
        strChar = Mid$(m_strDefaultFields, lngCharIdx, 1)
        If strChar = "," Then
            'start a new field
            lngMaxFields = lngMaxFields + 1
            ReDim Preserve a_strFields(lngMaxFields)
        Else
            'continue storing current field
            a_strFields(lngMaxFields) = a_strFields(lngMaxFields) & strChar
        End If
    Next lngCharIdx

    'second column fields
    If m_strSecColFields <> "" Then
        'at least one Second Column field
        lngMaxFields = lngMaxFields + 1
        ReDim Preserve a_strFields(lngMaxFields)
        For lngCharIdx = 1 To Len(m_strSecColFields)
            strChar = Mid$(m_strSecColFields, lngCharIdx, 1)
            If strChar = "," Then
                'start a new field
                lngMaxFields = lngMaxFields + 1
                ReDim Preserve a_strFields(lngMaxFields)
            Else
                'continue storing current field
                a_strFields(lngMaxFields) = a_strFields(lngMaxFields) & strChar
            End If
        Next lngCharIdx
    End If

    ''''''''''''''''''''''''''''''''''''
    'Compare and remove duplicate fields
    ''''''''''''''''''''''''''''''''''''
    
    'initialise the a_strCorrectFields() array and copy the first field into it
    ReDim a_strCorrectFields(0)
    lngMaxCorrectFields = 0
    a_strCorrectFields(0) = Trim$(a_strFields(0))
    
    'copy each field to the a_strCorrectFields() array if not already in there
    For lngFieldIdx = 1 To lngMaxFields
        blnFound = False
        'strip the alias
        lngAliasStart = InStr(a_strFields(lngFieldIdx), " AS ")
        If lngAliasStart = 0 Then
            strFieldWithoutAlias = Trim$(a_strFields(lngFieldIdx))
        Else
            strFieldWithoutAlias = Trim$(Left$(a_strFields(lngFieldIdx), lngAliasStart - 1))
        End If
        'check for duplicate
        For lngFindIdx = 0 To lngMaxCorrectFields
            If InStr(a_strCorrectFields(lngFindIdx), strFieldWithoutAlias) > 0 Then
                'preserve the Alias for the second columns
                If (InStr(a_strCorrectFields(lngFindIdx), " AS ") = 0) And (InStr(a_strFields(lngFieldIdx), " AS ") > 0) Then
                    'replace with the Aliased field
                    a_strCorrectFields(lngFindIdx) = Trim$(a_strFields(lngFieldIdx))
                End If
                blnFound = True
                Exit For
            End If
        Next lngFindIdx
        'add if no duplicate was found
        If blnFound = False Then
            lngMaxCorrectFields = lngMaxCorrectFields + 1
            ReDim Preserve a_strCorrectFields(lngMaxCorrectFields)
            a_strCorrectFields(lngMaxCorrectFields) = Trim$(a_strFields(lngFieldIdx))
        End If
    Next lngFieldIdx
    
    '''''''''''''''''''''''''''''''''
    'set the Fields part of the query
    '''''''''''''''''''''''''''''''''

    For lngCorrectFieldIdx = 0 To lngMaxCorrectFields
        GetTablesAndFields = GetTablesAndFields & a_strCorrectFields(lngCorrectFieldIdx) & ", "
    Next lngCorrectFieldIdx
    'take off excess ", " characters
    GetTablesAndFields = Left$(GetTablesAndFields, Len(GetTablesAndFields) - 2)
    'add a space before the FROM part
    GetTablesAndFields = GetTablesAndFields & " "
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    'get the unique tables from the visible fields
    ''''''''''''''''''''''''''''''''''''''''''''''

    'initialise the a_strTables() array and copy the first table into it
    ReDim a_strTables(0)
    lngMaxTables = 0
    a_strTables(0) = Left$(a_strCorrectFields(0), InStr(a_strCorrectFields(0), ".") - 1)
    
    'copy each table to the a_strTables() array if not already in there
    For lngCorrectFieldIdx = 1 To lngMaxCorrectFields
        'check for duplicate
        blnFound = False
        strTable = Left$(a_strCorrectFields(lngCorrectFieldIdx), InStr(a_strCorrectFields(lngCorrectFieldIdx), ".") - 1)
        For lngFindIdx = 0 To lngMaxTables
            If strTable = a_strTables(lngFindIdx) Then
                blnFound = True
                Exit For
            End If
        Next lngFindIdx
        'add if no duplicate was found
        If blnFound = False Then
            lngMaxTables = lngMaxTables + 1
            ReDim Preserve a_strTables(lngMaxTables)
            a_strTables(lngMaxTables) = strTable
        End If
    Next lngCorrectFieldIdx

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'get the unique tables used by the filters (for the inner joins)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'copy each table to the a_strTables() array if not already in there
    If m_lngFilterCount > 0 Then
        blnIncludePGenre = False
        blnIncludeGenre = False
        For lngFilterIdx = 1 To m_lngFilterCount
            'check for duplicate
            If m_blnGamesTable Then
                strTable = Trim$(ma_udtFilters(lngFilterIdx).strFieldTable)
                If strTable = "Genres" Then
                    blnIncludePGenre = True
                End If
               If strTable = "PGenres" Then
                    blnIncludeGenre = True
                End If
            Else
                strTable = Trim$(ma_udtFilters(lngFilterIdx).strMusicFieldTable)
            End If
            blnFound = False
            For lngFindIdx = 0 To lngMaxTables
                If strTable = a_strTables(lngFindIdx) Then
                    blnFound = True
                    Exit For
                End If
            Next lngFindIdx
            'add if no duplicate was found
            If blnFound = False Then
                lngMaxTables = lngMaxTables + 1
                ReDim Preserve a_strTables(lngMaxTables)
                a_strTables(lngMaxTables) = strTable
            End If
        Next lngFilterIdx
        'add PGenres if Genres exists
        If blnIncludePGenre Then
            blnFound = False
            For lngFindIdx = 0 To lngMaxTables
                If "PGenres" = a_strTables(lngFindIdx) Then
                    blnFound = True
                    Exit For
                End If
            Next lngFindIdx
            If blnFound = False Then
                lngMaxTables = lngMaxTables + 1
                ReDim Preserve a_strTables(lngMaxTables)
                a_strTables(lngMaxTables) = "PGenres"
            End If
        End If
        'add Genres if PGenres exists
        If blnIncludeGenre Then
            blnFound = False
            For lngFindIdx = 0 To lngMaxTables
                If "Genres" = a_strTables(lngFindIdx) Then
                    blnFound = True
                    Exit For
                End If
            Next lngFindIdx
            If blnFound = False Then
                lngMaxTables = lngMaxTables + 1
                ReDim Preserve a_strTables(lngMaxTables)
                a_strTables(lngMaxTables) = "Genres"
            End If
        End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Re-order tables so that PGenres is first and Games/Music
    'is last so that the INNER JOINs can be formulated
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If lngMaxTables > 0 Then
        'Move PGenres to first
        For lngTableIdx = 0 To lngMaxTables
            If a_strTables(lngTableIdx) = "PGenres" Then
                If lngTableIdx > 0 Then
                    'swap to first entry
                    strTable = a_strTables(0)
                    a_strTables(0) = a_strTables(lngTableIdx)
                    a_strTables(lngTableIdx) = strTable
                End If
                'done
                Exit For
            End If
        Next lngTableIdx
        'Move Games/Music to last
        For lngTableIdx = 0 To lngMaxTables
            If a_strTables(lngTableIdx) = m_strDefaultTable Then
                If lngTableIdx < lngMaxTables Then
                    'swap to last entry
                    strTable = a_strTables(lngMaxTables)
                    a_strTables(lngMaxTables) = a_strTables(lngTableIdx)
                    a_strTables(lngTableIdx) = strTable
                End If
                'done
                Exit For
            End If
        Next lngTableIdx
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Build the FROM <TABLE> INNER JOIN <TABLES> part of the SQL string
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GetTablesAndFields = GetTablesAndFields & "FROM "
    If lngMaxTables > 0 Then
        'inbound recursive INNER JOINS
        For lngTableIdx = 0 To lngMaxTables
            GetTablesAndFields = GetTablesAndFields & a_strTables(lngTableIdx)
            If lngTableIdx < lngMaxTables Then
                GetTablesAndFields = GetTablesAndFields & " INNER JOIN "
            End If
            If lngTableIdx < (lngMaxTables - 1) Then
                GetTablesAndFields = GetTablesAndFields & "("
            End If
        Next lngTableIdx
        'outbound recursive ON Table.Key=Games/Music.Key (except PGenres which uses Genres.PG_Id)
        For lngTableIdx = (lngMaxTables - 1) To 0 Step -1
            strLinkField = GetLinkField(a_strTables(lngTableIdx))
            If a_strTables(lngTableIdx) = "PGenres" Then
                strLinkTable = "Genres"
            Else
                strLinkTable = m_strDefaultTable
            End If
            GetTablesAndFields = GetTablesAndFields & " ON " & a_strTables(lngTableIdx) & "." & strLinkField & " = " & strLinkTable & "." & strLinkField
            If lngTableIdx > 0 Then
                GetTablesAndFields = GetTablesAndFields & ")"
            End If
        Next lngTableIdx
    Else
        'no inner joins, only one table
        GetTablesAndFields = GetTablesAndFields & a_strTables(0)
    End If

End Function

Private Function WCR(strSearchString_in As String) As String

    'WildCard Replacer

    WCR = Replace(strSearchString_in, "[", "[[]")
    WCR = Replace(WCR, "?", "[?]")
    WCR = Replace(WCR, "*", "[*]")
    WCR = Replace(WCR, "!", "[!]")
    WCR = Replace(WCR, "#", "[#]")

End Function


