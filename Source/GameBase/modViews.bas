Attribute VB_Name = "modViews"
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

'==============================================
'Contains all global variables and functions
'for the views engine, including the SQL engine
'and Second Column stuff.
'==============================================


'Enums for the Views engine
Public Enum ExtraColumnsTypeEnum       'extra columns
    e_SecColType_Normal
    e_SecColType_YesNo
    e_SecColType_Publisher
    e_SecColType_Year
    e_SecColType_Trainers
    e_SecColType_Players
    e_SecColType_Genres
    e_SecColType_Rating
    e_SecColType_PALNTSC
    e_SecColType_Control
    e_SecColType_Length
    e_SecColType_ReviewRating
End Enum
Public Enum FilterModeEnum          'AND or OR View mode
    e_FilterMode_AND
    e_FilterMode_OR
End Enum
Public Enum FilterClauseTypeEnum    'Each filter's clause type in a view
    e_ClauseType_FieldContainsText
    e_ClauseType_FileExists
    e_ClauseType_FieldIsFilled
    e_ClauseType_IsDatabaseField
    e_ClauseType_NoClause
End Enum
Public Enum FilterOperatorEnum      'Each filter's operator type in a view
    e_Op_Equal
    e_Op_DoesNotEqual
    e_Op_IsBefore
    e_Op_IsAfter
    e_Op_IsLessThan
    e_Op_IsMoreThan
    e_Op_StartsWith
    e_Op_EndsWith
End Enum
Public Enum TablesToIncludeEnum     'The tables to include in the view
    e_Table_GamesOnly
    e_Table_SidOnly
    e_Table_Both
End Enum

'Types for the Views engine
Public Type ExtraColumnFieldsType
    strFields As String
    strMusicFields As String
End Type
Public Type ExtraColumnsType                    'extra columns
    strFields As String
    strMusicFields As String
    strMenuText As String
    eType As ExtraColumnsTypeEnum
    strColumnTitle As String
    blnGameOnlyTable As Boolean
End Type
Public Type ExtraColumnsArrayType               'extra columns
    a_udtColumnItems() As ExtraColumnsType
    lngMaxColumns As Long
    strMenuText As String
End Type
Public Type CVFilterType                        'A Custom View's "Filter"
    strFieldTable As String
    strFieldName As String
    strMusicFieldTable As String
    strMusicFieldName As String
    eOperator As FilterOperatorEnum
    eClauseType As FilterClauseTypeEnum
    strClauseData As String
End Type
Public Type ExtraColumnType                     'each column in a view
    lngTypeIndex As Long
    lngFieldIndex As Long
    lngColumnHeaderWidth As Long
    strColumnName As String
End Type
Public Type ViewDataType                        'The Custom View's Main Type
    lngID As Long
    eFilterMode As FilterModeEnum
    strTitle As String
    lngFilterCount As Long
    a_udtFilters() As CVFilterType
    eIncludeMusicTable As TablesToIncludeEnum
    lngListViewType As Long
    lngSortColumn As Long
    lngSortOrder As Long
    lngCol1Width As Long
    strSelectedItem As String
    lngMaxExtraColumns As Long
    lngExtraColumnsCount As Long
    a_udtExtraColumns() As ExtraColumnType
    strExtraColumns As String
End Type
Public Type QuickFilterType                     'The Quick Filter View's Main Type
    strFieldTable(2) As String
    strFieldName(2) As String
    strMusicFieldTable(2) As String
    strMusicFieldName(2) As String
    strClauseData(2) As String
    eOperator(2) As FilterOperatorEnum
    eClauseType(2) As FilterClauseTypeEnum
    eTablesToInclude As TablesToIncludeEnum
    strTitle As String
    eFilterMode As FilterModeEnum
    lngNumFilters As Long
End Type

'The Views Engine variables
Public Const gc_lngDEFAULTCOLUMNWIDTH As Long = 3845            'Default Column width for the game list in twips
Public Const gc_lngDEFAULTEXTRACOLUMNWIDTH As Long = 2500       'Default Column width for the game list in twips
Public Const gc_lngACCEPTABLEWAIT As Long = 2000                'Views taking longer than this no. of milliseconds to load into the listview will show the "Please Wait..." message
Public Const gc_lngFAVOURITESINDEX As Long = 2                  'Index of the Favourites view
Public Const gc_lngDEFAULTVIEWSMAX As Long = 3                  'Total of #default views - 1
Public ga_udtViewData() As ViewDataType                         'Array for ALL the views
Public g_udtCustomView As ViewDataType                          'Storage for a new or edited custom view
Public g_udtCVFilter As CVFilterType                            'Storage for a new or edited custom view filter
Public g_udtQuickFilter As QuickFilterType                      'The Quick Filter storage
Public g_blnSaveCustomView As Boolean                           'Flag to say the user clicked OK to save a custom view
Public g_blnEditingCVFilter As Boolean                          'If True the custom view filter window will display the current filter settings
Public g_blnSettingCatFilter As Boolean                         'Flag for when we are setting a view from the Category window
Public ga_strViewModesLookupTable(1) As String                  'Lookup table for the View modes "AND" "OR"
Public g_lngRegViewIndex As Long                                'Index of the currently selected view
Public ga_blnViewsLoaded(gc_lngDEFAULTVIEWSMAX + 1) As Boolean  'To ascertain if each view has been loaded into the listbox yet
Public ga_strDefaultViewNames(gc_lngDEFAULTVIEWSMAX) As String  'Array for the default views
Public g_blnRegHideAdultItems As Boolean                        'Adult Filter flag
Public g_strAdultPassword As String                             'Password for the Adult Filter
Public ga_udtExtraColumns() As ExtraColumnsArrayType
Public g_lngMaxExtraColumnArrays As Long
Public g_udtExtraColumnFields As ExtraColumnFieldsType
Public g_blnViewsWereSorted As Boolean

Public Function GetFormattedDate(strDate_in As String) As String

    'formats the date from the database/registry ("dd-mm-yyyy") into "dd mmm yyyy"
    'country independant

    Dim lngMonth As Long
    Dim lngDay As Long
    Dim lngYear As Long

    If Len(strDate_in) = 10 Then
        lngDay = Val(Left$(strDate_in, 2))
        lngMonth = Val(Mid$(strDate_in, 4, 2))
        lngYear = Val(Right$(strDate_in, 4))
        If (lngDay < 1) Or (lngDay > 31) Or (lngMonth < 1) Or (lngMonth > 12) Or (lngYear < 2001) Or (lngYear > 9999) Then
            GetFormattedDate = ""
        Else
            GetFormattedDate = Format$(lngDay, "00") & " " & MonthName(lngMonth, True) & " " & Format$(lngYear, "0000")
        End If
    Else
        GetFormattedDate = ""
    End If

End Function
