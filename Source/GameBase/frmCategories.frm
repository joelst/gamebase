VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategories 
   Caption         =   "Categories"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3015
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvRecords 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox lstCategory 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmCategories"
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

Private Sub Form_Load()

    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips
    Const DesignScaleHeight As Single = 1935    'The required minimum scaleheight for the form
    Const DesignScaleWidth As Single = 3015     'The required minimum scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngCatWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngCatWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngCatWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngCatWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatWidth", sngRegWidth, False, gc_lngINVALIDWINDOWSIZE
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatHeight", sngRegHeight, False, gc_lngINVALIDWINDOWSIZE
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatTop", sngRegTop, False, gc_lngINVALIDWINDOWPOS
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatLeft", sngRegLeft, False, gc_lngINVALIDWINDOWPOS
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
    Top = sngRegTop
    Left = sngRegLeft
    Width = sngRegWidth
    Height = sngRegHeight

    'subclass the window
    If g_blnInVBIDE = False Then
        Call CatSubClass(Me.hwnd)
    End If

    'read in the ini file text
    With g_xobjFunc
        Caption = GetINIMessage("Category Window", "Caption", "Categories")
        lstCategory.AddItem GetINIMessage("Category Window", "Publishers", "Publishers")
        lstCategory.AddItem GetINIMessage("Category Window", "Years", "Years")
        lstCategory.AddItem GetINIMessage("Category Window", "Developers", "Developers")
        lstCategory.AddItem GetINIMessage("Category Window", "Programmers", "Programmers")
        lstCategory.AddItem GetINIMessage("Category Window", "Musicians", "Musicians")
        lstCategory.AddItem GetINIMessage("Category Window", "Artists", "Artists")
        lstCategory.AddItem GetINIMessage("Category Window", "Licenses", "Licenses")
        lstCategory.AddItem GetINIMessage("Category Window", "Genres", "Genres")
        lstCategory.AddItem GetINIMessage("Category Window", "Languages", "Languages")
        lstCategory.AddItem GetINIMessage("Category Window", "Rarities", "Rarities")
        lstCategory.AddItem GetINIMessage("Category Window", "Crackers", "Crackers")
    End With

    g_blnLoadedfrmCategories = True

End Sub


Private Sub Form_Resize()

    If WindowState = vbNormal Then
        lstCategory.Width = ScaleWidth
        lvRecords.Width = ScaleWidth
        lvRecords.Height = ScaleHeight - lvRecords.Top
        lvRecords.ColumnHeaders(1).Width = lvRecords.Width - 300
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'clear the records listbox
    Call SendMessageAny(lvRecords.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnCatSubclassed = True Then
            Call CatUnSubClass(Me.hwnd)
        End If
    End If

    'save window sizes etc...
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "CatLeft", Left, False
        End If
    End With
    
    g_blnLoadedfrmCategories = False

End Sub


Private Sub lstCategory_Click()

    Dim strSQL(10) As String
    Dim fldID As Field
    Dim fldName As Field
    Dim fldFilter As Field
    Dim fldFilter2 As Field

    If lstCategory.ListIndex > gc_lngNOLISTINDEX Then
    
        Call SendMessageAny(lvRecords.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
        Screen.MousePointer = vbHourglass
        DoEvents
        lvRecords.Enabled = False
        lvRecords.Visible = False
        strSQL(0) = "SELECT PU_Id AS ID, Publisher AS Name FROM Publishers ORDER BY Publisher"
        strSQL(1) = "SELECT YE_Id AS ID, Year AS Name FROM Years ORDER BY YE_Id"
        strSQL(2) = "SELECT DE_Id AS ID, Developer AS Name FROM Developers ORDER BY Developer"
        strSQL(3) = "SELECT PR_Id AS ID, Programmer AS Name FROM Programmers ORDER BY Programmer"
        strSQL(4) = "SELECT MU_Id AS ID, Musician AS Name FROM Musicians ORDER BY Musician"
        strSQL(5) = "SELECT AR_Id AS ID, Artist AS Name FROM Artists ORDER BY Artist"
        strSQL(6) = "SELECT LI_Id AS ID, License AS Name FROM Licenses ORDER BY License"
        strSQL(7) = "SELECT Genre, ParentGenre, GE_ID FROM PGenres INNER JOIN Genres ON PGenres.PG_Id = Genres.PG_Id WHERE PGenres.ParentGenre <> """" ORDER BY PGenres.ParentGenre, Genres.Genre"
        strSQL(8) = "SELECT LA_Id AS ID, Language AS Name FROM Languages ORDER BY Language"
        strSQL(9) = "SELECT RA_Id AS ID, Rarity AS Name FROM Rarities ORDER BY Rarity"
        strSQL(10) = "SELECT CR_Id AS ID, Cracker AS Name FROM Crackers ORDER BY Cracker"

        If lstCategory.ListIndex <> 1 And lstCategory.ListIndex <> 7 Then
        
            'NORMAL TABLE

            Set g_rsGames = g_dbGamebase.OpenRecordset(strSQL(lstCategory.ListIndex), dbOpenSnapshot)
            If Not g_rsGames.EOF Then
                Set fldID = g_rsGames!ID
                Set fldName = g_rsGames!Name
                g_rsGames.MoveFirst
                With lvRecords.ListItems
                    Do Until g_rsGames.EOF
                        Call .Add(, "I" & fldID, fldName)
                        g_rsGames.MoveNext
                    Loop
                End With
                Set fldID = Nothing
                Set fldName = Nothing
            End If
            g_rsGames.Close
            Set g_rsGames = Nothing

        ElseIf lstCategory.ListIndex = 1 Then

            'YEARS TABLE

            Set g_rsGames = g_dbGamebase.OpenRecordset(strSQL(1), dbOpenSnapshot)
            If Not g_rsGames.EOF Then
                Set fldID = g_rsGames!ID
                Set fldName = g_rsGames!Name
                g_rsGames.MoveFirst
                With lvRecords.ListItems
                    Do Until g_rsGames.EOF
                        Call .Add(, "I" & fldID, ga_strYearLookupTable(fldID))
                        g_rsGames.MoveNext
                    Loop
                End With
                Set fldID = Nothing
                Set fldName = Nothing
            End If
            g_rsGames.Close
            Set g_rsGames = Nothing
            

        Else
        
            'GENRES TABLE

            Set g_rsGames = g_dbGamebase.OpenRecordset(strSQL(7), dbOpenSnapshot)
            Set fldFilter = g_rsGames!Genre
            Set fldFilter2 = g_rsGames!ParentGenre
            Set fldID = g_rsGames!GE_Id
            With lvRecords.ListItems
                g_rsGames.MoveFirst
                Do Until g_rsGames.EOF
                    Call .Add(, "I" & fldID, fldFilter2 & " - " & fldFilter)
                    g_rsGames.MoveNext
                Loop
                g_rsGames.Close
            End With
            Set fldFilter = Nothing
            Set fldFilter2 = Nothing
            Set fldID = Nothing
            Set g_rsGames = Nothing

        End If
        
        lvRecords.Visible = True
        lvRecords.Enabled = True
        Screen.MousePointer = vbDefault
        lvRecords.SetFocus
        Set lvRecords.SelectedItem = Nothing
        If lvRecords.ListItems.Count > 0 Then
            lvRecords.ListItems(1).Selected = False
        End If

    End If

End Sub




Private Sub lvRecords_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim strID As Long

    If lstCategory.ListIndex > gc_lngNOLISTINDEX Then

        g_blnSettingCatFilter = True
        With Item
            strID = Right$(.Key, Len(.Key) - 1)
        End With
        Screen.MousePointer = vbHourglass

        If lstCategory.ListIndex = 0 Then

            'PUBLISHERS

            With g_udtQuickFilter
                .strTitle = g_strField_Publisher & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Publishers"
                .strFieldName(0) = "PU_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)
        
        ElseIf lstCategory.ListIndex = 1 Then

            'YEARS

            With g_udtQuickFilter
                .strTitle = g_strField_Year & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Years"
                .strFieldName(0) = "YE_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 2 Then

            'DEVELOPERS

            With g_udtQuickFilter
                .strTitle = g_strField_Developer & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Developers"
                .strFieldName(0) = "DE_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 3 Then

            'PROGRAMMERS

            With g_udtQuickFilter
                .strTitle = g_strField_Programmer & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Programmers"
                .strFieldName(0) = "PR_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 4 Then

            'MUSICIANS

            With g_udtQuickFilter
                .strTitle = g_strField_Musician & " = " & Item.Text
                .strMusicFieldTable(0) = "Musicians"
                .strMusicFieldName(0) = "MU_Id"
                .strFieldTable(0) = "Musicians"
                .strFieldName(0) = "MU_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_Both
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 5 Then

            'ARTISTS

            With g_udtQuickFilter
                .strTitle = g_strField_Artist & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Artists"
                .strFieldName(0) = "AR_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 6 Then

            'LICENCES

            With g_udtQuickFilter
                .strTitle = g_strField_License & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Licenses"
                .strFieldName(0) = "LI_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 7 Then

            'GENRES

            With g_udtQuickFilter
                .strTitle = g_strField_Genre & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Genres"
                .strFieldName(0) = "GE_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 8 Then

            'LANGUAGES

            With g_udtQuickFilter
                .strTitle = g_strField_Language & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Languages"
                .strFieldName(0) = "LA_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 9 Then

            'LICENCES

            With g_udtQuickFilter
                .strTitle = g_strField_Rarity & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Rarities"
                .strFieldName(0) = "RA_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        ElseIf lstCategory.ListIndex = 10 Then

            'CRACKERS

            With g_udtQuickFilter
                .strTitle = g_strField_Cracker & " = " & Item.Text
                .strMusicFieldTable(0) = ""
                .strMusicFieldName(0) = ""
                .strFieldTable(0) = "Crackers"
                .strFieldName(0) = "CR_Id"
                .strClauseData(0) = strID
                .eOperator(0) = e_Op_Equal
                .eClauseType(0) = e_ClauseType_IsDatabaseField
                .eTablesToInclude = e_Table_GamesOnly
                .eFilterMode = e_FilterMode_AND
                .lngNumFilters = 1
            End With
            Call frmMain.SetQuickFilter(False)

        End If

        g_blnSettingCatFilter = False
        Screen.MousePointer = vbDefault

    End If

End Sub


