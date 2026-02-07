VERSION 5.00
Begin VB.Form frmQuickFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Filter"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmQuickFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lstField 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame fraTables 
      Caption         =   "Include Tables"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
      Begin VB.OptionButton optTable 
         Caption         =   "&Both Tables"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Sids Table Only"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Games Table Only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   120
      MaxLength       =   64
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblContains 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contains text:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmQuickFilter"
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

Private Type QuickFieldType
    strListBoxName As String
    strFieldTable As String
    strMusicFieldName As String
    strMusicFieldTable As String
    strFieldName As String
    blnAllowSidTable As Boolean
End Type
Private ma_udtFields() As QuickFieldType

Private Const mc_lngMAXQFIELDS As Long = 11
Private Const mc_strILLEGALCHARS As String = """*[]"
Private Const mc_strILLEGALCHARSMSG As String = """   [   ]   *"

Private m_blnFormIsLoading As Boolean
Private m_blnClickedByCode As Boolean

Private Sub LoadQuickFilterFields()

    '---------------------------------------------------------
    'loads the quick filter fields into the quick filter array
    'so that it's ready to fill the field listbox
    '---------------------------------------------------------
    
    Dim lngFieldIdx As Long     'loop index
    
    'redim the array
    ReDim ma_udtFields(mc_lngMAXQFIELDS)
    
    'set the individual items
    With ma_udtFields(0)
        .strFieldName = "Name"
        .strFieldTable = "Games"
        .strMusicFieldName = "Name"
        .strMusicFieldTable = "Music"
        .strListBoxName = ga_QuickFilterNames(0)
        .blnAllowSidTable = True
    End With
    With ma_udtFields(1)
        .strFieldName = "Publisher"
        .strFieldTable = "Publishers"
        .strListBoxName = ga_QuickFilterNames(1)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(2)
        .strFieldName = "Developer"
        .strFieldTable = "Developers"
        .strListBoxName = ga_QuickFilterNames(2)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(3)
        .strFieldName = "Programmer"
        .strFieldTable = "Programmers"
        .strListBoxName = ga_QuickFilterNames(3)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(4)
        .strFieldName = "Musician"
        .strFieldTable = "Musicians"
        .strListBoxName = ga_QuickFilterNames(4)
        .strMusicFieldName = "Musician"
        .strMusicFieldTable = "Musicians"
        .blnAllowSidTable = True
    End With
    With ma_udtFields(5)
        .strFieldName = "Artist"
        .strFieldTable = "Artists"
        .strListBoxName = ga_QuickFilterNames(5)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(6)
        .strFieldName = "License"
        .strFieldTable = "Licenses"
        .strListBoxName = ga_QuickFilterNames(6)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(7)
        .strFieldName = "Genre"
        .strFieldTable = "Genres"
        .strListBoxName = ga_QuickFilterNames(7)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(8)
        .strFieldName = "Language"
        .strFieldTable = "Languages"
        .strListBoxName = ga_QuickFilterNames(8)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(9)
        .strFieldName = "Rarity"
        .strFieldTable = "Rarities"
        .strListBoxName = ga_QuickFilterNames(9)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(10)
        .strFieldName = "Comment"
        .strFieldTable = "Games"
        .strListBoxName = ga_QuickFilterNames(10)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With
    With ma_udtFields(11)
        .strFieldName = "V_Comment"
        .strFieldTable = "Games"
        .strListBoxName = ga_QuickFilterNames(11)
        .strMusicFieldName = ""
        .strMusicFieldTable = ""
        .blnAllowSidTable = False
    End With

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim lngCharIdx As Long          'loop index
    Dim lngChosenTableIdx As Long   'the chosen table index
    
    'test for illegal characters
    For lngCharIdx = 1 To Len(mc_strILLEGALCHARS)
        If InStr(txtFilter, Mid$(mc_strILLEGALCHARS, lngCharIdx, 1)) <> 0 Then
            MsgBox "The text you typed contains illegal characters!" & vbCrLf & vbCrLf & "You cannot use any of the following characters in the filter: -" & vbCrLf & vbCrLf & mc_strILLEGALCHARSMSG, vbExclamation, gc_strAPPNAME
            txtFilter.SetFocus
            Exit Sub
        End If
    Next lngCharIdx

    'easter egg?
    If LCase$(txtFilter) = "i want an easter egg" Then
        g_blnEasterEgg = True
        GoTo Egg_Exit
    End If

    'save the selected properties to the registry for next time
    If optTable(0).Value = True Then
        lngChosenTableIdx = 0
    ElseIf optTable(1).Value = True Then
        lngChosenTableIdx = 1
    Else
        lngChosenTableIdx = 2
    End If
    With g_xobjRegistry
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedField", lstField.ListIndex, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedText", txtFilter.Text, True
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedTable", lngChosenTableIdx, False
    End With

    'save info
    With g_udtQuickFilter
        .eTablesToInclude = lngChosenTableIdx
        .eFilterMode = e_FilterMode_AND
        .strClauseData(0) = txtFilter
        .strFieldName(0) = ma_udtFields(lstField.ListIndex).strFieldName
        .strFieldTable(0) = ma_udtFields(lstField.ListIndex).strFieldTable
        .strTitle = ma_udtFields(lstField.ListIndex).strListBoxName & " = *" & txtFilter & "*"
        .strMusicFieldName(0) = ma_udtFields(lstField.ListIndex).strMusicFieldName
        .strMusicFieldTable(0) = ma_udtFields(lstField.ListIndex).strMusicFieldTable
        .eOperator(0) = e_Op_Equal
        .eClauseType(0) = e_ClauseType_FieldContainsText
        .lngNumFilters = 1
    End With

Egg_Exit:
    Unload Me

End Sub


Private Sub Form_Activate()
    
    'start with focus on the textbox
    txtFilter.SetFocus

End Sub

Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 3255    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 2655     'The required scalewidth for the form
    Dim lngFieldIdx As Long                     'loop index
    Dim lngSelectedFieldIdx As Long             'selected field chosen last time
    Dim strSelectedText As String               'selected text chosen last time
    Dim lngSelectedTable As Long                'selected table option last time
    Dim lngCharIdx As Long                      'loop index
    
    'we are loading!
    m_blnFormIsLoading = True
    
    'no easter egg yet
    g_blnEasterEgg = False

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the default text
    With g_xobjFunc
        Caption = .GetINIMessage("Quick Filter", "Caption", "Quick Filter")
        lblField.Caption = .GetINIMessage("Quick Filter", "Field", "&Field") & ":"
        lblContains.Caption = .GetINIMessage("Quick Filter", "ContainsText", "&Contains text") & ":"
        fraTables.Caption = .GetINIMessage("Quick Filter", "Tab_Include_Tables", "Include Tables")
        optTable(0).Caption = .GetINIMessage("Quick Filter", "GamesTableOnly", "&Games Table Only")
        optTable(1).Caption = .GetINIMessage("Quick Filter", "MusicTableOnly", "&Music Table Only")
        optTable(2).Caption = .GetINIMessage("Quick Filter", "BothTables", "&Both Tables")
    End With
    
    'load the field info
    Call LoadQuickFilterFields
    
    'fill the fields listbox
    With lstField
        For lngFieldIdx = 0 To mc_lngMAXQFIELDS
            .AddItem ma_udtFields(lngFieldIdx).strListBoxName
        Next lngFieldIdx
    End With
    
    'retrieve the default properties (from last time)
    With g_xobjRegistry
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedField", lngSelectedFieldIdx, False, 0
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedText", strSelectedText, True, ""
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\QuickFilter", "SelectedTable", lngSelectedTable, False, 0
    End With
    'check field validity...
    'field index
    If lngSelectedFieldIdx < 0 Or lngSelectedFieldIdx > mc_lngMAXQFIELDS Then
        lngSelectedFieldIdx = 0
    End If
    'selected text
    For lngCharIdx = 1 To Len(mc_strILLEGALCHARS)
        If InStr(strSelectedText, Mid$(mc_strILLEGALCHARS, lngCharIdx, 1)) <> 0 Then
            strSelectedText = ""
            Exit For
        End If
    Next lngCharIdx
    If Len(strSelectedText) > 64 Then
        strSelectedText = Left$(strSelectedText, 64)
    End If
    'table
    If ma_udtFields(lngSelectedFieldIdx).blnAllowSidTable = False Then
        'only games table allowed
        lngSelectedTable = 0
    Else
        'both tables allowed
        If lngSelectedTable < 0 Or lngSelectedTable > 2 Then
            lngSelectedTable = 0
        End If
    End If
    
    'set the properties from last time
    lstField.ListIndex = lngSelectedFieldIdx
    txtFilter.Text = strSelectedText
    optTable(0).Value = (lngSelectedTable = 0)
    optTable(1).Value = (lngSelectedTable = 1)
    optTable(2).Value = (lngSelectedTable = 2)
    optTable(1).Enabled = (ma_udtFields(lngSelectedFieldIdx).blnAllowSidTable = True)
    optTable(2).Enabled = optTable(1).Enabled
    
    'user clicks first
    m_blnClickedByCode = False
    
    'finished loading!
    m_blnFormIsLoading = False
    
    'clear out the structure
    With g_udtQuickFilter
        For lngFieldIdx = 0 To 2
            .strClauseData(lngFieldIdx) = ""
            .strFieldName(lngFieldIdx) = ""
            .strFieldTable(lngFieldIdx) = ""
            .strMusicFieldName(lngFieldIdx) = ""
            .strMusicFieldTable(lngFieldIdx) = ""
            .eClauseType(lngFieldIdx) = e_ClauseType_FieldContainsText
            .eOperator(lngFieldIdx) = e_Op_Equal
        Next lngFieldIdx
        .strTitle = ""
        .eFilterMode = e_FilterMode_AND
        .lngNumFilters = 0
    End With
    
    'now goes to Form_Activate()
    
End Sub


Private Sub lstField_Click()

    If m_blnFormIsLoading = False Then
        'set the table buttons
        If ma_udtFields(lstField.ListIndex).blnAllowSidTable = False Then
            'sid table not allowed
            optTable(1).Enabled = False
            optTable(2).Enabled = False
            optTable(0).Value = True
        Else
            optTable(1).Enabled = True
            optTable(2).Enabled = True
        End If
    End If

End Sub


Private Sub txtFilter_Change()

    'enable/disable the ok button
    cmdOK.Enabled = (txtFilter.Text <> "")

End Sub

Private Sub txtFilter_GotFocus()

    With txtFilter
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtFilter_KeyPress(KeyAscii As Integer)

    'test for invalid character pressed
    If InStr(mc_strILLEGALCHARS, Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    

End Sub


