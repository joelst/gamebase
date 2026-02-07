VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColumns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Columns"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmColumns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvPossible 
      Height          =   8775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   15478
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdFieldRemoveAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdFieldRemove 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdFieldAddAll 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdFieldAdd 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   8055
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   14208
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblPrompt 
      Caption         =   "&Selected Columns:"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblPrompt 
      Caption         =   "&Possible Columns:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmColumns"
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

Private Type ColumnSelectorColumnType
    strKey As String
    strName As String
    lngTypeIdx As Long
    lngColumnIdx As Long
    lngIdx As Long
End Type
Private ma_udtColumns() As ColumnSelectorColumnType
Private m_lngTotalColumns As Long
Private ma_udtPossible() As ColumnSelectorColumnType
Private m_lngTotalPossible As Long
Private ma_udtSelected() As ColumnSelectorColumnType
Private m_lngTotalSelected As Long

Private Function FindColumnIdx(strKey As String)

    Dim lngIdx As Long
    
    For lngIdx = 0 To m_lngTotalColumns - 1
        If strKey = ma_udtColumns(lngIdx).strKey Then
            FindColumnIdx = lngIdx
            Exit Function
        End If
    Next lngIdx
    
    FindColumnIdx = m_lngTotalColumns
    Exit Function

End Function

Private Function GetColumnArrayIndexForKey(strKey_in As String)

    Dim lngColumnIdx As Long

    For lngColumnIdx = 0 To m_lngTotalColumns - 1
        If ma_udtColumns(lngColumnIdx).strKey = strKey_in Then
            GetColumnArrayIndexForKey = lngColumnIdx
            Exit Function
        End If
    Next lngColumnIdx
    
    'should never happen
    GetColumnArrayIndexForKey = 0

End Function

Private Sub MoveToPossible()

    Dim LItem As ListItem
    Dim PLitem As ListItem
    Dim a_lngToRemove() As Long
    Dim lngRemoveCount As Long
    Dim lngIdx As Long
    Dim lngSelectedIdx As Long
    Dim lngPossibleIdx As Long
    Dim blnFound As Boolean
    Dim lngFirstSelectedIdx As Long
    Dim lngLastIdx As Long

    lvPossible.Enabled = False
    lvSelected.Enabled = False
    lvPossible.Visible = False
    lvSelected.Visible = False
    
    lngFirstSelectedIdx = -1
    lngRemoveCount = 0
    For Each LItem In lvSelected.ListItems
        With LItem
            If .Selected Then
                If lngFirstSelectedIdx = -1 Then lngFirstSelectedIdx = LItem.Index
                lngSelectedIdx = GetColumnArrayIndexForKey(.Key)
                blnFound = False
                For Each PLitem In lvPossible.ListItems
                    lngPossibleIdx = GetColumnArrayIndexForKey(PLitem.Key)
                    If lngPossibleIdx > lngSelectedIdx Then
                        lvPossible.ListItems.Add PLitem.Index, .Key, .Text
                        blnFound = True
                        Exit For
                    End If
                Next PLitem
                If Not blnFound Then
                    lvPossible.ListItems.Add , .Key, .Text
                End If
                ReDim Preserve a_lngToRemove(lngRemoveCount)
                a_lngToRemove(lngRemoveCount) = .Index
                lngRemoveCount = lngRemoveCount + 1
            End If
        End With
    Next LItem

    For lngIdx = lngRemoveCount - 1 To 0 Step -1
        lvSelected.ListItems.Remove a_lngToRemove(lngIdx)
    Next lngIdx

    lvPossible.Enabled = True
    lvSelected.Enabled = True
    lvPossible.Visible = True
    lvSelected.Visible = True

    With lvSelected
        If .ListItems.Count > 0 And lngFirstSelectedIdx > -1 Then
            lngLastIdx = .ListItems.Count - 1
            If lngLastIdx < lngFirstSelectedIdx Then
                lngFirstSelectedIdx = lngLastIdx
            Else
                lngFirstSelectedIdx = lngFirstSelectedIdx - 1
            End If
            .SelectedItem = .ListItems(lngFirstSelectedIdx + 1)
            .SelectedItem.Selected = True
            .SetFocus
        End If
        If .ListItems.Count = 0 Then
            lvPossible.SetFocus
        End If
    End With
    
    Call SetButtonsEnabledState

End Sub
Private Sub MoveToSelected()

    Dim LItem As ListItem
    Dim a_lngToRemove() As Long
    Dim lngRemoveCount As Long
    Dim lngIdx As Long
    Dim lngFirstSelectedIdx As Long
    Dim lngLastIdx As Long
    
    lvPossible.Enabled = False
    lvSelected.Enabled = False
    lvPossible.Visible = False
    lvSelected.Visible = False

    lngFirstSelectedIdx = -1
    lngRemoveCount = 0
    For Each LItem In lvPossible.ListItems
        With LItem
            If .Selected Then
                If lngFirstSelectedIdx = -1 Then lngFirstSelectedIdx = LItem.Index
                lvSelected.ListItems.Add , .Key, .Text
                ReDim Preserve a_lngToRemove(lngRemoveCount)
                a_lngToRemove(lngRemoveCount) = .Index
                lngRemoveCount = lngRemoveCount + 1
            End If
        End With
    Next LItem
    
    For lngIdx = lngRemoveCount - 1 To 0 Step -1
        lvPossible.ListItems.Remove a_lngToRemove(lngIdx)
    Next lngIdx

    lvPossible.Enabled = True
    lvSelected.Enabled = True
    lvPossible.Visible = True
    lvSelected.Visible = True

    With lvPossible
        If .ListItems.Count > 0 And lngFirstSelectedIdx > -1 Then
            lngLastIdx = .ListItems.Count - 1
            If lngLastIdx < lngFirstSelectedIdx Then
                lngFirstSelectedIdx = lngLastIdx
            Else
                lngFirstSelectedIdx = lngFirstSelectedIdx - 1
            End If
            .SelectedItem = .ListItems(lngFirstSelectedIdx + 1)
            .SelectedItem.Selected = True
            .SelectedItem.EnsureVisible
            .SetFocus
        End If
        If .ListItems.Count = 0 Then
            lvSelected.SetFocus
        End If
    End With
    
    Call SetButtonsEnabledState

End Sub

Private Sub SetButtonsEnabledState()

    cmdFieldRemove.Enabled = (lvSelected.ListItems.Count > 0)
    cmdFieldRemoveAll.Enabled = cmdFieldRemove.Enabled
    cmdFieldAdd.Enabled = (lvPossible.ListItems.Count > 0)
    cmdFieldAddAll.Enabled = cmdFieldAdd.Enabled

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdFieldAdd_Click()

    Call MoveToSelected

End Sub

Private Sub cmdFieldAddAll_Click()

    Dim LItem As ListItem

    For Each LItem In lvPossible.ListItems
        LItem.Selected = True
    Next LItem
    
    Call MoveToSelected

End Sub


Private Sub cmdFieldRemove_Click()

    Call MoveToPossible

End Sub

Private Sub cmdFieldRemoveAll_Click()

    Dim LItem As ListItem

    For Each LItem In lvSelected.ListItems
        LItem.Selected = True
    Next LItem
    Call MoveToPossible

End Sub


Private Sub cmdOK_Click()

    Dim LItem As ListItem
    Dim strExtraColumns As String
    Dim lngColumnIdx As Long

    strExtraColumns = ""
    For Each LItem In lvSelected.ListItems
    
        strExtraColumns = strExtraColumns & LItem.Key & ","
    
    Next LItem
    If Right$(strExtraColumns, 1) = "," Then
        strExtraColumns = Left$(strExtraColumns, Len(strExtraColumns) - 1)
    End If
    
    With ga_udtViewData(g_lngRegViewIndex)
        .strExtraColumns = strExtraColumns
        .lngExtraColumnsCount = ParseExtraColumns(.a_udtExtraColumns, strExtraColumns)
        .lngMaxExtraColumns = UBound(.a_udtExtraColumns)
        
        'now the number of columns is known, get the extra columns widths
        For lngColumnIdx = 0 To .lngExtraColumnsCount - 1
            .a_udtExtraColumns(lngColumnIdx).lngColumnHeaderWidth = gc_lngDEFAULTCOLUMNWIDTH
        Next lngColumnIdx

    End With
    
    g_blnOKClicked = True
    Unload Me

End Sub

Private Sub Form_Load()

    Dim lngTypeIdx As Long
    Dim lngColumnIdx As Long
    Dim lngIdx As Long
    Dim lngSelectedIdx As Long
    Dim blnFound As Boolean

    Const DesignScaleHeight As Single = 9015    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6015     'The required scalewidth for the form
    
    g_blnOKClicked = False
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText

    With g_xobjFunc
        'with shortcuts
        lblPrompt(0).Caption = .GetINIMessage("Columns Window", "PossibleLabel", "&Possible Columns") & ":"
        lblPrompt(1).Caption = .GetINIMessage("Columns Window", "SelectedLabel", "&Selected Columns") & ":"
        'without shortcuts
        Caption = .GetINIMessage("Columns Window", "Caption", "Select Extra Columns")
    End With

    'all columns
    lngIdx = 0
    If ga_udtViewData(g_lngRegViewIndex).eIncludeMusicTable <> e_Table_SidOnly Then
        'games table (or games and music)
        m_lngTotalColumns = 0
        For lngTypeIdx = 0 To g_lngMaxExtraColumnArrays
            For lngColumnIdx = 0 To ga_udtExtraColumns(lngTypeIdx).lngMaxColumns
                If ga_udtExtraColumns(lngTypeIdx).a_udtColumnItems(lngColumnIdx).strColumnTitle <> "" Then
                    ReDim Preserve ma_udtColumns(m_lngTotalColumns)
                    With ma_udtColumns(m_lngTotalColumns)
                        .strKey = lngTypeIdx & "-" & lngColumnIdx
                        .strName = ga_udtExtraColumns(lngTypeIdx).a_udtColumnItems(lngColumnIdx).strColumnTitle
                        .lngTypeIdx = lngTypeIdx
                        .lngColumnIdx = lngColumnIdx
                        .lngIdx = .lngIdx + 1
                    End With
                    lngIdx = lngIdx + 1
                    m_lngTotalColumns = m_lngTotalColumns + 1
                End If
            Next lngColumnIdx
        Next lngTypeIdx
    Else
        'music table only
        m_lngTotalColumns = 2
        ReDim Preserve ma_udtColumns(m_lngTotalColumns - 1)
        With ma_udtColumns(0)
            .strKey = "0-4" 'musician
            .strName = ga_udtExtraColumns(0).a_udtColumnItems(4).strColumnTitle
            .lngTypeIdx = 0
            .lngColumnIdx = 4
            .lngIdx = 0
        End With
        With ma_udtColumns(1)
            .strKey = "1-16" 'music filename
            .strName = ga_udtExtraColumns(1).a_udtColumnItems(16).strColumnTitle
            .lngTypeIdx = 1
            .lngColumnIdx = 16
            .lngIdx = 1
        End With
    End If

    'selected columns
    m_lngTotalSelected = 0
    With ga_udtViewData(g_lngRegViewIndex)
        For lngIdx = 0 To .lngExtraColumnsCount - 1
            ReDim Preserve ma_udtSelected(m_lngTotalSelected)
            With .a_udtExtraColumns(m_lngTotalSelected)
                ma_udtSelected(m_lngTotalSelected).strKey = .lngTypeIndex & "-" & .lngFieldIndex
                ma_udtSelected(m_lngTotalSelected).strName = .strColumnName
                ma_udtSelected(m_lngTotalSelected).lngTypeIdx = .lngTypeIndex
                ma_udtSelected(m_lngTotalSelected).lngColumnIdx = .lngFieldIndex
                ma_udtSelected(m_lngTotalSelected).lngIdx = FindColumnIdx(ma_udtSelected(m_lngTotalSelected).strKey)
            End With
            m_lngTotalSelected = m_lngTotalSelected + 1
        Next lngIdx
    End With

    'possible columns
    m_lngTotalPossible = 0
    For lngIdx = 0 To m_lngTotalColumns - 1
        blnFound = False
        For lngSelectedIdx = 0 To m_lngTotalSelected - 1
            If ma_udtColumns(lngIdx).strKey = ma_udtSelected(lngSelectedIdx).strKey Then
                blnFound = True
                Exit For
            End If
        Next lngSelectedIdx
        If Not blnFound Then
            ReDim Preserve ma_udtPossible(m_lngTotalPossible)
            ma_udtPossible(m_lngTotalPossible) = ma_udtColumns(lngIdx)
            m_lngTotalPossible = m_lngTotalPossible + 1
        End If
    Next lngIdx

    'fill listboxes
    lvPossible.ColumnHeaders.Add , , , lvPossible.Width - 100
    lvSelected.ColumnHeaders.Add , , , lvSelected.Width - 100
    Call FillListboxes
    
    Call SetButtonsEnabledState

End Sub

Private Sub FillListboxes()

    Dim lngIdx As Long

    lvPossible.Enabled = False
    lvSelected.Enabled = False
    lvPossible.Visible = False
    lvSelected.Visible = False

    For lngIdx = 0 To m_lngTotalPossible - 1
        lvPossible.ListItems.Add , ma_udtPossible(lngIdx).strKey, ma_udtPossible(lngIdx).strName
    Next lngIdx
    For lngIdx = 0 To m_lngTotalSelected - 1
        lvSelected.ListItems.Add , ma_udtSelected(lngIdx).strKey, ma_udtSelected(lngIdx).strName
    Next lngIdx

    lvPossible.Enabled = True
    lvSelected.Enabled = True
    lvPossible.Visible = True
    lvSelected.Visible = True

End Sub

Private Sub lvPossible_DblClick()

    Call MoveToSelected

End Sub


Private Sub lvSelected_DblClick()

    Call MoveToPossible

End Sub

