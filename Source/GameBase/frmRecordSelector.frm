VERSION 5.00
Begin VB.Form frmRecordSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Selector"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmRecordSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRecs 
      Caption         =   "Select Record"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtFilter 
         Height          =   375
         Left            =   6240
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstRecs 
         Enabled         =   0   'False
         Height          =   2985
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label txtFilterLabel 
         Caption         =   "Filter by:"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmRecordSelector"
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

Private ma_lngNoEdit() As Long
Private m_lngMaxNoEdit As Long
Private m_blnRecordChanged As Boolean
Private m_strFieldKey As String
Private m_strTableName As String
Private m_strFieldName As String
Private m_strFrameCaption As String

Private Const mc_strILLEGALCHARS As String = """*[]"
Private Const mc_strILLEGALCHARSMSG As String = """   [   ]   *"

Private Function AllowEdit(ByVal lngID_in As Long) As Boolean

    Dim lngIdx As Long

    For lngIdx = 1 To m_lngMaxNoEdit
        If lngID_in = ma_lngNoEdit(lngIdx) Then
            AllowEdit = False
            Exit Function
        End If
    Next lngIdx
    AllowEdit = True

End Function

Private Sub cmdCancel_Click()

    g_udtRecordSelector.blnOKClicked = m_blnRecordChanged
    Unload Me

End Sub

Private Sub cmdEdit_Click()

    If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Musician Then
        'musician
        With g_udtMusicianProps
            .blnAdding = False
            .blnOKClicked = False
            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Musicians WHERE MU_Id = " & lstRecs.ItemData(lstRecs.ListIndex), dbOpenDynaset)
            If g_rsGames.EOF = False Then
                g_rsGames.MoveFirst
                .strGroup = g_rsGames!Grp
                .strNick = g_rsGames!Nick
                .strMusician = g_rsGames!Musician
                .strPhoto = g_rsGames!Photo
                .strFocus = "Name"
                frmEditMusicianProps.Show vbModal, Me
                If .blnOKClicked = True Then
                    g_rsGames.Edit
                    g_rsGames!Grp = .strGroup
                    g_rsGames!Musician = .strMusician
                    g_rsGames!Nick = .strNick
                    g_rsGames!Photo = .strPhoto
                    g_rsGames.Update
                    lstRecs.List(lstRecs.ListIndex) = .strMusician
                    g_udtRecordSelector.strReturnedName = .strMusician
                    m_blnRecordChanged = True
                End If
            End If
            g_rsGames.Close
            Set g_rsGames = Nothing
        End With
    Else
        'everything else
        With g_udtInputBox
            .blnAllowEmptyReturn = False
            .blnOKClicked = False
            .lngMaxLength = 255
            .strCancelButtonText = g_strButtonCancelText
            .strCaption = g_xobjFunc.GetINIMessage("Add Edit Record Window", "EditCaption", "Edit Record") & "..."
            .strInvalidChars = ""
            .strOKButtonText = g_xobjFunc.GetINIMessage("General Buttons", "Save", "&Save")
            If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Cracker Then
                'cracker
                .strPrompt = GetINIMessage("Add Edit Record Window", "CrackerPrompt", "Cracker") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Language Then
                'language
                .strPrompt = GetINIMessage("Add Edit Record Window", "LanguagePrompt", "Language") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Programmer Then
                'programmer
                .strPrompt = GetINIMessage("Add Edit Record Window", "ProgrammerPrompt", "Programmer") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Artist Then
                'artist
                .strPrompt = GetINIMessage("Add Edit Record Window", "ArtistPrompt", "Artist") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Developer Then
                'developer
                .strPrompt = GetINIMessage("Add Edit Record Window", "DeveloperPrompt", "Developer") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_License Then
                'license
                .strPrompt = GetINIMessage("Add Edit Record Window", "LicensePrompt", "License") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Rarity Then
                'rarity
                .strPrompt = GetINIMessage("Add Edit Record Window", "RarityPrompt", "Rarity") & ":"
            Else 'publisher
                .strPrompt = GetINIMessage("Add Edit Record Window", "PublisherPrompt", "Publisher") & ":"
            End If
            .strReturnedText = ""
            .strStartingText = lstRecs.List(lstRecs.ListIndex)
            frmInputBox.Show vbModal, Me
            If .blnOKClicked = True Then
                If .strReturnedText <> lstRecs.List(lstRecs.ListIndex) Then
                    'save record
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT " & m_strFieldName & " AS Name FROM " & m_strTableName & " WHERE " & m_strFieldKey & " = " & lstRecs.ItemData(lstRecs.ListIndex), dbOpenDynaset)
                    With g_rsGames
                        If Not .EOF Then
                            .MoveFirst
                            .Edit
                            !Name = g_udtInputBox.strReturnedText
                            .Update
                            lstRecs.List(lstRecs.ListIndex) = g_udtInputBox.strReturnedText
                            g_udtRecordSelector.strReturnedName = g_udtInputBox.strReturnedText
                            m_blnRecordChanged = True
                        End If
                        .Close
                    End With
                    Set g_rsGames = Nothing
                End If
            End If
        End With
    End If
    lstRecs.SetFocus

End Sub

Private Sub cmdFilter_Click()
    Dim strQuery As String
    Dim fldName As Field
    Dim fldID As Field
    Dim lngID As Long
    Dim lngRecordIdx As Long
    Dim lngCharIdx As Long
    
    'test for illegal characters in textbox
    For lngCharIdx = 1 To Len(mc_strILLEGALCHARS)
        If InStr(txtFilter, Mid$(mc_strILLEGALCHARS, lngCharIdx, 1)) <> 0 Then
            MsgBox "The text you typed contains illegal characters!" & vbCrLf & vbCrLf & "You cannot use any of the following characters in the filter: -" & vbCrLf & vbCrLf & mc_strILLEGALCHARSMSG, vbExclamation, gc_strAPPNAME
            txtFilter.SetFocus
            Exit Sub
        End If
    Next lngCharIdx
    
    lstRecs.Clear
    cmdEdit.Enabled = False
    fraRecs.Caption = m_strFrameCaption

    strQuery = "SELECT " & m_strFieldName & " AS Name, " & m_strFieldKey & " as ID FROM " & m_strTableName
    If txtFilter <> "" Then
        strQuery = strQuery & " WHERE (" & m_strFieldName & " LIKE ""*" & txtFilter & "*"")"
        fraRecs.Caption = fraRecs.Caption & " (""" & txtFilter & """)"
    End If
    strQuery = strQuery & " ORDER BY " & m_strFieldName
    Debug.Print strQuery
    
    Set g_rsGames = g_dbGamebase.OpenRecordset(strQuery, dbOpenSnapshot)
    
    If Not g_rsGames.EOF Then
        Set fldName = g_rsGames!Name
        Set fldID = g_rsGames!ID
        g_rsGames.MoveFirst
        With lstRecs
            Do Until g_rsGames.EOF
                .AddItem fldName
                .ItemData(.NewIndex) = fldID
                g_rsGames.MoveNext
            Loop
        End With
        Set fldName = Nothing
        Set fldID = Nothing
    End If
    g_rsGames.Close
    Set g_rsGames = Nothing '

    'find the ID
    lngID = g_udtRecordSelector.lngID
    With lstRecs
        If .ListCount > 0 Then
            For lngRecordIdx = 0 To (.ListCount - 1)
                If .ItemData(lngRecordIdx) = lngID Then
                    .ListIndex = lngRecordIdx
                    Exit For
                End If
            Next lngRecordIdx
            If .ListIndex = gc_lngNOLISTINDEX Then
                .ListIndex = 0
            End If
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End With
    
    txtFilter = ""
    cmdFilter.Default = False
    cmdOK.Default = True
End Sub

Private Sub cmdNew_Click()

    If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Musician Then
        'musician
        With g_udtMusicianProps
            .blnAdding = False
            .blnOKClicked = False
            .strGroup = ""
            .strNick = ""
            .strMusician = ""
            .strPhoto = ""
            .strFocus = "Name"
            frmEditMusicianProps.Show vbModal, Me
            If .blnOKClicked = True Then
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Musicians", dbOpenDynaset)
                g_rsGames.AddNew
                g_rsGames!Grp = .strGroup
                g_rsGames!Musician = .strMusician
                g_rsGames!Nick = .strNick
                g_rsGames!Photo = .strPhoto
                lstRecs.AddItem .strMusician
                lstRecs.ItemData(lstRecs.NewIndex) = g_rsGames!MU_Id
                g_rsGames.Update
                g_rsGames.Close
                Set g_rsGames = Nothing
                lstRecs.ListIndex = gc_lngNOLISTINDEX
                lstRecs.ListIndex = lstRecs.NewIndex
            End If
        End With
    Else
        With g_udtInputBox
            .blnAllowEmptyReturn = False
            .blnOKClicked = False
            .lngMaxLength = 255
            .strCancelButtonText = g_strButtonCancelText
            .strCaption = g_xobjFunc.GetINIMessage("Add Edit Record Window", "NewCaption", "New Record") & "..."
            .strInvalidChars = ""
            .strOKButtonText = g_xobjFunc.GetINIMessage("General Buttons", "Save", "&Save")
            If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Cracker Then
                'cracker
                .strPrompt = GetINIMessage("Add Edit Record Window", "CrackerPrompt", "Cracker") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Language Then
                'language
                .strPrompt = GetINIMessage("Add Edit Record Window", "LanguagePrompt", "Language") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Programmer Then
                'programmer
                .strPrompt = GetINIMessage("Add Edit Record Window", "ProgrammerPrompt", "Programmer") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Artist Then
                'artist
                .strPrompt = GetINIMessage("Add Edit Record Window", "ArtistPrompt", "Artist") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Developer Then
                'developer
                .strPrompt = GetINIMessage("Add Edit Record Window", "DeveloperPrompt", "Developer") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_License Then
                'license
                .strPrompt = GetINIMessage("Add Edit Record Window", "LicensePrompt", "License") & ":"
            ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Rarity Then
                'rarity
                .strPrompt = GetINIMessage("Add Edit Record Window", "RarityPrompt", "Rarity") & ":"
            Else 'publisher
                .strPrompt = GetINIMessage("Add Edit Record Window", "PublisherPrompt", "Publisher") & ":"
            End If
            .strReturnedText = ""
            .strStartingText = ""
            frmInputBox.Show vbModal, Me
            If .blnOKClicked = True Then
                If .strReturnedText <> "" Then
                    'save record
                    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT " & m_strFieldName & " AS Name, " & m_strFieldKey & " AS ID FROM " & m_strTableName, dbOpenDynaset)
                    With g_rsGames
                        .AddNew
                        !Name = g_udtInputBox.strReturnedText
                        lstRecs.AddItem g_udtInputBox.strReturnedText
                        lstRecs.ItemData(lstRecs.NewIndex) = !ID
                        .Update
                        .Close
                    End With
                    Set g_rsGames = Nothing
                    lstRecs.ListIndex = gc_lngNOLISTINDEX
                    lstRecs.ListIndex = lstRecs.NewIndex
                End If
            End If
        End With
    End If
    lstRecs.SetFocus

End Sub

Private Sub cmdOK_Click()

    If (m_blnRecordChanged = True) Or (lstRecs.ItemData(lstRecs.ListIndex) <> g_udtRecordSelector.lngID) Then
        g_udtRecordSelector.blnOKClicked = True
        g_udtRecordSelector.lngID = lstRecs.ItemData(lstRecs.ListIndex)
    End If
    Unload Me

End Sub

Private Sub txtFilter_Change()
    cmdFilter.Enabled = True
    cmdFilter.Default = True
    cmdOK.Default = False
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)

    'test for invalid character pressed
    If InStr(mc_strILLEGALCHARS, Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 4215    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 7695     'The required scalewidth for the form
    Dim fldName As Field
    Dim fldID As Field
    Dim lngID As Long
    Dim lngRecordIdx As Long
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        cmdNew.Caption = .GetINIMessage("Record Selector Window", "NewButton", "&New")
        cmdEdit.Caption = .GetINIMessage("Record Selector Window", "EditButton", "&Edit")
        cmdFilter.Caption = .GetINIMessage("Record Selector Window", "FilterButton", "&Filter")
        txtFilterLabel.Caption = .GetINIMessage("Record Selector Window", "FilterCaption", "Filter by:")
        If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Publisher Then
            Caption = GetINIMessage("Record Selector Window", "CaptionPublisher", "Select Publisher") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Publishers", "Publishers")
        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Programmer Then
            Caption = GetINIMessage("Record Selector Window", "CaptionProgrammer", "Select Programmer") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Programmers", "Programmers")
        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Musician Then
            Caption = GetINIMessage("Record Selector Window", "CaptionMusician", "Select Musician") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Musicians", "Musicians")
        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Language Then
            Caption = GetINIMessage("Record Selector Window", "CaptionLanguage", "Select Language") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Languages", "Languages")

        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Artist Then
            Caption = GetINIMessage("Record Selector Window", "CaptionArtist", "Select Artist") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Artists", "Artists")

        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Developer Then
            Caption = GetINIMessage("Record Selector Window", "CaptionDeveloper", "Select Developer") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Developer", "Developer")
            
        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_License Then
            Caption = GetINIMessage("Record Selector Window", "CaptionLicense", "Select License") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Licenses", "Licenses")
            
        ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Rarity Then
            Caption = GetINIMessage("Record Selector Window", "CaptionRarity", "Select Rarity") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Rarities", "Rarities")

        Else 'crackers
            Caption = GetINIMessage("Record Selector Window", "CaptionCracker", "Select Cracker") & "..."
            fraRecs.Caption = GetINIMessage("Record Selector Window", "Crackers", "Crackers")
        End If
    End With
    m_strFrameCaption = fraRecs.Caption
    
    'set the no-delete array and db sql words
    If g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Publisher Then
        m_lngMaxNoEdit = 2
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 37
        ma_lngNoEdit(2) = 39
        m_strFieldName = "Publisher"
        m_strFieldKey = "PU_Id"
        m_strTableName = "Publishers"
    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Programmer Then
        m_lngMaxNoEdit = 1
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        m_strFieldName = "Programmer"
        m_strFieldKey = "PR_Id"
        m_strTableName = "Programmers"
    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Musician Then
        m_lngMaxNoEdit = 2
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 38
        ma_lngNoEdit(2) = 39
        m_strFieldName = "Musician"
        m_strFieldKey = "MU_Id"
        m_strTableName = "Musicians"
    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Language Then
        m_lngMaxNoEdit = 1
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        m_strFieldName = "Language"
        m_strFieldKey = "LA_Id"
        m_strTableName = "Languages"

    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Artist Then
        m_lngMaxNoEdit = 2
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        ma_lngNoEdit(2) = 2
        m_strFieldName = "Artist"
        m_strFieldKey = "AR_Id"
        m_strTableName = "Artists"

    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Developer Then
        m_lngMaxNoEdit = 2
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        ma_lngNoEdit(2) = 2
        m_strFieldName = "Developer"
        m_strFieldKey = "DE_Id"
        m_strTableName = "Developers"

    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_License Then
        m_lngMaxNoEdit = 1
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        m_strFieldName = "License"
        m_strFieldKey = "LI_Id"
        m_strTableName = "Licenses"

    ElseIf g_udtRecordSelector.eRecordSelectorType = e_RecordSelector_Rarity Then
        m_lngMaxNoEdit = 1
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 1
        m_strFieldName = "Rarity"
        m_strFieldKey = "RA_Id"
        m_strTableName = "Rarities"

    Else
        m_lngMaxNoEdit = 2
        ReDim ma_lngNoEdit(m_lngMaxNoEdit)
        ma_lngNoEdit(1) = 7
        ma_lngNoEdit(2) = 8
        m_strFieldName = "Cracker"
        m_strFieldKey = "CR_Id"
        m_strTableName = "Crackers"
    End If
    ma_lngNoEdit(0) = 0
    
    'read in the table
    Set g_rsGames = g_dbGamebase.OpenRecordset(g_udtRecordSelector.strSQL, dbOpenSnapshot)
    
    If Not g_rsGames.EOF Then
        Set fldName = g_rsGames!Name
        Set fldID = g_rsGames!ID
        g_rsGames.MoveFirst
        With lstRecs
            Do Until g_rsGames.EOF
                .AddItem fldName
                .ItemData(.NewIndex) = fldID
                g_rsGames.MoveNext
            Loop
        End With
        Set fldName = Nothing
        Set fldID = Nothing
    End If
    g_rsGames.Close
    Set g_rsGames = Nothing

    'find the ID
    lngID = g_udtRecordSelector.lngID
    With lstRecs
        If .ListCount > 0 Then
            For lngRecordIdx = 0 To (.ListCount - 1)
                If .ItemData(lngRecordIdx) = lngID Then
                    .ListIndex = lngRecordIdx
                    Exit For
                End If
            Next lngRecordIdx
            If .ListIndex = gc_lngNOLISTINDEX Then
                .ListIndex = 0
            End If
            cmdOK.Enabled = True
            lstRecs.Enabled = True
        End If
    End With
    
    'finish up
    lstRecs.Visible = True
    Screen.MousePointer = vbDefault
    g_udtRecordSelector.blnOKClicked = False
    m_blnRecordChanged = False

End Sub

Private Sub lstRecs_Click()

    With lstRecs
        If .ListIndex > gc_lngNOLISTINDEX Then
            cmdEdit.Enabled = AllowEdit(.ItemData(.ListIndex))
            g_udtRecordSelector.strReturnedName = .List(.ListIndex)
            cmdOK.Enabled = True
        Else
            cmdEdit.Enabled = False
            cmdOK.Enabled = False
        End If
    End With

End Sub


