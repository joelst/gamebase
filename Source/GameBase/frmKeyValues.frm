VERSION 5.00
Begin VB.Form frmKeyValues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GEMUS ""key=value"" pairs"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   Icon            =   "frmKeyValues.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGame 
      Caption         =   "Current Selected Game"
      Height          =   7095
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdPeek 
         Caption         =   "&Peek (F1)"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   6600
         Width           =   1095
      End
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4695
         TabIndex        =   19
         Top             =   240
         Width           =   4695
         Begin VB.OptionButton optEdit 
            Caption         =   "Edit All Games in Current View"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   21
            Top             =   90
            Width           =   2535
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Edit Current Game Only"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   90
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.OptionButton optSave 
         Caption         =   "SAVE to games in current view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   5280
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optSave 
         Caption         =   "UPDATE games in current view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   5520
         Width           =   2775
      End
      Begin VB.OptionButton optSave 
         Caption         =   "ADD to games in current view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   5760
         Width           =   2655
      End
      Begin VB.OptionButton optSave 
         Caption         =   "REMOVE matching keys from games in current view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   6000
         Width           =   4335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   6600
         Width           =   1095
      End
      Begin VB.OptionButton optSave 
         Caption         =   "SAVE for current game only"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   15
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optSave 
         Caption         =   "REMOVE matching key=values from games in current view"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   6240
         Width           =   4695
      End
      Begin VB.TextBox txtPairs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   0
         Top             =   630
         Width           =   4695
      End
   End
   Begin VB.Frame fraSets 
      Caption         =   "KVSets"
      Height          =   7095
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdRename 
         Caption         =   "&Rename"
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   1080
      End
      Begin VB.TextBox txtSet 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   120
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   1560
         Width           =   4695
      End
      Begin VB.ComboBox lstSets 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Saved OK"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1260
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmKeyValues"
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

Private m_strPreChangedPairs As String
Private m_blnByCode As Boolean
Private m_strPreChangedSet As String
Private m_lngPreviousSetListIndex As Long
Private m_blnFormIsLoading As Boolean
Private m_blnIsPeekPossible As Boolean
Private Sub AddSet()

    Dim blnFound As Boolean
    Dim lngIdx As Long
    Dim strSetName As String
    Dim strSetNameUCase As String
    Dim strText As String
    Dim strError As String

    With g_udtInputBox
        .blnAllowEmptyReturn = False
        .blnOKClicked = False
        .lngMaxLength = 64
        .strCancelButtonText = g_strButtonCancelText
        .strCaption = "New GEMUS KVSet"
        .strInvalidChars = "[""];"
        .strOKButtonText = g_strButtonOKText
        .strPrompt = "Enter the name of the new KVSet:"
        .strReturnedText = ""
        .strStartingText = ""
        frmInputBox.Show vbModal, Me
        If .blnOKClicked = True Then
            strSetName = LCase$(Trim$(.strReturnedText))
            strSetNameUCase = UCase$(strSetName)
            If strSetName = "" Then
                MsgBox "You must enter a name for the new KVSet", vbExclamation, gc_strAPPNAME
            ElseIf strSetNameUCase = "GLOBAL" Then
                MsgBox "You cannot call a new KVSet ""global""", vbExclamation, gc_strAPPNAME
            Else
                'only add it if there isn't one existing of the same name
                blnFound = False
                With lstSets
                    For lngIdx = 0 To .ListCount - 1
                        If strSetNameUCase = Trim$(UCase$(.List(lngIdx))) Then
                            blnFound = True
                            Exit For
                        End If
                    Next lngIdx
                End With
                If blnFound = True Then
                    MsgBox "The KVSet """ & strSetName & """ already exists", vbExclamation, gc_strAPPNAME
                Else
                    m_strPreChangedSet = ""
                    txtSet.Text = ""
                    lstSets.AddItem strSetName
                    lstSets.ListIndex = -1
                    lstSets.ListIndex = lstSets.ListCount - 1
                    lblStatus.Caption = "KVSet """ & strSetName & """ Added"
                    strText = ""
                    Call g_xobjGEMUS.SaveKVSet(strSetName, strText, strError)
                End If
            End If
        End If
    End With
    
    txtSet.SetFocus

End Sub

Private Sub RenameSet(strSetName_in As String)

    Dim blnFound As Boolean
    Dim lngIdx As Long
    Dim strSetName As String
    Dim strSetNameUCase As String
    Dim strError As String

    Dim mbRes As VbMsgBoxResult
    
    mbRes = MsgBox( _
            "There may be some games that use the KVSet """ & strSetName_in & """." & vbCrLf & vbCrLf & _
                "If so, you will have to update those games manually." & vbCrLf & vbCrLf & _
                "(A future version of GameBase might do that for you.)" & vbCrLf & vbCrLf & _
                "ARE YOU SURE you still want to rename it?", _
            vbYesNo Or vbDefaultButton1, _
            gc_strAPPNAME _
        )

    If mbRes = vbYes Then
        With g_udtInputBox
            .blnAllowEmptyReturn = False
            .blnOKClicked = False
            .lngMaxLength = 64
            .strCancelButtonText = g_strButtonCancelText
            .strCaption = "Rename GEMUS KVSet """ & strSetName_in & """"
            .strInvalidChars = "[]"";"
            .strOKButtonText = g_strButtonOKText
            .strPrompt = "Enter the new name for the KVSet:"
            .strReturnedText = ""
            .strStartingText = strSetName_in
            frmInputBox.Show vbModal, Me
            If .blnOKClicked = True Then
                strSetName = Trim$(.strReturnedText)
                strSetNameUCase = UCase$(strSetName)
                If strSetName = "" Then
                    MsgBox "You must enter a new name for the KVSet", vbExclamation, gc_strAPPNAME
                ElseIf strSetNameUCase = "GLOBAL" Then
                    MsgBox "You cannot rename a KVSet to ""global""", vbExclamation, gc_strAPPNAME
                Else
                    'let gemus rename it, and if successful, change the name in the list
                    If g_xobjGEMUS.RenameKVSet(strSetName_in, strSetName, strError) Then
                        lstSets.List(lstSets.ListIndex) = strSetName
                        lblStatus.Caption = "KVSet """ & strSetName_in & """ renamed to """ & strSetName & """"
                    Else
                        MsgBox "Unable to rename the KVSet """ & strSetName_in & """: " & vbCrLf & vbCrLf & strError, vbExclamation, gc_strAPPNAME
                    End If
                End If
            End If
        End With
    End If
    
    txtSet.SetFocus

End Sub
Private Sub DeleteSet(strSetName_in As String)

    Dim blnFound As Boolean
    Dim lngIdx As Long
    Dim strSetName As String
    Dim strSetNameUCase As String
    Dim strError As String
    Dim lngCurrentListIndex As Long

    Dim mbRes As VbMsgBoxResult
    
    mbRes = MsgBox( _
            "There may be some games that use the KVSet """ & strSetName_in & """." & vbCrLf & vbCrLf & _
                "If so, you will have to update those games manually." & vbCrLf & vbCrLf & _
                "(A future version of GameBase might do that for you.)" & vbCrLf & vbCrLf & _
                "ARE YOU SURE you still want to delete it?", _
            vbYesNo Or vbDefaultButton1, _
            gc_strAPPNAME _
        )

    If mbRes = vbYes Then
        'let gemus delete it, and if successful, remove the name in the list
        If g_xobjGEMUS.DeleteKVSet(strSetName_in, strError) Then
            With lstSets
                lngCurrentListIndex = .ListIndex
                Call .RemoveItem(lngCurrentListIndex)
                .ListIndex = -1
                If lngCurrentListIndex < .ListCount Then
                    .ListIndex = lngCurrentListIndex
                Else
                    .ListIndex = lngCurrentListIndex - 1
                End If
            End With
            lblStatus.Caption = "KVSet """ & strSetName_in & """ Deleted"
        Else
            MsgBox "Unable to delete the KVSet """ & strSetName_in & """: " & vbCrLf & vbCrLf & strError, vbExclamation, gc_strAPPNAME
        End If
    End If
    
    txtSet.SetFocus

End Sub
Private Function SaveSet(strSetName_in As String) As Boolean

    Dim strText As String
    Dim strError As String
    Dim strSetName As String

    lblStatus.Caption = ""
    Screen.MousePointer = vbHourglass
    DoEvents

    strSetName = strSetName_in
    If strSetName = "<GLOBAL>" Then strSetName = "global"
    strText = txtSet.Text

    If g_xobjGEMUS.SaveKVSet(strSetName, strText, strError) Then
        m_strPreChangedSet = txtSet.Text
        lblStatus.Caption = "Saved OK"
        cmdSave.Enabled = False
        SaveSet = True
    Else
        MsgBox "Unable to save the KVSet """ & strSetName_in & """: " & vbCrLf & vbCrLf & strError, vbExclamation, gc_strAPPNAME
        SaveSet = False
    End If
    
    txtSet.SetFocus
    Screen.MousePointer = vbDefault

End Function
Private Function SaveSetIfChangedAndUserWantsTo() As Boolean

    SaveSetIfChangedAndUserWantsTo = True

    With lstSets
        If (Not m_blnFormIsLoading) And (m_lngPreviousSetListIndex > -1) Then
            If txtSet.Text <> m_strPreChangedSet Then
                If MsgBox( _
                    "The KVSet """ & .List(m_lngPreviousSetListIndex) & """ has changed." & vbCrLf & vbCrLf & "Do you want to save it?", _
                    vbQuestion Or vbYesNo, _
                    gc_strAPPNAME) = vbYes _
                Then
                    If Not SaveSet(.List(m_lngPreviousSetListIndex)) Then
                        SaveSetIfChangedAndUserWantsTo = False
                    End If
                End If
            End If
        End If
    End With

End Function

Private Sub SetOptSaveForecolor(intSelectedIndex As Integer)

    Dim lngIdx As Long
    
    For lngIdx = 0 To optSave.Count - 1
        optSave(lngIdx).ForeColor = IIf(lngIdx = intSelectedIndex, &HFF0000, Me.ForeColor)
    Next lngIdx

End Sub

Private Sub cmdCancel_Click()

    If m_strPreChangedPairs <> txtPairs.Text Then
        If MsgBox("Discard changed Key=Value Pairs?", vbQuestion Or vbYesNo Or vbDefaultButton2, gc_strAPPNAME) = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me
End Sub

Private Sub cmdDelete_Click()

    lblStatus.Caption = ""
    
    If lstSets.ListIndex = 0 Then Exit Sub
    
    Call DeleteSet(lstSets.List(lstSets.ListIndex))

End Sub

Private Sub cmdNew_Click()

    If Not SaveSetIfChangedAndUserWantsTo Then Exit Sub
    lblStatus.Caption = ""
    Call AddSet

End Sub

Private Sub cmdOK_Click()

    Dim msgBoxResult As VbMsgBoxResult
    Dim q As String
    
    'save the key=value pairs
    cmdOK.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    If optEdit(0).Value = True Then
        'current game update
        g_udtGame.strPairs = txtPairs.Text
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Gemus FROM Games WHERE GA_Id=" & g_udtGame.lngID, dbOpenDynaset)
        With g_rsGames
            If Not .EOF Then
                .MoveFirst
                .Edit
                !Gemus = txtPairs.Text
                .Update
            End If
            .Close
        End With
        Set g_rsGames = Nothing
    Else

        'build question string for msgbox
        If optSave(1).Value = True Then
            q = "You are about to COMPLETELY OVERWRITE the key=value pairs for ALL GAMES in the current view."
        ElseIf optSave(2).Value = True Then
            q = "You are about to ADD these key=value pairs to ALL GAMES in the current view." & _
                vbCrLf & vbCrLf & _
                "Where keys already exist for a game, they will be overwritten with these values."
        ElseIf optSave(3).Value = True Then
            q = "You are about to ADD these key=value pairs to ALL GAMES in the current view." & _
                vbCrLf & vbCrLf & _
                "Where keys already exist for a game, their values will not be updated."
        ElseIf optSave(4).Value = True Then
            q = "You are about to REMOVE these key=value pairs from ALL GAMES in the current view." & _
                vbCrLf & vbCrLf & _
                "Keys will be removed regardless of their values."
        Else
            q = "You are about to REMOVE these key=value pairs from ALL GAMES in the current view." & _
                vbCrLf & vbCrLf & _
                "Keys will only be removed if their values match these ones exactly."
        End If

        msgBoxResult = MsgBox( _
                q & vbCrLf & vbCrLf & "ARE YOU SURE you want to do this?", _
                vbYesNo Or vbQuestion Or vbDefaultButton2, _
                gc_strAPPNAME _
            )

        If (msgBoxResult = vbCancel) Or (msgBoxResult = vbNo) Then
            Screen.MousePointer = vbDefault
            cmdOK.Enabled = True
            Exit Sub
        Else
            'batch update for all games in current view
            g_strUpdateKVText = txtPairs.Text
            If optSave(1).Value = True Then
                g_UpdateKVTypeEnum = e_Save_All
            ElseIf optSave(2).Value = True Then
                g_UpdateKVTypeEnum = e_Add_All_Overwrite
            ElseIf optSave(3).Value = True Then
                g_UpdateKVTypeEnum = e_Add_All_Leave
            ElseIf optSave(4).Value = True Then
                g_UpdateKVTypeEnum = e_Remove_All_Matching_Keys
            Else
                g_UpdateKVTypeEnum = e_Remove_All_Matching_Exactly
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Unload Me

End Sub


Private Sub cmdPeek_Click()

    Dim lngEmuIdx As Long
    Dim strError As String
    Dim lngGPIdx As Long
    Dim strExt As String
    Dim lngDotPos As Long
    Dim strFileToRun As String
    Dim eExtractResult As Gemus.GB64ExtractResultEnum
    Dim Extractor As Gemus.clsExtract
    Dim strExtractedFilename As String
    Dim strFileList As String
    Dim strDefaultFile As String
    Dim strStatus As String
    Dim strArchiveCaption As String
    Dim strArchivePrompt As String
    Dim strStatusArchiveContents As String
    Dim strImageName As String
    Dim lngTotalPlayed As Long
    Dim strOldSBText As String
    Dim strKV As String
    Dim strPairs As String
    Dim mbRes As VbMsgBoxResult

    If cmdSave.Enabled = True Then
        mbRes = MsgBox("Do you want to save the currently selected KVSet?", vbQuestion Or vbYesNoCancel, gc_strAPPNAME)
        If mbRes = vbCancel Then Exit Sub
        If mbRes = vbYes Then cmdSave.Value = True
    End If

    cmdPeek.Enabled = False
    MousePointer = vbHourglass
    DoEvents
    strPairs = txtPairs.Text

    'find actual game path\file
    If g_udtGame.blnGameFileExists = True Then
        strError = FindGameFile(lngGPIdx, g_udtGame.strGameFilename)
        If strError <> "" Then
            'not found or aborted
            If strError <> "~~ABORTED~~" Then
                MsgBox "Could not Peek, need game file to exist in your game paths!", vbExclamation, gc_strAPPNAME
            End If
        Else
            'game found!
            strDefaultFile = g_udtGame.strFileToRun
            strFileToRun = ga_udtRegPathGame(lngGPIdx).strWithBackSlash & g_udtGame.strGameFilename
            With g_udtGame

                'extract and run game
                Call g_xobjGEMUS.SetRezipNotify(g_blnZipNotify)
                Call g_xobjGEMUS.SetRezip(g_blnReZip, g_blnRezipAsk, ge_RezipAdditional)
                lngEmuIdx = -1 'default emulator
                strError = g_xobjGEMUS.SimulateRunEmulator(strKV, lngEmuIdx, strFileToRun _
                , .lngGameFilenameIndex, strDefaultFile, g_udtPathGBGame.strRealPath, True, .lngPlayersFrom _
                , .lngPlayersTo, .lngControl, .strComment, .strVersionComment, .lngNTSC, IIf(.lngTrueDriveEmulation = 1, True, False), strPairs)
            End With

            If strError = "" Then
                Load frmKVPeek
                frmKVPeek.Caption = "[Esc or F1 to close] Applied key=value pairs for " & fraGame.Caption
                frmKVPeek.txtPairs.Text = strKV
                frmKVPeek.Show vbModal, Me
            Else
                MsgBox "Could not Peek!" & vbCrLf & vbCrLf & strError, vbInformation, gc_strAPPNAME
            End If
        End If
    Else
        MsgBox "Could not Peek, need game file to exist!", vbExclamation, gc_strAPPNAME
    End If

    MousePointer = vbDefault
    cmdPeek.Enabled = True

End Sub

Private Sub cmdRename_Click()
    
    If lstSets.ListIndex = -1 Then Exit Sub
    If lstSets.ListIndex = 0 Then
        MsgBox "You cannot rename the <GLOBAL> KVSet.", vbExclamation, gc_strAPPNAME
        Exit Sub
    End If
    
    If Not SaveSetIfChangedAndUserWantsTo Then Exit Sub
    lblStatus.Caption = ""
    Call RenameSet(lstSets.List(lstSets.ListIndex))

End Sub

Private Sub cmdSave_Click()

    If lstSets.ListIndex > -1 Then
        Call SaveSet(lstSets.List(lstSets.ListIndex))
    End If

End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 Then
        If KeyCode = 112 Then
            If cmdPeek.Enabled = True Then
                cmdPeek.Value = True
            End If
        End If
    End If

End Sub

Private Sub Form_Load()

    Dim a_strSets() As String
    Dim lngTotalSets As Long
    Dim lngIdx As Long
    Dim strLastViewedSet As String

    m_blnByCode = False
    m_blnFormIsLoading = True
    m_blnIsPeekPossible = frmMain.cmdPlay.Enabled
    cmdPeek.Enabled = m_blnIsPeekPossible

    g_UpdateKVTypeEnum = e_Save_Current
    g_strUpdateKVText = ""

    Width = 10335 + (Width - ScaleWidth)
    Height = 7335 + (Height - ScaleHeight)

    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText

    fraGame.Caption = frmMain.lblName
    txtPairs.Text = g_udtGame.strPairs
    txtPairs.SelStart = Len(txtPairs)
    m_strPreChangedPairs = txtPairs.Text
    
    Call SetOptSaveForecolor(0)

    'get all kvsets
    ReDim a_strSets(0)
    lngTotalSets = g_xobjGEMUS.GetKVSetNames(a_strSets)

    'get the last viewed set
'    Call g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM( _
'            "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", _
'            "LastPerGameKVSETViewed", _
'            strLastViewedSet, _
'            True, _
'            "global" _
'        )
    strLastViewedSet = "global"

    'fill the listbox
    With lstSets
        .Enabled = False
        .Visible = False
        .ListIndex = -1
        .AddItem "<GLOBAL>"
        If strLastViewedSet = "global" Then .ListIndex = 0
        For lngIdx = 0 To lngTotalSets - 1
            .AddItem a_strSets(lngIdx)
            If strLastViewedSet = a_strSets(lngIdx) Then .ListIndex = lngIdx + 1
        Next lngIdx
        If .ListIndex = -1 Then .ListIndex = 0
        .Enabled = True
        .Visible = True
        m_lngPreviousSetListIndex = .ListIndex
    End With

    m_blnFormIsLoading = False

End Sub





Private Sub Form_Unload(Cancel As Integer)

    Dim strLastViewedSet As String

    If Not SaveSetIfChangedAndUserWantsTo Then Exit Sub

    'save the last viewed set
    With lstSets
        If .Visible = True And .ListIndex > -1 Then
            strLastViewedSet = IIf(.ListIndex = 0, "global", .List(.ListIndex))
            Call g_xobjRegistry.SaveSetting( _
                    "HKEY_CURRENT_USER", _
                    "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", _
                    "LastPerGameKVSETViewed", _
                    strLastViewedSet, _
                    True _
                )
        End If
    End With

End Sub


Private Sub lstSets_Click()

    Dim a_udtKV() As NotesKeysType
    Dim lngTotalKV As Long
    Dim lngIdx As Long
    Dim strError As String

    If m_blnByCode = True Then Exit Sub

    If Not SaveSetIfChangedAndUserWantsTo Then
        m_blnByCode = True
        lstSets.ListIndex = m_lngPreviousSetListIndex
        m_blnByCode = False
        Exit Sub
    End If
    
    lblStatus.Caption = ""
    txtSet.Text = ""
    cmdRename.Enabled = False
    cmdDelete.Enabled = False

    With lstSets
    
        If .ListIndex = -1 Then Exit Sub
        
        If .ListIndex = 0 Then
            Call g_xobjGEMUS.GetKVSetAsText("global", m_strPreChangedSet, strError)
        Else
            Call g_xobjGEMUS.GetKVSetAsText(.List(.ListIndex), m_strPreChangedSet, strError)
            cmdRename.Enabled = True
            cmdDelete.Enabled = True
        End If

        m_lngPreviousSetListIndex = .ListIndex

    End With

    txtSet.Text = m_strPreChangedSet
    cmdSave.Enabled = False
    If Not m_blnFormIsLoading Then lstSets.SetFocus

End Sub

Private Sub optEdit_Click(Index As Integer)

    Dim lngIdx As Long

    For lngIdx = 1 To 5
        optSave(lngIdx).Enabled = (Index = 1)
        If Index = 1 Then
            If optSave(lngIdx).Value = True Then
                Call SetOptSaveForecolor(CInt(lngIdx))
            End If
        End If
    Next lngIdx
    
    With txtPairs
        .SelStart = 0
        If Index = 1 Then
            .SelLength = Len(.Text)
        Else
            .SelLength = 0
        End If
        .SetFocus
    End With
    
    If Index = 0 Then
        cmdPeek.Enabled = m_blnIsPeekPossible
    Else
        cmdPeek.Enabled = False
    End If

End Sub

Private Sub optSave_Click(Index As Integer)

    Call SetOptSaveForecolor(Index)

End Sub

Private Sub txtPairs_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        With txtPairs
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub


Private Sub txtSet_Change()

    cmdSave.Enabled = True
    lblStatus.Caption = ""

End Sub


Private Sub txtSet_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 2 And KeyCode = 65 Then
        With txtSet
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If

End Sub


