VERSION 5.00
Begin VB.Form frmRecordSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Selector"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmRecordSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRecs 
      Caption         =   "Select Record"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   4800
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
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
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
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
' HVSC2GB64 - Tool to update GB64 database with
' the latest version of the HVSC
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

    g_udtRecordSelector.blnOKClicked = False
    Unload Me

End Sub

Private Sub cmdEdit_Click()

    Dim rsGames As DAO.Recordset

    'edit musician
    With g_udtMusicianProps2
        .blnAdding = False
        .blnOKClicked = False
        Set rsGames = g_dbGB64.OpenRecordset("SELECT * FROM Musicians WHERE MU_Id = " & lstRecs.ItemData(lstRecs.ListIndex), dbOpenDynaset)
        If rsGames.EOF = False Then
            rsGames.MoveFirst
            .strGroup = rsGames!Grp
            .strNick = rsGames!Nick
            .strMusician = rsGames!Musician
            .strPhoto = rsGames!Photo
            .strFocus = "Name"
            frmEditMusicianProps.Show vbModal, Me
            If .blnOKClicked = True Then
                rsGames.Edit
                rsGames!Grp = .strGroup
                rsGames!Musician = .strMusician
                rsGames!Nick = .strNick
                rsGames!Photo = .strPhoto
                rsGames.Update
                lstRecs.List(lstRecs.ListIndex) = .strMusician
                g_udtRecordSelector.strReturnedName = .strMusician
                If g_udtRecordSelector.blnGNPChanged = False Then
                    g_udtRecordSelector.blnGNPChanged = .blnGNPChanged
                End If
                m_blnRecordChanged = True
            End If
        End If
        rsGames.Close
        Set rsGames = Nothing
    End With
    lstRecs.SetFocus

End Sub

Private Sub cmdNew_Click()

    Dim rsGames As DAO.Recordset

    'new musician
    With g_udtMusicianProps2
        .blnAdding = False
        .blnOKClicked = False
        .strGroup = ""
        .strNick = ""
        .strMusician = ""
        .strPhoto = ""
        .strFocus = "Name"
        frmEditMusicianProps.Show vbModal, Me
        If .blnOKClicked = True Then
            Set rsGames = g_dbGB64.OpenRecordset("SELECT * FROM Musicians", dbOpenDynaset)
            rsGames.AddNew
            rsGames!Grp = .strGroup
            rsGames!Musician = .strMusician
            rsGames!Nick = .strNick
            rsGames!Photo = .strPhoto
            lstRecs.AddItem .strMusician
            lstRecs.ItemData(lstRecs.NewIndex) = rsGames!MU_Id
            rsGames.Update
            rsGames.Close
            Set rsGames = Nothing
            lstRecs.ListIndex = -1
            lstRecs.ListIndex = lstRecs.NewIndex
            g_udtRecordSelector.blnGNPChanged = True
        End If
    End With
    lstRecs.SetFocus

End Sub

Private Sub cmdOK_Click()

    Dim rsGames As DAO.Recordset

    g_udtRecordSelector.blnOKClicked = True
    g_udtRecordSelector.lngID = lstRecs.ItemData(lstRecs.ListIndex)
    Set rsGames = g_dbGB64.OpenRecordset("SELECT Grp, Nick, Photo FROM Musicians WHERE MU_Id=" & g_udtRecordSelector.lngID, dbOpenSnapshot)
    With rsGames
        .MoveFirst
        g_udtRecordSelector.strGroup = !Grp
        g_udtRecordSelector.strNick = !Nick
        g_udtRecordSelector.strPhoto = !Photo
        .Close
    End With
    Set rsGames = Nothing
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 4215    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6255     'The required scalewidth for the form
    Dim fldName As Field
    Dim fldID As Field
    Dim lngID As Long
    Dim lngRecordIdx As Long
    Dim rsGames As DAO.Recordset
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    With g_xobjFunc
        Caption = "Select GB64 Musician..."
        fraRecs.Caption = "GB64 Musicians"
    End With

    'set the no-delete array and db sql words
    m_lngMaxNoEdit = 2
    ReDim ma_lngNoEdit(m_lngMaxNoEdit)
    ma_lngNoEdit(1) = 38
    ma_lngNoEdit(2) = 39
    m_strFieldName = "Musician"
    m_strFieldKey = "MU_Id"
    m_strTableName = "Musicians"
    ma_lngNoEdit(0) = 0
    
    'read in the table
    Set rsGames = g_dbGB64.OpenRecordset(g_udtRecordSelector.strSQL, dbOpenSnapshot)
    If Not rsGames.EOF Then
        Set fldName = rsGames!Name
        Set fldID = rsGames!ID
        rsGames.MoveFirst
        With lstRecs
            Do Until rsGames.EOF
                .AddItem fldName
                .ItemData(.NewIndex) = fldID
                rsGames.MoveNext
            Loop
        End With
        Set fldName = Nothing
        Set fldID = Nothing
    End If
    rsGames.Close
    Set rsGames = Nothing

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
            If .ListIndex = -1 Then
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
    g_udtRecordSelector.blnGNPChanged = True
    m_blnRecordChanged = False

End Sub

Private Sub lstRecs_Click()

    With lstRecs
        If .ListIndex > -1 Then
            cmdEdit.Enabled = AllowEdit(.ItemData(.ListIndex))
            g_udtRecordSelector.strReturnedName = .List(.ListIndex)
            cmdOK.Enabled = True
        Else
            cmdEdit.Enabled = False
            cmdOK.Enabled = False
        End If
    End With

End Sub


