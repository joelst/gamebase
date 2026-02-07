VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMusicProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Music Info Properties"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmMusicProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProps 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Include"
      TabPicture(0)   =   "frmMusicProps.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAll"
      Tab(0).Control(1)=   "cmdNone"
      Tab(0).Control(2)=   "fraUpdate"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fields"
      TabPicture(1)   =   "frmMusicProps.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPrompt(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblMusician"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkAdult"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkFav"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdMusician"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdMusician 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkFav 
         Caption         =   "Fa&vourite Tune"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkAdult 
         Caption         =   "Ad&ult Tune"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Fields to Update"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   5055
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Musician"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Favourite Tune"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Adult Tune"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   -73200
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblMusician 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   4695
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Musician:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Label lblDummy 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmMusicProps"
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

Private m_blnSeenOptions As Boolean
Private m_blnLoading As Boolean
Private m_blnStartup As Boolean
Private m_lngMusicianID As Long
Private m_strMusician As String
Private m_blnFirstTime As Boolean

Private Sub SetOKButton()

    cmdOK.Enabled = m_blnSeenOptions

End Sub

Private Sub chkUpdate_Click(Index As Integer)

    Dim lngIdx As Long
    Dim blnOK As Boolean

    If g_udtEditMusicInfo.blnEditAll = True Then
        If Index = 0 Then
            If chkUpdate(0).Value = vbChecked Then
                lblMusician.BackColor = lblDummy.BackColor
                lblMusician.Enabled = True
                cmdMusician.Enabled = True
            Else
                lblMusician.BackColor = BackColor
                cmdMusician.Enabled = False
                lblMusician.Enabled = False
            End If
        End If
        If Index = 1 Then
            If chkUpdate(1).Value = vbChecked Then
                chkFav.Enabled = True
            Else
                chkFav.Enabled = False
            End If
        End If
        If Index = 2 Then
            If chkUpdate(2).Value = vbChecked Then
                chkAdult.Enabled = True
            Else
                chkAdult.Enabled = False
            End If
        End If
        
        'any selected?
        blnOK = False
        For lngIdx = 0 To 2
            If chkUpdate(lngIdx).Value = vbChecked Then
                blnOK = True
                Exit For
            End If
        Next lngIdx
        If Not blnOK Then
            cmdOK.Enabled = False
            tabProps.TabEnabled(1) = False
        Else
            tabProps.TabEnabled(1) = True
            Call SetOKButton
        End If
    End If

End Sub



Private Sub cmdAll_Click()

    Dim lngIdx As Long
    
    For lngIdx = 0 To 2
        chkUpdate(lngIdx).Value = vbChecked
    Next lngIdx

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdMusician_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Musician
        .lngID = m_lngMusicianID
        .strSQL = "SELECT Musician AS Name, MU_Id as ID FROM Musicians ORDER BY Musician"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditMusicInfo.blnTableEdited = True
            m_lngMusicianID = .lngID
            m_strMusician = .strReturnedName
            lblMusician.Caption = m_strMusician
        End If
    End With
    cmdMusician.SetFocus

End Sub

Private Sub cmdNone_Click()

    Dim lngIdx As Long
    
    For lngIdx = 0 To 2
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx

End Sub

Private Sub cmdOK_Click()

    'confirmation
    If g_udtEditMusicInfo.blnEditAll Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "MusicAreYouSure", "Are you sure you want to update ALL MUSIC in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If
    'save all info
    With g_udtEditMusicInfo
        .blnOKClicked = True
        If .blnEditAll = True Then
            .blnUpdateMusician = (chkUpdate(0).Value = vbChecked)
            .blnUpdateFav = (chkUpdate(1).Value = vbChecked)
            .blnUpdateAdult = (chkUpdate(2).Value = vbChecked)
        End If
        .lngMusicianID = m_lngMusicianID
        .blnFav = (chkFav.Value = vbChecked)
        .blnAdult = (chkAdult.Value = vbChecked)
    End With

    Unload Me

End Sub


Private Sub Form_Activate()

    'each time the form is shown

    Dim lngIdx As Long
    Dim lngLengthAbs As Long

    If m_blnFirstTime = True Then
        m_blnFirstTime = False
    Else
        Exit Sub
    End If

    m_blnStartup = True
    m_blnLoading = False

    'clear all checkmarks
    For lngIdx = 0 To 2
        chkUpdate(lngIdx).Value = vbChecked
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx

    'are we changing all the games?
    If g_udtEditMusicInfo.blnEditAll = False Then
        lblMusician.BackColor = lblDummy.BackColor
        lblMusician.Enabled = True
        cmdMusician.Enabled = True
        chkFav.Enabled = True
        chkAdult.Enabled = True
        tabProps.TabEnabled(0) = False
        tabProps.TabEnabled(1) = True
        tabProps.Tab = 1
    Else
        tabProps.TabEnabled(0) = True
        tabProps.TabEnabled(1) = False
        tabProps.Tab = 0
    End If

    'init props
    chkFav.Value = vbUnchecked
    chkAdult.Value = vbUnchecked

    'fill props
    lblMusician.Caption = m_strMusician
    chkFav.Value = Abs(g_udtEditMusicInfo.blnFav)
    chkAdult.Value = Abs(g_udtEditMusicInfo.blnAdult)

    'set the focus
    If tabProps.Tab = 1 Then
        If g_udtEditMusicInfoFocus = e_MusicFocus_Musician Then
            cmdMusician.SetFocus
        ElseIf g_udtEditMusicInfoFocus = e_MusicFocus_Fav Then
            chkFav.SetFocus
        Else
            chkAdult.SetFocus
        End If
    Else
        chkUpdate(0).SetFocus
    End If

    'set ok button and finish off
    m_blnSeenOptions = Not g_udtEditMusicInfo.blnEditAll
    Call SetOKButton
    g_udtEditMusicInfo.blnOKClicked = False
    m_blnStartup = False
    g_udtEditMusicInfo.blnTableEdited = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 6375    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5775     'The required scalewidth for the form
    Dim lngIdx As Long

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    m_blnLoading = True
    m_lngMusicianID = g_udtEditMusicInfo.lngMusicianID

    'read in the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Music Info Properties Window", "Caption", "Music Info Properties") & "..."
        tabProps.TabCaption(0) = .GetINIMessage("Music Info Properties Window", "IncludeTab", "&Include")
        tabProps.TabCaption(1) = .GetINIMessage("Music Info Properties Window", "FieldsTab", "&Fields")
        cmdAll.Caption = .GetINIMessage("Music Info Properties Window", "SelectAllButton", "Select &All")
        cmdNone.Caption = .GetINIMessage("Music Info Properties Window", "SelectNoneButton", "Select &None")
        fraUpdate.Caption = .GetINIMessage("Music Info Properties Window", "FieldsToUpdate", "Fields to Update")
        chkUpdate(0).Caption = GetINIMessage("Music Info Properties Window", "UpdateMusicianBox", "Musician")
        chkUpdate(1).Caption = GetINIMessage("Music Info Properties Window", "UpdateFavBox", "Favourite Tune")
        chkUpdate(2).Caption = GetINIMessage("Music Info Properties Window", "UpdateAdultBox", "Adult Tune")
        lblPrompt(2).Caption = GetINIMessage("Music Info Properties Window", "MusicianPrompt", "&Musician") & ":"
        chkFav.Caption = GetINIMessage("Music Info Properties Window", "FavPrompt", "Fa&vourite Tune")
        chkAdult.Caption = GetINIMessage("Music Info Properties Window", "AdultPrompt", "Ad&ult Tune")
    End With

    'read in the musician
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MU_Id as ID, Musician as Name FROM Musicians WHERE MU_Id = " & m_lngMusicianID, dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            m_strMusician = !Name
        End If
        .Close
    End With
    Set g_rsGames = Nothing

    m_blnFirstTime = True

End Sub

Private Sub tabProps_Click(PreviousTab As Integer)

    If tabProps.Tab <> PreviousTab Then
        If tabProps.Tab = 1 Then
            If Not m_blnLoading Then
                If cmdMusician.Enabled = True Then
                    cmdMusician.SetFocus
                ElseIf chkFav.Enabled = True Then
                    chkFav.SetFocus
                ElseIf chkAdult.Enabled = True Then
                    chkAdult.SetFocus
                End If
            End If
        Else
            If chkUpdate(0).Enabled = True And (Not m_blnLoading) Then
                chkUpdate(0).SetFocus
            End If
        End If
    End If
    If tabProps.Tab = 1 Then
        m_blnSeenOptions = True
        Call SetOKButton
    End If

End Sub

