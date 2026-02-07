VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonalProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personal Info Properties"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmPersonalProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDummy 
      Enabled         =   0   'False
      Height          =   255
      ItemData        =   "frmPersonalProps.frx":08CA
      Left            =   120
      List            =   "frmPersonalProps.frx":08CC
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   5880
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProps 
      Height          =   5535
      Left            =   120
      TabIndex        =   23
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
      TabPicture(0)   =   "frmPersonalProps.frx":08CE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraUpdate"
      Tab(0).Control(1)=   "cmdNone"
      Tab(0).Control(2)=   "cmdAll"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fields"
      TabPicture(1)   =   "frmPersonalProps.frx":08EA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPrompt(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPrompt(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPrompt(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgRating"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "imgClassic"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPrompt(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkTP"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkFav"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkLP"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkAdult"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lstDifficulty"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtHighScore"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lstRating"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtNotes"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtNotes 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3840
         Width           =   5055
      End
      Begin VB.ComboBox lstRating 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtHighScore 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox lstDifficulty 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   -73200
         TabIndex        =   22
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Fields to Update"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   12
         Top             =   480
         Width           =   5055
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Notes"
            Height          =   255
            Index           =   7
            Left            =   2640
            TabIndex        =   20
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Favourite Game"
            Height          =   255
            Index           =   6
            Left            =   2640
            TabIndex        =   19
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Adult Game"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   18
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Date Last Played"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   17
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Times Played"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Rating"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Difficulty"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "High Score"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.CheckBox chkAdult 
         Caption         =   "Ad&ult Game"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox chkLP 
         Caption         =   "Reset &Date &Last Played"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CheckBox chkFav 
         Caption         =   "Fa&vourite Game"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CheckBox chkTP 
         Caption         =   "Reset &Times Played"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Not&es:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   5055
      End
      Begin VB.Image imgClassic 
         Height          =   330
         Left            =   3360
         Picture         =   "frmPersonalProps.frx":0906
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgRating 
         Height          =   330
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   390
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Difficulty:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&High Score:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rating:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   5055
      End
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   6
      Left            =   2520
      Picture         =   "frmPersonalProps.frx":0C90
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   0
      Left            =   360
      Picture         =   "frmPersonalProps.frx":0D16
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   5
      Left            =   2160
      Picture         =   "frmPersonalProps.frx":0D92
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   4
      Left            =   1800
      Picture         =   "frmPersonalProps.frx":0E18
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   3
      Left            =   1440
      Picture         =   "frmPersonalProps.frx":0E9D
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   2
      Left            =   1080
      Picture         =   "frmPersonalProps.frx":0F21
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgRatingStore 
      Height          =   330
      Index           =   1
      Left            =   720
      Picture         =   "frmPersonalProps.frx":0FA6
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "frmPersonalProps"
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

Private Sub SetOKButton()

    If m_blnSeenOptions = False Then
        cmdOK.Enabled = False
    Else
        If (lstRating.ListIndex > -1) And _
           (lstDifficulty.ListIndex > -1) Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End If

End Sub

Private Sub chkUpdate_Click(Index As Integer)

    Dim lngIdx As Long
    Dim blnOK As Boolean

    If g_udtEditPersonalInfo.blnEditAll = True Then
        If Index = 0 Then
            If chkUpdate(0).Value = vbChecked Then
                txtHighScore.Enabled = True
                txtHighScore.BackColor = lstDummy.BackColor
            Else
                txtHighScore.Enabled = False
                txtHighScore.BackColor = BackColor
            End If
        End If
        If Index = 1 Then
            If chkUpdate(1).Value = vbChecked Then
                lstDifficulty.Enabled = True
                lstDifficulty.BackColor = lstDummy.BackColor
            Else
                lstDifficulty.Enabled = False
                lstDifficulty.BackColor = BackColor
            End If
        End If
        If Index = 2 Then
            If chkUpdate(2).Value = vbChecked Then
                lstRating.Enabled = True
                lstRating.BackColor = lstDummy.BackColor
                imgRating.Picture = imgRatingStore(lstRating.ListIndex).Picture
                imgRating.Visible = True
                imgClassic.Visible = (lstRating.ListIndex = 6)
            Else
                lstRating.Enabled = False
                lstRating.BackColor = BackColor
                imgRating.Visible = False
                imgClassic.Visible = False
            End If
        End If
        If Index = 3 Then
            If chkUpdate(3).Value = vbChecked Then
                chkTP.Enabled = True
            Else
                chkTP.Enabled = False
            End If
        End If
        If Index = 4 Then
            If chkUpdate(4).Value = vbChecked Then
                chkLP.Enabled = True
            Else
                chkLP.Enabled = False
            End If
        End If
        If Index = 5 Then
            If chkUpdate(5).Value = vbChecked Then
                chkAdult.Enabled = True
            Else
                chkAdult.Enabled = False
            End If
        End If
        If Index = 6 Then
            If chkUpdate(6).Value = vbChecked Then
                chkFav.Enabled = True
            Else
                chkFav.Enabled = False
            End If
        End If
        If Index = 7 Then
            If chkUpdate(7).Value = vbChecked Then
                txtNotes.Enabled = True
                txtNotes.BackColor = lstDummy.BackColor
            Else
                txtNotes.Enabled = False
                txtNotes.BackColor = BackColor
            End If
        End If

        'any selected?
        blnOK = False
        For lngIdx = 0 To 7
            If (chkUpdate(lngIdx).Value = vbChecked) And (chkUpdate(lngIdx).Enabled = True) Then
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

    For lngIdx = 0 To 7
        If chkUpdate(lngIdx).Enabled = True Then
            chkUpdate(lngIdx).Value = vbChecked
        End If
    Next lngIdx

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNone_Click()

    Dim lngIdx As Long

    For lngIdx = 0 To 7
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx

End Sub

Private Sub cmdOK_Click()

    'confirmation
    If g_udtEditPersonalInfo.blnEditAll Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSure", "Are you sure you want to update ALL GAMES in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If
    'save all info
    With g_udtEditPersonalInfo
        .blnOKClicked = True
        If .blnEditAll = True Then
            .blnUpdateAdult = (chkUpdate(5).Value = vbChecked)
            .blnUpdateDiff = (chkUpdate(1).Value = vbChecked)
            .blnUpdateFav = (chkUpdate(6).Value = vbChecked)
            .blnUpdateHighScore = (chkUpdate(0).Value = vbChecked)
            .blnUpdateLP = (chkUpdate(4).Value = vbChecked)
            .blnUpdateNotes = (chkUpdate(7).Value = vbChecked)
            .blnUpdateRating = (chkUpdate(2).Value = vbChecked)
            .blnUpdateTP = (chkUpdate(3).Value = vbChecked)
        End If
        .blnAdult = (chkAdult.Value = vbChecked)
        .blnClassic = (lstRating.ListIndex = 6)
        .blnFav = (chkFav.Value = vbChecked)
        .blnLP = (chkLP.Value = vbChecked)
        .blnTP = (chkTP.Value = vbChecked)
        .lngDiffID = lstDifficulty.ItemData(lstDifficulty.ListIndex)
        .lngRating = lstRating.ListIndex
        If .lngRating = 6 Then .lngRating = 5
        .strHighScore = Trim$(txtHighScore.Text)
        If .strHighScore = g_strNotEntered Then
            .strHighScore = ""
        End If
        .strNotes = txtNotes.Text
    End With

    Unload Me

End Sub
Private Sub Form_Activate()

    'each time the form is shown

    Dim lngIdx As Long

    m_blnLoading = False
    m_blnStartup = True

    'clear all checkmarks
    For lngIdx = 0 To 7
        chkUpdate(lngIdx).Value = vbChecked
        chkUpdate(lngIdx).Value = vbUnchecked
        If lngIdx = 5 Then
            chkUpdate(5).Enabled = (frmMain.mnuToolsHideAdult.Checked = False)
        End If
    Next lngIdx

    'are we changing all the games?
    If g_udtEditPersonalInfo.blnEditAll = False Then
        txtHighScore.BackColor = lstDummy.BackColor
        txtHighScore.Enabled = True
        lstDifficulty.BackColor = lstDummy.BackColor
        lstDifficulty.Enabled = True
        lstRating.BackColor = lstDummy.BackColor
        lstRating.Enabled = True
        chkTP.Enabled = True
        chkLP.Enabled = True
        chkFav.Enabled = True
        chkAdult.Enabled = (frmMain.mnuToolsHideAdult.Checked = False)
        txtNotes.BackColor = lstDummy.BackColor
        txtNotes.Enabled = True
        tabProps.TabEnabled(0) = False
        tabProps.TabEnabled(1) = True
        tabProps.Tab = 1
    Else
        tabProps.TabEnabled(0) = True
        tabProps.TabEnabled(1) = False
        tabProps.Tab = 0
    End If

    'init props
    txtHighScore.Text = ""
    If lstDifficulty.ListCount > 7 Then
        lstDifficulty.ListIndex = 7
    Else
        lstDifficulty.ListIndex = 0
    End If
    lstRating.ListIndex = 0
    chkTP.Value = vbUnchecked
    chkLP.Value = vbUnchecked
    chkAdult.Value = vbUnchecked
    chkFav.Value = vbUnchecked
    txtNotes.Text = ""

    'fill props
    txtHighScore.Text = g_udtEditPersonalInfo.strHighScore
    With lstDifficulty
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditPersonalInfo.lngDiffID Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    With lstRating
        If g_udtEditPersonalInfo.blnClassic = True Then
            .ListIndex = 6
        Else
            .ListIndex = g_udtEditPersonalInfo.lngRating
        End If
    End With
    chkTP.Value = Abs(g_udtEditPersonalInfo.blnTP)
    chkLP.Value = Abs(g_udtEditPersonalInfo.blnLP)
    chkAdult.Value = Abs(g_udtEditPersonalInfo.blnAdult)
    chkFav.Value = Abs(g_udtEditPersonalInfo.blnFav)
    txtNotes.Text = g_udtEditPersonalInfo.strNotes

    'set the focus
    If tabProps.Tab = 1 Then
        If g_udtEditPersonalInfoFocus = e_PersonalFocus_Adult Then
            If chkAdult.Enabled = True Then
                chkAdult.SetFocus
            Else
                txtHighScore.SetFocus
            End If
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_Difficulty Then
            lstDifficulty.SetFocus
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_Favourite Then
            chkFav.SetFocus
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_HighScore Then
            txtHighScore.SetFocus
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_LP Then
            chkLP.SetFocus
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_Memo Then
            txtNotes.SetFocus
        ElseIf g_udtEditPersonalInfoFocus = e_PersonalFocus_Rating Then
            lstRating.SetFocus
        Else 'times played
            chkTP.SetFocus
        End If
    Else
        chkUpdate(0).SetFocus
    End If

    'set ok button and finish off
    m_blnSeenOptions = Not g_udtEditPersonalInfo.blnEditAll
    Call SetOKButton
    g_udtEditPersonalInfo.blnOKClicked = False
    m_blnStartup = False
    Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 6375    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5775     'The required scalewidth for the form
    Dim lngIdx As Long
    Dim fldName As Field
    Dim fldID As Field

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    m_blnLoading = True

    'read in the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Personal Info Properties Window", "Caption", "Personal Info Properties") & "..."
        tabProps.TabCaption(0) = .GetINIMessage("Personal Info Properties Window", "IncludeTab", "&Include")
        tabProps.TabCaption(1) = .GetINIMessage("Personal Info Properties Window", "FieldsTab", "&Fields")
        cmdAll.Caption = .GetINIMessage("Personal Info Properties Window", "SelectAllButton", "Select &All")
        cmdNone.Caption = .GetINIMessage("Personal Info Properties Window", "SelectNoneButton", "Select &None")
        fraUpdate.Caption = .GetINIMessage("Personal Info Properties Window", "FieldsToUpdate", "Fields to Update")
        chkUpdate(0).Caption = GetINIMessage("Personal Info Properties Window", "UpdateHighScoreBox", "High Score")
        chkUpdate(1).Caption = GetINIMessage("Personal Info Properties Window", "UpdateDifficultyBox", "Difficulty")
        chkUpdate(2).Caption = GetINIMessage("Personal Info Properties Window", "UpdateRatingBox", "Rating")
        chkUpdate(3).Caption = GetINIMessage("Personal Info Properties Window", "UpdateTimesPlayedBox", "Times Played")
        chkUpdate(4).Caption = GetINIMessage("Personal Info Properties Window", "UpdateDateLastPlayedBox", "Date Last Played")
        chkUpdate(5).Caption = GetINIMessage("Personal Info Properties Window", "UpdateAdultGameBox", "Adult Game")
        chkUpdate(6).Caption = GetINIMessage("Personal Info Properties Window", "UpdateFavouriteGameBox", "Favourite Game")
        chkUpdate(7).Caption = GetINIMessage("Personal Info Properties Window", "UpdateNotesBox", "Notes")
        lblPrompt(2).Caption = GetINIMessage("Personal Info Properties Window", "HighScorePrompt", "&High Score") & ":"
        lblPrompt(9).Caption = GetINIMessage("Personal Info Properties Window", "DifficultyPrompt", "&Difficulty") & ":"
        lblPrompt(3).Caption = GetINIMessage("Personal Info Properties Window", "RatingPrompt", "&Rating") & ":"
        lblPrompt(0).Caption = GetINIMessage("Personal Info Properties Window", "NotesPrompt", "Not&es") & ":"
        chkTP.Caption = GetINIMessage("Personal Info Properties Window", "TimesPlayedPrompt", "Reset &Times Played")
        chkLP.Caption = GetINIMessage("Personal Info Properties Window", "DateLastPlayedPrompt", "Reset Date &Last Played")
        chkAdult.Caption = GetINIMessage("Personal Info Properties Window", "AdultPrompt", "Ad&ult Game")
        chkFav.Caption = GetINIMessage("Personal Info Properties Window", "FavouritePrompt", "Fa&vourite Game")
    End With
    
    'load custom game rating images?
    Call LoadGBImageFromFile(imgRatingStore(5), g_strGameBasePathWBS & "Gfx\VeryGoodRating_13x11.gif")
    Call LoadGBImageFromFile(imgRatingStore(4), g_strGameBasePathWBS & "Gfx\QuiteGoodRating_13x11.gif")
    Call LoadGBImageFromFile(imgRatingStore(3), g_strGameBasePathWBS & "Gfx\AverageRating_13x11.gif")
    Call LoadGBImageFromFile(imgRatingStore(2), g_strGameBasePathWBS & "Gfx\PoorRating_13x11.gif")
    Call LoadGBImageFromFile(imgRatingStore(1), g_strGameBasePathWBS & "Gfx\TerribleRating_13x11.gif")
    Call LoadGBImageFromFile(imgRatingStore(0), g_strGameBasePathWBS & "Gfx\UnknownRating_13x11.gif")
    imgRatingStore(6).Picture = imgRatingStore(5).Picture
    Call LoadGBImageFromFile(imgClassic, g_strGameBasePathWBS & "Gfx\ClassicGame_16x11.gif")
    
    'read in the difficulty levels
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT DI_Id as ID, Difficulty as Name FROM Difficulty ORDER BY DI_Id", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            Set fldID = !ID
            Set fldName = !Name
            .MoveFirst
            With lstDifficulty
                Do Until g_rsGames.EOF
                    .AddItem fldName
                    .ItemData(.NewIndex) = fldID
                    g_rsGames.MoveNext
                Loop
            End With
            Set fldID = Nothing
            Set fldName = Nothing
        End If
        .Close
    End With
    Set g_rsGames = Nothing
    lstDifficulty.Visible = True

    'read in ratings
    With lstRating
        .AddItem ga_strRatingLookupTable(0)
        .AddItem ga_strRatingLookupTable(1)
        .AddItem ga_strRatingLookupTable(2)
        .AddItem ga_strRatingLookupTable(3)
        .AddItem ga_strRatingLookupTable(4)
        .AddItem ga_strRatingLookupTable(5)
        .AddItem ga_strRatingLookupTable(6)
        .Visible = True
    End With

End Sub

Private Sub lstDifficulty_Click()
    Call SetOKButton
End Sub
Private Sub lstRating_Click()

    With lstRating
        If .ListIndex > -1 Then
            imgRating.Picture = imgRatingStore(.ListIndex).Picture
            imgClassic.Visible = (.ListIndex = 6)
        Else
            imgRating.Picture = imgRatingStore(0).Picture
            imgClassic.Visible = False
        End If
    End With
    Call SetOKButton

End Sub
Private Sub tabProps_Click(PreviousTab As Integer)
    
    If tabProps.Tab <> PreviousTab Then
        If tabProps.Tab = 1 Then
            If Not m_blnLoading Then
                If txtHighScore.Enabled = True Then
                    txtHighScore.SetFocus
                ElseIf lstDifficulty.Enabled = True Then
                    lstDifficulty.SetFocus
                ElseIf lstRating.Enabled = True Then
                    lstRating.SetFocus
                ElseIf chkTP.Enabled = True Then
                    chkTP.SetFocus
                ElseIf chkLP.Enabled = True Then
                    chkLP.SetFocus
                ElseIf chkAdult.Enabled = True Then
                    chkAdult.SetFocus
                ElseIf chkFav.Enabled = True Then
                    chkFav.SetFocus
                ElseIf txtNotes.Enabled = True Then
                    txtNotes.SetFocus
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

Private Sub txtHighScore_GotFocus()

    With txtHighScore
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtNotes_GotFocus()

    cmdOK.Default = False
    With txtNotes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtNotes_LostFocus()

    cmdOK.Default = True

End Sub
