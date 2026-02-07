VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVersionProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Version Info Properties..."
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmVersionProps.frx":0000
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
      TabIndex        =   39
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   38
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ListBox lstDummy 
      Enabled         =   0   'False
      Height          =   255
      ItemData        =   "frmVersionProps.frx":08CA
      Left            =   120
      List            =   "frmVersionProps.frx":08CC
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabProps 
      Height          =   5535
      Left            =   120
      TabIndex        =   41
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
      TabCaption(0)   =   "&Include"
      TabPicture(0)   =   "frmVersionProps.frx":08CE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraUpdate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNone"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Fields"
      TabPicture(1)   =   "frmVersionProps.frx":08EA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPrompt(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPrompt(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPrompt(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblPrompt(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPrompt(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblCracker"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblPrompt(15)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblPrompt(14)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtComment"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdCracker"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkDocs"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkLoadingScreen"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkTDE"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkHSS"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lstLength"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lstNTSC"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lstTrainers"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtLength"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtWebLinkURL"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtWebLinkName"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "chkPlayable"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkOriginal"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "chkTitleScreen"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      Begin VB.CheckBox chkTitleScreen 
         Caption         =   "&Title Screen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkOriginal 
         Caption         =   "&Original"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox chkPlayable 
         Caption         =   "&Playable"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtWebLinkName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3600
         Width           =   5055
      End
      Begin VB.TextBox txtWebLinkURL 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   4320
         Width           =   5055
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox lstTrainers 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox lstNTSC 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox lstLength 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1530
      End
      Begin VB.CheckBox chkHSS 
         Caption         =   "&High Score Saver"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkTDE 
         Caption         =   "True Drive &Emulation"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox chkLoadingScreen 
         Caption         =   "&Loading Screen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CheckBox chkDocs 
         Caption         =   "Included &Docs"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Fields to Update"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   5055
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Download"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Original"
            Height          =   255
            Index           =   11
            Left            =   2640
            TabIndex        =   30
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Playable"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "&Title Screen"
            Height          =   255
            Index           =   9
            Left            =   2640
            TabIndex        =   32
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Cracker"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "No. Trainers"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "High Score Saver"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   34
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Game Length"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "True Drive Emulation"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   35
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Loading Screen"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   31
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "PAL/NTSC"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   27
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Version Comment"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   2520
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Included Docs"
            Height          =   255
            Index           =   7
            Left            =   2640
            TabIndex        =   33
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   -73200
         TabIndex        =   37
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdCracker 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtComment 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   255
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   5010
         Width           =   5055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Download &Name:"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   3855
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Download &URL:"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label lblCracker 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&PAL/NTSC:"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Game Length:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "No. &Trainers:"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "C&racker:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "&Version Comment:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   20
         Top             =   4800
         Width           =   5055
      End
   End
   Begin VB.Label lblDummy 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   1200
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmVersionProps"
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

Private m_blnDocsGrayed As Boolean
Private m_blnTDEGrayed As Boolean
Private m_blnHSSGrayed As Boolean
Private m_blnLoadingScreenGrayed As Boolean
Private m_blnTitleScreenGrayed As Boolean
Private m_blnOriginalGrayed As Boolean
Private m_blnPlayableGrayed As Boolean

Private m_blnByCode As Boolean
Private m_blnSeenOptions As Boolean
Private m_blnLoading As Boolean
Private m_blnStartup As Boolean
Private m_lngCrackerID As Long
Private m_strCracker As String
Private m_blnFirstTime As Boolean
Private Const mc_strLENGTHCHARS As String = "0123456789?"
Private Sub SetOKButton()

    If m_blnSeenOptions = False Then
        cmdOK.Enabled = False
    Else
        If (lstTrainers.ListIndex > -1) And _
           (lstLength.ListIndex > -1) And _
           (lstNTSC.ListIndex > -1) And _
           (Len(txtLength) > 0) Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End If

End Sub

Private Sub chkDocs_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnDocsGrayed = True Then
            chkDocs.Value = vbUnchecked
            m_blnDocsGrayed = False
        Else
            If chkDocs.Value = vbUnchecked Then
                m_blnDocsGrayed = True
                chkDocs.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkHSS_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnHSSGrayed = True Then
            chkHSS.Value = vbUnchecked
            m_blnHSSGrayed = False
        Else
            If chkHSS.Value = vbUnchecked Then
                m_blnHSSGrayed = True
                chkHSS.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkLoadingScreen_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnLoadingScreenGrayed = True Then
            chkLoadingScreen.Value = vbUnchecked
            m_blnLoadingScreenGrayed = False
        Else
            If chkLoadingScreen.Value = vbUnchecked Then
                m_blnLoadingScreenGrayed = True
                chkLoadingScreen.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkOriginal_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnOriginalGrayed = True Then
            chkOriginal.Value = vbUnchecked
            m_blnOriginalGrayed = False
        Else
            If chkOriginal.Value = vbUnchecked Then
                m_blnOriginalGrayed = True
                chkOriginal.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkPlayable_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnPlayableGrayed = True Then
            chkPlayable.Value = vbUnchecked
            m_blnPlayableGrayed = False
        Else
            If chkPlayable.Value = vbUnchecked Then
                m_blnPlayableGrayed = True
                chkPlayable.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkTDE_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnTDEGrayed = True Then
            chkTDE.Value = vbUnchecked
            m_blnTDEGrayed = False
        Else
            If chkTDE.Value = vbUnchecked Then
                m_blnTDEGrayed = True
                chkTDE.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkTitleScreen_Click()

    If Not m_blnByCode Then

        m_blnByCode = True

        If m_blnTitleScreenGrayed = True Then
            chkTitleScreen.Value = vbUnchecked
            m_blnTitleScreenGrayed = False
        Else
            If chkTitleScreen.Value = vbUnchecked Then
                m_blnTitleScreenGrayed = True
                chkTitleScreen.Value = vbGrayed
            End If
        End If

        m_blnByCode = False

    End If

End Sub

Private Sub chkUpdate_Click(Index As Integer)

    Dim lngIdx As Long
    Dim blnOK As Boolean

    If g_udtEditVersionInfo.blnEditAll = True Then
        If Index = 0 Then
            If chkUpdate(0).Value = vbChecked Then
                lblCracker.BackColor = lblDummy.BackColor
                lblCracker.Enabled = True
                cmdCracker.Enabled = True
            Else
                lblCracker.BackColor = BackColor
                cmdCracker.Enabled = False
                lblCracker.Enabled = False
            End If
        End If
        If Index = 1 Then
            If chkUpdate(1).Value = vbChecked Then
                lstTrainers.Enabled = True
                lstTrainers.BackColor = lstDummy.BackColor
            Else
                lstTrainers.Enabled = False
                lstTrainers.BackColor = BackColor
            End If
        End If
        If Index = 2 Then
            If chkUpdate(2).Value = vbChecked Then
                chkHSS.Enabled = True
            Else
                chkHSS.Enabled = False
            End If
        End If
        If Index = 3 Then
            If chkUpdate(3).Value = vbChecked Then
                lstLength.Enabled = True
                lstLength.BackColor = lstDummy.BackColor
                txtLength.Enabled = True
                txtLength.BackColor = lstDummy.BackColor
            Else
                lstLength.Enabled = False
                lstLength.BackColor = BackColor
                txtLength.Enabled = False
                txtLength.BackColor = BackColor
            End If
        End If
        If Index = 4 Then
            If chkUpdate(4).Value = vbChecked Then
                chkTDE.Enabled = True
            Else
                chkTDE.Enabled = False
            End If
        End If
        If Index = 5 Then
            If chkUpdate(5).Value = vbChecked Then
                chkLoadingScreen.Enabled = True
            Else
                chkLoadingScreen.Enabled = False
            End If
        End If
        If Index = 6 Then
            If chkUpdate(6).Value = vbChecked Then
                lstNTSC.Enabled = True
                lstNTSC.BackColor = lstDummy.BackColor
            Else
                lstNTSC.Enabled = False
                lstNTSC.BackColor = BackColor
            End If
        End If
        If Index = 7 Then
            If chkUpdate(7).Value = vbChecked Then
                chkDocs.Enabled = True
            Else
                chkDocs.Enabled = False
            End If
        End If
        If Index = 8 Then
            If chkUpdate(8).Value = vbChecked Then
                txtComment.Enabled = True
                txtComment.BackColor = lstDummy.BackColor
            Else
                txtComment.Enabled = False
                txtComment.BackColor = BackColor
            End If
        End If
        If Index = 9 Then
            If chkUpdate(9).Value = vbChecked Then
                chkTitleScreen.Enabled = True
            Else
                chkTitleScreen.Enabled = False
            End If
        End If
        If Index = 10 Then
            If chkUpdate(10).Value = vbChecked Then
                chkPlayable.Enabled = True
            Else
                chkPlayable.Enabled = False
            End If
        End If
        If Index = 11 Then
            If chkUpdate(11).Value = vbChecked Then
                chkOriginal.Enabled = True
            Else
                chkOriginal.Enabled = False
            End If
        End If
        If Index = 12 Then
            If chkUpdate(12).Value = vbChecked Then
                txtWebLinkName.Enabled = True
                txtWebLinkName.BackColor = lstDummy.BackColor
                txtWebLinkURL.Enabled = True
                txtWebLinkURL.BackColor = lstDummy.BackColor
            Else
                txtWebLinkName.Enabled = False
                txtWebLinkName.BackColor = BackColor
                txtWebLinkURL.Enabled = False
                txtWebLinkURL.BackColor = BackColor
            End If
        End If

        'any selected?
        blnOK = False
        For lngIdx = 0 To 12
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
    
    For lngIdx = 0 To 12
        chkUpdate(lngIdx).Value = vbChecked
    Next lngIdx

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCracker_Click()

    With g_udtRecordSelector
        .blnOKClicked = False
        .eRecordSelectorType = e_RecordSelector_Cracker
        .lngID = m_lngCrackerID
        .strSQL = "SELECT Cracker AS Name, CR_Id as ID FROM Crackers ORDER BY Cracker"
        frmRecordSelector.Show vbModal, Me
        If .blnOKClicked = True Then
            g_udtEditVersionInfo.blnTableEdited = True
            m_lngCrackerID = .lngID
            m_strCracker = .strReturnedName
            lblCracker.Caption = m_strCracker
        End If
    End With
    cmdCracker.SetFocus

End Sub

Private Sub cmdNone_Click()

    Dim lngIdx As Long
    
    For lngIdx = 0 To 12
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx
    
End Sub

Private Sub cmdOK_Click()

    'confirmation
    If g_udtEditVersionInfo.blnEditAll Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSure", "Are you sure you want to update ALL GAMES in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If
    'save all info
    With g_udtEditVersionInfo
        .blnOKClicked = True
        If .blnEditAll = True Then
            .blnUpdateCracker = (chkUpdate(0).Value = vbChecked)
            .blnUpdateDocs = (chkUpdate(7).Value = vbChecked)
            .blnUpdateHSS = (chkUpdate(2).Value = vbChecked)
            .blnUpdateLength = (chkUpdate(3).Value = vbChecked)
            .blnUpdateLoadingScreen = (chkUpdate(5).Value = vbChecked)
            .blnUpdateNTSC = (chkUpdate(6).Value = vbChecked)
            .blnUpdateTDE = (chkUpdate(4).Value = vbChecked)
            .blnUpdateTrainers = (chkUpdate(1).Value = vbChecked)
            .blnUpdateVComment = (chkUpdate(8).Value = vbChecked)
            .blnUpdateOriginal = (chkUpdate(11).Value = vbChecked)
            .blnUpdatePlayable = (chkUpdate(10).Value = vbChecked)
            .blnUpdateTitleScreen = (chkUpdate(9).Value = vbChecked)
            .blnUpdateWebLink = (chkUpdate(12).Value = vbChecked)
        End If
        .lngDocs = chkDocs.Value
        .lngHSS = chkHSS.Value
        .lngLoadingScreen = chkLoadingScreen.Value
        .lngTDE = chkTDE.Value
        .lngPlayable = chkPlayable.Value
        .lngTitleScreen = chkTitleScreen.Value
        .lngOriginal = chkOriginal.Value
        .lngCrackerID = m_lngCrackerID
        .lngGameLengthType = lstLength.ListIndex
        .lngGameLength = Val(txtLength)
        .lngNTSC = lstNTSC.ItemData(lstNTSC.ListIndex)
        .lngTrainers = lstTrainers.ItemData(lstTrainers.ListIndex)
        .strVComment = txtComment.Text
        .strWebLinkName = txtWebLinkName.Text
        .strWebLinkURL = txtWebLinkURL.Text
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
    For lngIdx = 0 To 12
        chkUpdate(lngIdx).Value = vbChecked
        chkUpdate(lngIdx).Value = vbUnchecked
    Next lngIdx

    'are we changing all the games?
    If g_udtEditVersionInfo.blnEditAll = False Then
        lblCracker.BackColor = lblDummy.BackColor
        lblCracker.Enabled = True
        cmdCracker.Enabled = True
        lstNTSC.BackColor = lstDummy.BackColor
        lstNTSC.Enabled = True
        lstLength.BackColor = lstDummy.BackColor
        lstLength.Enabled = True
        txtLength.BackColor = lstDummy.BackColor
        txtLength.Enabled = True
        lstTrainers.BackColor = lstDummy.BackColor
        lstTrainers.Enabled = True
        chkHSS.Enabled = True
        chkLoadingScreen.Enabled = True
        chkDocs.Enabled = True
        chkTDE.Enabled = True
        chkPlayable.Enabled = True
        chkOriginal.Enabled = True
        chkTitleScreen.Enabled = True
        txtComment.BackColor = lstDummy.BackColor
        txtComment.Enabled = True
        txtWebLinkName.BackColor = lstDummy.BackColor
        txtWebLinkName.Enabled = True
        txtWebLinkURL.BackColor = lstDummy.BackColor
        txtWebLinkURL.Enabled = True
        tabProps.TabEnabled(0) = False
        tabProps.TabEnabled(1) = True
        tabProps.Tab = 1
    Else
        tabProps.TabEnabled(0) = True
        tabProps.TabEnabled(1) = False
        tabProps.Tab = 0
    End If

    'init props
    lstNTSC.ListIndex = 2
    lstLength.ListIndex = 0
    txtLength = "?"
    lstTrainers.ListIndex = 0
    chkHSS.Value = vbGrayed
    chkLoadingScreen.Value = vbGrayed
    chkDocs.Value = vbGrayed
    chkTDE.Value = vbGrayed
    chkTitleScreen = vbGrayed
    chkOriginal = vbGrayed
    chkPlayable = vbGrayed
    txtComment.Text = ""
    txtWebLinkName.Text = ""
    txtWebLinkURL.Text = ""

    'fill props
    lblCracker.Caption = m_strCracker
    With lstTrainers
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditVersionInfo.lngTrainers Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    With lstNTSC
        For lngIdx = 0 To .ListCount - 1
            If .ItemData(lngIdx) = g_udtEditVersionInfo.lngNTSC Then
                .ListIndex = lngIdx
                Exit For
            End If
        Next lngIdx
    End With
    lstLength.ListIndex = g_udtEditVersionInfo.lngGameLengthType
    txtLength = g_udtEditVersionInfo.lngGameLength
    chkHSS.Value = g_udtEditVersionInfo.lngHSS
    chkLoadingScreen.Value = g_udtEditVersionInfo.lngLoadingScreen
    chkDocs.Value = g_udtEditVersionInfo.lngDocs
    chkTDE.Value = g_udtEditVersionInfo.lngTDE
    chkOriginal.Value = g_udtEditVersionInfo.lngOriginal
    chkPlayable.Value = g_udtEditVersionInfo.lngPlayable
    chkTitleScreen.Value = g_udtEditVersionInfo.lngTitleScreen
    txtComment.Text = g_udtEditVersionInfo.strVComment
    txtWebLinkName.Text = g_udtEditVersionInfo.strWebLinkName
    txtWebLinkURL.Text = g_udtEditVersionInfo.strWebLinkURL
    
    m_blnDocsGrayed = (chkDocs.Value = vbGrayed)
    m_blnTDEGrayed = (chkTDE.Value = vbGrayed)
    m_blnHSSGrayed = (chkHSS.Value = vbGrayed)
    m_blnLoadingScreenGrayed = (chkLoadingScreen.Value = vbGrayed)
    m_blnTitleScreenGrayed = (chkTitleScreen.Value = vbGrayed)
    m_blnOriginalGrayed = (chkOriginal.Value = vbGrayed)
    m_blnPlayableGrayed = (chkPlayable.Value = vbGrayed)

    'set the focus
    If tabProps.Tab = 1 Then
        If g_udtEditVersionInfoFocus = e_VersionFocus_Cracker Then
            cmdCracker.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_Docs Then
            chkDocs.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_HSS Then
            chkHSS.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_Length Then
            txtLength.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_LoadingScreen Then
            chkLoadingScreen.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_NTSC Then
            lstNTSC.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_TDE Then
            chkTDE.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_Trainers Then
            lstTrainers.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_Original Then
            chkOriginal.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_Playable Then
            chkPlayable.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_TitleScreen Then
            chkTitleScreen.SetFocus
        ElseIf g_udtEditVersionInfoFocus = e_VersionFocus_WebLink Then
            txtWebLinkName.SetFocus
        Else 'comment
            txtComment.SetFocus
        End If
    Else
        chkUpdate(0).SetFocus
    End If

    'set ok button and finish off
    m_blnSeenOptions = Not g_udtEditVersionInfo.blnEditAll
    Call SetOKButton
    g_udtEditVersionInfo.blnOKClicked = False
    m_blnStartup = False
    g_udtEditVersionInfo.blnTableEdited = False
    Screen.MousePointer = vbDefault
    m_blnByCode = False

End Sub
Private Sub Form_Load()
    
    Const DesignScaleHeight As Single = 6375    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5775     'The required scalewidth for the form
    Dim lngIdx As Long

    m_blnByCode = True

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    m_blnLoading = True
    m_lngCrackerID = g_udtEditVersionInfo.lngCrackerID

    'read in the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Version Info Properties Window", "Caption", "Version Info Properties") & "..."
        tabProps.TabCaption(0) = .GetINIMessage("Version Info Properties Window", "IncludeTab", "&Include")
        tabProps.TabCaption(1) = .GetINIMessage("Version Info Properties Window", "FieldsTab", "&Fields")
        cmdAll.Caption = .GetINIMessage("Version Info Properties Window", "SelectAllButton", "Select &All")
        cmdNone.Caption = .GetINIMessage("Version Info Properties Window", "SelectNoneButton", "Select &None")
        fraUpdate.Caption = .GetINIMessage("Version Info Properties Window", "FieldsToUpdate", "Fields to Update")
        chkUpdate(0).Caption = GetINIMessage("Version Info Properties Window", "UpdateCrackerBox", "Cracker")
        chkUpdate(1).Caption = GetINIMessage("Version Info Properties Window", "UpdateTrainersBox", "No. Trainers")
        chkUpdate(2).Caption = GetINIMessage("Version Info Properties Window", "UpdateHighScoreSaverBox", "High Score Saver")
        chkUpdate(3).Caption = GetINIMessage("Version Info Properties Window", "UpdateGameLengthBox", "Game Length")
        chkUpdate(6).Caption = GetINIMessage("Version Info Properties Window", "UpdatePalNTSCBox", "PAL/NTSC")
        chkUpdate(4).Caption = GetINIMessage("Version Info Properties Window", "UpdateTrueDriveEmuBox", "True Drive Emulation")
        chkUpdate(5).Caption = GetINIMessage("Version Info Properties Window", "UpdateLoadingScreenBox", "Loading Screen")
        chkUpdate(7).Caption = GetINIMessage("Version Info Properties Window", "UpdateIncludedDocsBox", "Included Docs")
        chkUpdate(8).Caption = GetINIMessage("Version Info Properties Window", "UpdateVersionCommentBox", "Version Comment")
        chkUpdate(9).Caption = GetINIMessage("Version Info Properties Window", "UpdateVersionTitleScreenBox", "Title Screen")
        chkUpdate(10).Caption = GetINIMessage("Version Info Properties Window", "UpdateVersionPlayableBox", "Playable")
        chkUpdate(11).Caption = GetINIMessage("Version Info Properties Window", "UpdateVersionOriginalBox", "Original")
        chkUpdate(12).Caption = GetINIMessage("Version Info Properties Window", "UpdateVersionWebLinkBox", "Download")
        lblPrompt(2).Caption = GetINIMessage("Version Info Properties Window", "CrackerPrompt", "C&racker") & ":"
        lblPrompt(0).Caption = GetINIMessage("Version Info Properties Window", "TrainersPrompt", "No. &Trainers") & ":"
        lblPrompt(7).Caption = GetINIMessage("Version Info Properties Window", "PalNTSCPrompt", "&PAL/NTSC") & ":"
        lblPrompt(3).Caption = GetINIMessage("Version Info Properties Window", "GameLengthPrompt", "&Game Length ('0' or '?' = Unknown)") & ":"
        chkHSS.Caption = GetINIMessage("Version Info Properties Window", "HighScoreSaverPrompt", "&High Score Saver")
        chkLoadingScreen.Caption = GetINIMessage("Version Info Properties Window", "LoadingScreenPrompt", "&Loading Screen")
        chkDocs.Caption = GetINIMessage("Version Info Properties Window", "IncludedDocsPrompt", "Included &Docs")
        chkTDE.Caption = GetINIMessage("Version Info Properties Window", "TrueDriveEmuPrompt", "True Drive &Emulation")
        chkOriginal.Caption = GetINIMessage("Version Info Properties Window", "OriginalPrompt", "&Original")
        chkPlayable.Caption = GetINIMessage("Version Info Properties Window", "PlayablePrompt", "&Playable")
        chkTitleScreen.Caption = GetINIMessage("Version Info Properties Window", "TitleScreenPrompt", "&Title Screen")
        lblPrompt(14).Caption = GetINIMessage("Version Info Properties Window", "WebLinkNamePrompt", "Download &Name") & ":"
        lblPrompt(15).Caption = GetINIMessage("Version Info Properties Window", "WebLinkURLPrompt", "Download &URL") & ":"
        lblPrompt(9).Caption = GetINIMessage("Version Info Properties Window", "VersionCommentPrompt", "&Version Comment") & ":"
    End With

    'read in the pal/ntsc selections
    With lstNTSC
        .AddItem ga_strPalNTSCLookupTable(0)
        .AddItem ga_strPalNTSCLookupTable(1)
        .AddItem ga_strPalNTSCLookupTable(2)
        .AddItem ga_strPalNTSCLookupTable(3)
        .ItemData(0) = 0
        .ItemData(1) = 1
        .ItemData(2) = 2
        .ItemData(3) = 3
        .Visible = True
    End With
    'read in the game length selections
    With lstLength
        For lngIdx = 0 To 19
            .AddItem ga_strLengthTypeLookupTable(lngIdx)
        Next lngIdx
        .Visible = True
    End With
    'read in the trainer selections
    With lstTrainers
        .AddItem "(" & g_strUnknown & ")"
        .ItemData(.NewIndex) = gc_lngTRAINERSUNKNOWN
        For lngIdx = 0 To 5000
            .AddItem lngIdx
            .ItemData(.NewIndex) = lngIdx
        Next lngIdx
        .Visible = True
    End With

    'read in the cracker
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT CR_Id as ID, Cracker as Name FROM Crackers WHERE CR_Id = " & m_lngCrackerID, dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            m_strCracker = !Name
        End If
        .Close
    End With
    Set g_rsGames = Nothing

    m_blnFirstTime = True
    
End Sub


Private Sub lstLength_Click()
    Call SetOKButton
End Sub

Private Sub lstNTSC_Click()
    Call SetOKButton
End Sub

Private Sub lstTrainers_Click()
    Call SetOKButton
End Sub

Private Sub tabProps_Click(PreviousTab As Integer)
    
    If tabProps.Tab <> PreviousTab Then
        If tabProps.Tab = 1 Then
            If Not m_blnLoading Then
                If cmdCracker.Enabled = True Then
                    cmdCracker.SetFocus
                ElseIf lstTrainers.Enabled = True Then
                    lstTrainers.SetFocus
                ElseIf lstNTSC.Enabled = True Then
                    lstNTSC.SetFocus
                ElseIf txtLength.Enabled = True Then
                    txtLength.SetFocus
                ElseIf lstLength.Enabled = True Then
                    lstLength.SetFocus
                ElseIf chkHSS.Enabled = True Then
                    chkHSS.SetFocus
                ElseIf chkLoadingScreen.Enabled = True Then
                    chkLoadingScreen.SetFocus
                ElseIf chkDocs.Enabled = True Then
                    chkDocs.SetFocus
                ElseIf chkTDE.Enabled = True Then
                    chkTDE.SetFocus
                ElseIf chkOriginal.Enabled = True Then
                    chkOriginal.SetFocus
                ElseIf chkPlayable.Enabled = True Then
                    chkPlayable.SetFocus
                ElseIf chkTitleScreen.Enabled = True Then
                    chkTitleScreen.SetFocus
                ElseIf txtWebLinkName.Enabled = True Then
                    txtWebLinkName.SetFocus
                ElseIf txtWebLinkURL.Enabled = True Then
                    txtWebLinkURL.SetFocus
                ElseIf txtComment.Enabled = True Then
                    txtComment.SetFocus
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

Private Sub txtComment_GotFocus()

    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtLength_Change()
    Call SetOKButton
End Sub

Private Sub txtLength_GotFocus()

    With txtLength
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    Else
        If InStr(mc_strLENGTHCHARS, Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtWebLinkName_GotFocus()

    With txtWebLinkName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtWebLinkURL_GotFocus()

    With txtWebLinkURL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
