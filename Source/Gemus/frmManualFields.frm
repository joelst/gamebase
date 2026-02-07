VERSION 5.00
Begin VB.Form frmManualFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set GEMUS Script Fields"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmManualFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBack 
      BackColor       =   &H80000005&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   5835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   240
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3360
         Width           =   5415
      End
      Begin VB.ComboBox lstTDE 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3375
      End
      Begin VB.ComboBox lstPalNTSC 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox lstPlayers 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   125
         TabIndex        =   10
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtVComment 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   125
         TabIndex        =   12
         Top             =   2640
         Width           =   3375
      End
      Begin VB.ComboBox lstControl 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Keys && Values:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   1680
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   19
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&VersionComment"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&GameComment"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&TrueDriveEmu"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&PalNTSC"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&NumPlayers"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cont&rol"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmManualFields"
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
'     GEMUS - Gamebase EMUlator Script
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

Private m_strMessFile As String
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    With g_udtManualFields
        .blnOKClicked = True
        .blnTDE = (lstTDE.ListIndex = 1)
        .lngControl = lstControl.ListIndex
        .lngPalNTSC = lstPalNTSC.ListIndex
        .lngPlayers = lstPlayers.ListIndex + 1
        .strComment = txtComment.Text
        .strVComment = txtVComment.Text
        .strNotes = txtNotes.Text
    End With
    
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 6615    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6135     'The required scalewidth for the form
    Dim lngIdx As Long                          'loop index

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'language ini file text
    m_strMessFile = App.Path
    If Right$(m_strMessFile, 1) <> "\" Then m_strMessFile = m_strMessFile & "\"
    m_strMessFile = m_strMessFile & "GemusText.ini"
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", m_strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", m_strMessFile)
    If g_udtManualFields.blnDefaults = True Then
        Caption = GetINIMessageFromFile("GEMUS Script Fields Window", "DefaultCaption", "Set Default GEMUS Script Fields", m_strMessFile)
    Else
        Caption = GetINIMessageFromFile("GEMUS Script Fields Window", "Caption", "Set GEMUS Script Fields", m_strMessFile)
    End If

    'fill listboxes
    With lstControl
        .AddItem "JoyPort2"
        .AddItem "JoyPort1"
        .AddItem "Keyboard"
        .AddItem "PaddlePort2"
        .AddItem "PaddlePort1"
        .AddItem "Mouse"
        .AddItem "LightPen"
        .AddItem "KoalaPad"
        .AddItem "LightGun"
    End With
    With lstPlayers
        For lngIdx = 1 To 100
            .AddItem lngIdx
        Next lngIdx
    End With
    With lstPalNTSC
        .AddItem "PAL"
        .AddItem "BOTH"
        .AddItem "NTSC"
    End With
    With lstTDE
        .AddItem "NO"
        .AddItem "YES"
    End With

    'set values
    With g_udtManualFields
        If .lngControl > -1 And .lngControl < 9 Then
            lstControl.ListIndex = .lngControl
        Else
            lstControl.ListIndex = 0
        End If
        If .lngPlayers > 0 And .lngPlayers < 101 Then
            lstPlayers.ListIndex = .lngPlayers - 1
        Else
            lstPlayers.ListIndex = 0
        End If
        If .lngPalNTSC > -1 And .lngPalNTSC < 3 Then
            lstPalNTSC.ListIndex = .lngPalNTSC
        Else
            lstPalNTSC.ListIndex = 0
        End If
        lstTDE.ListIndex = Abs(.blnTDE)
        If Len(.strComment) > txtComment.MaxLength Then
            txtComment.Text = Left$(.strComment, txtComment.MaxLength)
        Else
            txtComment.Text = .strComment
        End If
        If Len(.strVComment) > txtVComment.MaxLength Then
            txtVComment.Text = Left$(.strVComment, txtVComment.MaxLength)
        Else
            txtVComment.Text = .strVComment
        End If
        If Len(.strNotes) > txtNotes.MaxLength Then
            txtNotes.Text = Left$(.strNotes, txtNotes.MaxLength)
        Else
            txtNotes.Text = .strNotes
        End If
    End With

    'reset flag
    g_udtManualFields.blnOKClicked = False

End Sub

Private Sub txtComment_GotFocus()

    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtNotes_GotFocus()
    cmdOK.Default = False
End Sub


Private Sub txtNotes_LostFocus()
    cmdOK.Default = True
End Sub


Private Sub txtVComment_GotFocus()

    With txtVComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


