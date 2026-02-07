VERSION 5.00
Begin VB.Form frmGameLinkSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Link Selector"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmGameLinkSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      Caption         =   "Prompt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.OptionButton optField 
         Caption         =   "Clone &Of"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   9495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&Related Game"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   9495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&Sequel"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   9495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&Prequel"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   9495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmGameLinkSelector"
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


Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    If optField(0).Value = True Then
        g_strGameLinkField = "Prequel"
    ElseIf optField(1).Value = True Then
        g_strGameLinkField = "Sequel"
    ElseIf optField(2).Value = True Then
        g_strGameLinkField = "Related"
    Else
        g_strGameLinkField = "CloneOf"
    End If
    g_strDefaultGameLinkField = g_strGameLinkField
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3135    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 10095     'The required scalewidth for the form
    Dim strIs As String
    Dim strTo As String
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'read ini file text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        If g_blnGameLinkSetFlag = True Then
            Caption = .GetINIMessage("Game Link Selector", "SetCaption", "Set Game Link")
            fraFields.Caption = g_strMarkedGame & " -> " & g_udtGame.strName
        Else
            Caption = .GetINIMessage("Game Link Selector", "ClearCaption", "Clear Game Link")
            fraFields.Caption = .GetINIMessage("Game Link Selector", "ClearPrompt", "Clear which link?")
        End If
        optField(0).Caption = .GetINIMessage("Game Link Selector", "Precursor", "&Precursor")
        optField(1).Caption = .GetINIMessage("Game Link Selector", "Sequel", "&Sequel")
        optField(2).Caption = .GetINIMessage("Game Link Selector", "Related", "&Related Game")
        optField(3).Caption = .GetINIMessage("Game Link Selector", "CloneOf", "Clone &Of")
        If g_blnGameLinkSetFlag = True Then
            strIs = .GetINIMessage("Game Link Selector", "Is", "is") & " "
            strTo = " " & .GetINIMessage("Game Link Selector", "To", "to") & " "
            optField(0).Caption = g_strMarkedGame & " " & strIs & optField(0).Caption & strTo & g_udtGame.strName
            optField(1).Caption = g_strMarkedGame & " " & strIs & optField(1).Caption & strTo & g_udtGame.strName
            optField(2).Caption = g_strMarkedGame & " " & strIs & optField(2).Caption & strTo & g_udtGame.strName
            optField(3).Caption = g_strMarkedGame & " " & strIs & optField(3).Caption & " " & g_udtGame.strName
        End If
    End With

    'set the initial field
    Select Case g_strGameLinkField
        Case "Prequel"
            optField(0).Value = True
        Case "Sequel"
            optField(1).Value = True
        Case "Related"
            optField(2).Value = True
        Case Else
            optField(3).Value = True
    End Select

    With g_udtGame
        'disable stuff if necessary
        If g_blnGameLinkSetFlag = False Then
            optField(0).Enabled = (.lngPrequel > 0)
            optField(1).Enabled = (.lngSequel > 0)
            optField(2).Enabled = (.lngRelated > 0)
            optField(3).Enabled = (.lngCloneOf > 0)
            cmdOK.Enabled = ((.lngPrequel > 0) Or (.lngSequel > 0) Or (.lngRelated > 0) Or (.lngCloneOf > 0))
        End If
    End With

    'assume user presses cancel
    g_strGameLinkField = ""

End Sub


Private Sub optField_Click(Index As Integer)

    cmdOK.Enabled = ((optField(0).Value = True) Or (optField(1).Value = True) Or (optField(2).Value = True) Or (optField(3).Value = True))

End Sub


