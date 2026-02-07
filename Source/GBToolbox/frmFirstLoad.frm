VERSION 5.00
Begin VB.Form frmFirstLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Settings"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmFirstLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGameBaseWindowTitle 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   2
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   2775
      Left            =   120
      MaxLength       =   65526
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox chkGemusAsk 
      Caption         =   "&Prompt for GEMUS Setup Wizard"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "(appears in GameBase About Screen)"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblPrompt 
      Caption         =   "GameBase Window Title:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Database Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblPrompt 
      Caption         =   "First Load Message:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmFirstLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2010 James Burrows
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

Private m_db As DAO.Database
Private m_rs As DAO.Recordset

Private Sub cmdCancel_Click()
    
    m_db.Close
    Set m_db = Nothing

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Set m_rs = m_db.OpenRecordset("SELECT DatabaseName, GameBaseWindowTitle, FirstLoadMessage, FirstLoadGemusAsk FROM Config", dbOpenDynaset)
    With m_rs
        .Edit
        !DatabaseName = txtDatabaseName.Text
        !GameBaseWindowTitle = txtGameBaseWindowTitle.Text
        !FirstLoadMessage = Left$(txtMessage.Text, gc_lngMAXMEMOTEXTLENGTH)
        !FirstLoadGemusAsk = chkGemusAsk.Value
        .Update
        .Close
    End With
    m_db.Close
    Set m_rs = Nothing
    Set m_db = Nothing

    Unload Me

End Sub


Private Sub Form_Activate()

    txtDatabaseName.SetFocus

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 5295    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5895     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    Set m_db = OpenDatabase(g_strDatabase)
    Set m_rs = m_db.OpenRecordset("SELECT DatabaseName, GameBaseWindowTitle, FirstLoadMessage, FirstLoadGemusAsk FROM Config", dbOpenSnapshot)
    txtDatabaseName.Text = m_rs!DatabaseName
    txtGameBaseWindowTitle.Text = m_rs!GameBaseWindowTitle
    txtMessage.Text = m_rs!FirstLoadMessage
    If m_rs!FirstLoadGemusAsk = 0 Then
        chkGemusAsk.Value = vbUnchecked
    Else
        chkGemusAsk.Value = vbChecked
    End If
    m_rs.Close
    Set m_rs = Nothing

    txtMessage.Visible = True

End Sub



