VERSION 5.00
Begin VB.Form frmFirstLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GameBase"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmFirstLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkGemusAsk 
      Caption         =   "&Setup my emulators now with the GEMUS Setup Wizard"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
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

Private Sub cmdOK_Click()

    With g_udtFirstLoad
        If .blnGemus = True Then
            .blnGemus = (chkGemusAsk.Value = vbChecked)
        End If
    End With
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3630    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5895     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    Caption = "[" & g_strGameBaseName & "]" & " GameBase"
    chkGemusAsk.Caption = g_xobjFunc.GetINIMessage("FirstLoad Window", "GemusPrompt2", "&Setup my emulators now with the GEMUS Setup Wizard")
    cmdOK.Caption = g_strButtonOKText
    With g_udtFirstLoad
        txtMessage.Text = .strMessage
        chkGemusAsk.Visible = .blnGemus
        chkGemusAsk.Value = vbChecked
    End With
    txtMessage.Visible = True

End Sub


