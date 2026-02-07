VERSION 5.00
Begin VB.Form frmGB64Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select file to load into the emulators"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.FileListBox FileBox 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "This archive file has more than one compatible game file.  Please select the file to load into the emulators from the list below:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmGB64Main"
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

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()

    g_strDefaultFile = lstFiles.List(lstFiles.ListIndex)
    Hide

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3255    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5055     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    Caption = g_strArchiveCaption
    lblPrompt = g_strArchivePrompt
    cmdOK.Caption = g_strOKButton
    cmdCancel.Caption = g_strCancelButton
    
End Sub


Private Sub lstFiles_DblClick()
    DoEvents
    g_strDefaultFile = lstFiles.List(lstFiles.ListIndex)
    Hide
End Sub


