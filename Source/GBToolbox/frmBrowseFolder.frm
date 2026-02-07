VERSION 5.00
Begin VB.Form frmBrowseFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder..."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmBrowseFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.DirListBox DirBox 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frmBrowseFolder"
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
'                The GBToolbox
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

Private m_strLastValidPath As String    'the last valid non-error selected path
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    g_strBrowseForFolder = lblPath
    Unload Me

End Sub


Private Sub DirBox_Change()

    lblPath = DirBox.List(DirBox.ListIndex)
    m_strLastValidPath = DirBox.Path

End Sub

Private Sub DirBox_Click()

    DirBox.Path = DirBox.List(DirBox.ListIndex)

End Sub


Private Sub DriveBox_Change()

    On Error GoTo Err_Handler
    
    MousePointer = vbHourglass
    DoEvents
    DirBox.Path = Left$(DriveBox.Drive, 2) & "\"
    MousePointer = vbDefault
    Exit Sub

Err_Handler:
    MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "GBToolbox"
    On Error GoTo 0
    DriveBox.Drive = Left$(m_strLastValidPath, 2) & "\"
    DirBox.Path = m_strLastValidPath

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3855    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4695     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'assumption: If (g_strBrowseForFolder <> "") then contains a _valid_ path
    If g_strBrowseForFolder <> "" Then
        DriveBox.Drive = Left$(g_strBrowseForFolder, 2) & "\"
        DirBox.Path = g_strBrowseForFolder
        m_strLastValidPath = g_strBrowseForFolder
        lblPath = DirBox.Path
    Else
        'set to app.path
        DriveBox.Drive = Left$(App.Path, 2) & "\"
        DirBox.Path = App.Path
        m_strLastValidPath = App.Path
        lblPath = DirBox.Path
    End If

    'return an empty string unless cmdOK is clicked
    g_strBrowseForFolder = ""

End Sub


