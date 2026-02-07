VERSION 5.00
Begin VB.Form frmFileSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Select"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmFileSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox FileBox 
      Height          =   3600
      Left            =   0
      Pattern         =   "*.gif;*.jpg;*.bmp"
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmFileSelect"
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

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    'save props
    With g_udtFileSelectorProps
        .blnOKClicked = True
        .strReturnedFile = FileBox.List(FileBox.ListIndex)
    End With

    'exit
    Unload Me

End Sub


Private Sub FileBox_Click()

    cmdOK.Enabled = (FileBox.ListIndex > -1)

End Sub

Private Sub FileBox_DblClick()

    If (FileBox.ListIndex > -1) And (cmdOK.Enabled = True) Then
        cmdOK.Value = True
    End If

End Sub


Private Sub Form_Load()

    Dim strUpper As String
    Dim lngIdx As Long
    Const DesignScaleHeight As Single = 4335    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4095     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the starting values
    With g_udtFileSelectorProps
        Caption = .strCaption
        FileBox.Path = .strFolder
        FileBox.Pattern = .strPattern
        FileBox.Refresh
        .blnOKClicked = False
        If .strSelectFile <> "" Then
            strUpper = UCase$(.strSelectFile)
            With FileBox
                .ListIndex = -1
                For lngIdx = 0 To (.ListCount - 1)
                    If UCase$(.List(lngIdx)) = strUpper Then
                        .ListIndex = lngIdx
                        Exit For
                    End If
                Next lngIdx
            End With
        End If
    End With
    cmdOK.Enabled = (FileBox.ListIndex > -1)
    
End Sub


