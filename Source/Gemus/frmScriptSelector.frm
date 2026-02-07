VERSION 5.00
Begin VB.Form frmScriptSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GEMUS"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmScriptSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Load"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.FileListBox FileBox 
      Height          =   2235
      Left            =   0
      Pattern         =   "*.txt"
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmScriptSelector"
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

    Unload Me

End Sub

Private Sub cmdOK_Click()

    g_udtScriptSelectorType.blnOKClicked = True
    g_udtScriptSelectorType.strSelectedFile = FileBox.List(FileBox.ListIndex)
    Unload Me

End Sub

Private Sub FileBox_Click()

    cmdOK.Enabled = (FileBox.ListIndex > -1)

End Sub


Private Sub FileBox_DblClick()

    DoEvents
    If FileBox.ListIndex > -1 Then
        g_udtScriptSelectorType.blnOKClicked = True
        g_udtScriptSelectorType.strSelectedFile = FileBox.List(FileBox.ListIndex)
        Unload Me
    End If

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2895    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 3855     'The required scalewidth for the form
    Dim lngIdx As Long
    Dim strAppPathWithBackSlash As String
    Dim strMessFile As String

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the text
    strAppPathWithBackSlash = App.Path
    If Right$(strAppPathWithBackSlash, 1) <> "\" Then strAppPathWithBackSlash = strAppPathWithBackSlash & "\"
    strMessFile = strAppPathWithBackSlash & "GemusText.ini"
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "Load", "&Load", strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", strMessFile)

    With FileBox
        .Path = g_udtScriptSelectorType.strInitPath
        .Refresh
        If .ListCount > 0 Then
            For lngIdx = 0 To .ListCount - 1
                If UCase$(.List(lngIdx)) = UCase$(g_udtScriptSelectorType.strSelectedFile) Then
                    .ListIndex = lngIdx
                    Exit For
                End If
            Next lngIdx
            If .ListIndex < 0 Then
                .ListIndex = 0
            End If
        End If
    End With

    g_udtScriptSelectorType.blnOKClicked = False

End Sub


