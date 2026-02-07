VERSION 5.00
Begin VB.Form frmNotes 
   Caption         =   "Game Notes"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6135
   Begin VB.TextBox txtNotes 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   65526
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmNotes"
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

Public Sub SetNotes(strNotes As String)

    If Len(strNotes) > 65526 Then
        txtNotes.Text = Left$(strNotes, 65526)
    Else
        txtNotes.Text = strNotes
    End If

End Sub


Private Sub Form_Load()

    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord

    g_blnLoadedfrmNotes = True

    'get current size/position/state variables
    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesWidth", sngRegWidth, False, 6135
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesHeight", sngRegHeight, False, 4905
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesTop", sngRegTop, False, gc_lngINVALIDWINDOWPOS
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesLeft", sngRegLeft, False, gc_lngINVALIDWINDOWPOS
    End With

    'make sure not to big or small...
    If sngRegWidth <= 100 Then
        sngRegWidth = 6135
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight <= 100 Then
        sngRegHeight = 4905
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    '...or out of full screen view
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (Height / 2)
        sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    'set the position and size
    Top = sngRegTop
    Left = sngRegLeft
    Width = sngRegWidth
    Height = sngRegHeight

End Sub


Private Sub Form_Resize()

    If ScaleWidth > 100 Then
        txtNotes.Width = ScaleWidth
    End If
    If ScaleHeight > 100 Then
        txtNotes.Height = ScaleHeight
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    If WindowState <> vbMinimized And WindowState <> vbMaximized Then
        With g_xobjRegistry
            'main window size/position
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "NotesLeft", Left, False
        End With
    End If

    g_blnLoadedfrmNotes = False

End Sub


