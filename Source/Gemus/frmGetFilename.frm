VERSION 5.00
Begin VB.Form frmGetFilename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save GEMUS Script File"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmGetFilename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   120
      MaxLength       =   61
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".TXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4150
      TabIndex        =   4
      Top             =   390
      Width           =   435
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Script &Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmGetFilename"
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

Private m_strFolder As String
Private m_strPrompt As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim xobjFunc As GBFuncs.clsFunctions
    Dim f As Integer
    Dim strPrompt As String
    Dim blnAargh As Boolean

    'check doesn't exits already
    Set xobjFunc = New GBFuncs.clsFunctions
    blnAargh = xobjFunc.FileExists(m_strFolder & txtFilename.Text & ".txt", False, "GEMUS")
    Set xobjFunc = Nothing
    If blnAargh = True Then
        If MsgBox(m_strPrompt, vbYesNoCancel, "GEMUS") = vbYes Then
            blnAargh = False
        End If
    End If
    
    If blnAargh = False Then
        g_strScriptSaveVar = txtFilename.Text & ".txt"
    Else
        g_strScriptSaveVar = ""
    End If
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 1335    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4680     'The required scalewidth for the form
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
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "Save", "&Save", strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", strMessFile)
    Caption = GetINIMessageFromFile("Get Script Filename Window", "Caption", "Save GEMUS Sript", strMessFile)
    lblPrompt.Caption = GetINIMessageFromFile("Get Script Filename Window", "Prompt", "Script &Filename", strMessFile) & "..."
    m_strPrompt = GetINIMessageFromFile("Get Script Filename Window", "OverwritePrompt", "This file exists.  Overwrite?", strMessFile)
    
    m_strFolder = g_strScriptSaveVar
    g_strScriptSaveVar = ""

End Sub


Private Sub txtFilename_Change()

    cmdOK.Enabled = (txtFilename.Text <> "")

End Sub


Private Sub txtFilename_GotFocus()

    With txtFilename
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtFilename_KeyPress(KeyAscii As Integer)

    If InStr("\/:*?""<>|", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub


