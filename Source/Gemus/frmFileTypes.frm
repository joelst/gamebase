VERSION 5.00
Begin VB.Form frmFileTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Emulator FileTypes"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmFileTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      MaxLength       =   4096
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblDelim 
      BackStyle       =   0  'Transparent
      Caption         =   "DelimText"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Default Native Emulator File Types:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmFileTypes"
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

Private m_xobjFunc As GBFuncs.clsFunctions
Private m_strMessFile As String
Private m_strAppPathWithBackSlash As String
Private m_strEmuFile As String
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If m_xobjFunc.INI_WriteSetting("File Types", "SupportTypes", txtInput.Text, m_strEmuFile) = True Then
        g_blnFileTypesChanged = True
    Else
        MsgBox "There was an error saving the Supported File Types to " & m_strEmuFile, vbExclamation, "GEMUS"
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 1335    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5775     'The required scalewidth for the form
    Dim strFileTypes As String

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    m_strAppPathWithBackSlash = App.Path
    If Right$(m_strAppPathWithBackSlash, 1) <> "\" Then m_strAppPathWithBackSlash = m_strAppPathWithBackSlash & "\"
    m_strMessFile = m_strAppPathWithBackSlash & "GemusText.ini"
    m_strEmuFile = g_strFileTypesEmuFile

    'buttons
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", m_strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", m_strMessFile)
    Caption = GetINIMessageFromFile("Supported Filetypes Window", "Caption", "GEMUS Default Emulator File Types", m_strMessFile)
    lblPrompt.Caption = GetINIMessageFromFile("Supported Filetypes Window", "Prompt", "&Native Emulator File Types", m_strMessFile) & ":"
    lblDelim.Caption = GetINIMessageFromFile("Supported Filetypes Window", "DelimText", "Delimit with", m_strMessFile) & " ;"

    Set m_xobjFunc = New GBFuncs.clsFunctions
    strFileTypes = m_xobjFunc.INI_GetSetting("File Types", "SupportTypes", "", m_strEmuFile)
    txtInput = strFileTypes
    g_blnFileTypesChanged = False

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set m_xobjFunc = Nothing

End Sub


Private Sub txtInput_Change()

    Dim lngIdx As Long
    
    With txtInput
        If .Text = "" Then
            cmdOK.Enabled = True
            Exit Sub
        End If
        If InStr(.Text, ";;") > 0 Then
            cmdOK.Enabled = False
            Exit Sub
        End If
        For lngIdx = 1 To Len(.Text)
            If Mid$(.Text, lngIdx, 1) <> ";" Then
                cmdOK.Enabled = True
                Exit Sub
            End If
        Next lngIdx
        cmdOK.Enabled = False
    End With
    
End Sub

Private Sub txtInput_GotFocus()

    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)

    If InStr(" \/:*?""<>|", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub


