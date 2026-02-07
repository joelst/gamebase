VERSION 5.00
Begin VB.Form frmInputBox2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Box 2"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmInputBox2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFlag 
      Caption         =   "Flag"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmInputBox2"
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

Private m_strInvalid As String
Private m_blnAllowEmptyReturn As Boolean

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim lngCharIdx As Long

    If m_strInvalid <> "" Then
        If txtInput.Text <> "" Then
            For lngCharIdx = 1 To Len(txtInput.Text)
                If InStr(m_strInvalid, Mid$(txtInput.Text, lngCharIdx, 1)) > 0 Then
                    MsgBox g_xobjFunc.GetINIMessage("Messages", "InvalidChars", "You have entered one or more invalid characters!"), vbExclamation, gc_strAPPNAME
                    txtInput.SetFocus
                    Exit Sub
                End If
            Next lngCharIdx
        End If
    End If

    g_udtInputBox2.blnOKClicked = True
    g_udtInputBox2.strReturnedText = txtInput.Text
    g_udtInputBox2.lngCheckBoxReturnedValue = chkFlag.Value
    Unload Me

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 1695    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4575     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the properties
    With g_udtInputBox2
        Caption = .strCaption
        chkFlag.Value = .lngCheckBoxStartingValue
        chkFlag.Caption = .strCheckBoxText
        lblPrompt = .strPrompt
        txtInput.MaxLength = .lngMaxLength
        m_strInvalid = .strInvalidChars
        cmdOK.Caption = .strOKButtonText
        cmdCancel.Caption = .strCancelButtonText
        .strReturnedText = ""
        .blnOKClicked = False
        m_blnAllowEmptyReturn = .blnAllowEmptyReturn
        txtInput = .strStartingText
        If m_blnAllowEmptyReturn Then
            cmdOK.Enabled = True
        End If
    End With

End Sub


Private Sub txtInput_Change()

    If Not m_blnAllowEmptyReturn Then
        cmdOK.Enabled = (txtInput <> "")
    End If

End Sub


Private Sub txtInput_GotFocus()

    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)

    If m_strInvalid <> "" Then
        If InStr(m_strInvalid, Chr$(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If

End Sub



