VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "frmNewPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConfirm 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Please remember the password you type in here, you will need it to change this setting in the future."
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a password that will need to be typed when enabling and disabling the Adult filter."
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewPassword"
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

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    'get password from textbox and store in the registry
    g_strAdultPassword = Muddle(txtPassword)
    g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\OregonASFD", "Check", g_strAdultPassword, True

    Unload Me

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3975    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 2550     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    'read in the ini file text
    With g_xobjFunc
        Caption = .GetINIMessage("Password Window", "Caption", "Adult Filter")
        lblPrompt = .GetINIMessage("Password Window", "Prompt", "Enter a password that will need to be typed when disabling the Adult filter.")
        lblPassword = .GetINIMessage("Password Window", "Password", "Password") & ":"
        lblConfirm = .GetINIMessage("Password Window", "Confirm", "Confirm") & ":"
        lblWarning = .GetINIMessage("Password Window", "Warning", "Please remember the password you type in here, you will need it to change this setting in the future.")
    End With
    
    'set the button text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    
    'no password yet
    g_strAdultPassword = ""

End Sub


Private Sub txtConfirm_Change()

    cmdOK.Enabled = (txtPassword <> "") And (txtPassword = txtConfirm)

End Sub

Private Sub txtConfirm_GotFocus()
    
    With txtConfirm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtPassword_Change()

    cmdOK.Enabled = (txtPassword <> "") And (txtPassword = txtConfirm)

End Sub


Private Sub txtPassword_GotFocus()
    
    With txtPassword
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


