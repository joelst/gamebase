VERSION 5.00
Begin VB.Form frmVersionCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GameBase Version Checker"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "frmVersionCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkStats 
      Caption         =   "Send anonymous usage statistics on startup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CheckBox chkVersion 
      Caption         =   "Check for new &Version of GameBase on startup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVersionCheck.frx":08CA
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The emulators you are using."
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The databases you are using."
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVersionCheck.frx":0957
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You will need to be connected to the internet for this to work."
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "frmVersionCheck"
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

    Dim blnWritten As Boolean

    blnWritten = g_xobjFunc.INI_WriteSetting("Startup", "VersionCheck", IIf((chkVersion.Value = vbChecked), "1", "0"), g_strGameBaseDotINI)
    blnWritten = blnWritten Or g_xobjFunc.INI_WriteSetting("Startup", "SendUsageStats", IIf((chkStats.Value = vbChecked), "1", "0"), g_strGameBaseDotINI)

    If Not blnWritten Then
        MsgBox "Unable to save the Version Check settings in " & g_strGameBaseDotINI, vbExclamation Or vbOKOnly, gc_strAPPNAME
    End If

    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3855    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5190     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

End Sub


