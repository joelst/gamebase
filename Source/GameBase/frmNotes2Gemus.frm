VERSION 5.00
Begin VB.Form frmNotes2Gemus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GameBase Database Update"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optGemus 
      Caption         =   "My Notes field is ONLY used for ""key=value"" pairs."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4335
   End
   Begin VB.OptionButton optGemus 
      Caption         =   "I use my Notes field for ""key=value"" pairs and other stuff too."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4815
   End
   Begin VB.OptionButton optGemus 
      Caption         =   "I don't use any ""key=value"" pairs in my Notes field."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Value           =   -1  'True
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNotes2Gemus.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNotes2Gemus.frx":010B
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmNotes2Gemus"
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
    Hide
End Sub


