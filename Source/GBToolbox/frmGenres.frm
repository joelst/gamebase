VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGenres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genres Wizard"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmGenres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   0
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.Frame fraWizardSelect 
         Caption         =   "What do you want to do?"
         Height          =   4815
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   5535
         Begin VB.OptionButton optWizSelect 
            Caption         =   "Add Sub Genres"
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
            Index           =   4
            Left            =   240
            TabIndex        =   29
            Top             =   3840
            Width           =   4815
         End
         Begin VB.OptionButton optWizSelect 
            Caption         =   "Add Parent Genres"
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
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   3000
            Width           =   4815
         End
         Begin VB.OptionButton optWizSelect 
            Caption         =   "Remove Sub Genres"
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
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   2160
            Width           =   4815
         End
         Begin VB.OptionButton optWizSelect 
            Caption         =   "Remove Parent Genres"
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
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   4815
         End
         Begin VB.OptionButton optWizSelect 
            Caption         =   "Clear Genres"
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
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Lets you add Sub Genres to a single Parent Genre."
            Height          =   495
            Index           =   4
            Left            =   480
            TabIndex        =   34
            Top             =   4080
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Lets you add Parent Genres to the list of Genres."
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   33
            Top             =   3240
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Removes Sub Genres from the Genres list and lets you move affected Games to a new Genre."
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   32
            Top             =   2400
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Removes Parent Genres from the Genres list and lets you move affected Games to a new Genre."
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   31
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Clears the entire Genre and Parent Genre lists from the database.  All games will point back to [uncategorized]."
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   4815
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   2
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   14
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame1 
         Caption         =   "Clear Genres"
         Height          =   4815
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   5535
         Begin VB.Label lblPrompt 
            Caption         =   "All Games now point to the [Uncategorized] Genre."
            Height          =   495
            Index           =   21
            Left            =   240
            TabIndex        =   23
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Click"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   22
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Back"
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
            Index           =   19
            Left            =   720
            TabIndex        =   21
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            Caption         =   "to return to the Wizard selector screen."
            Height          =   255
            Index           =   18
            Left            =   1320
            TabIndex        =   20
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Click"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   19
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPrompt 
            Caption         =   "Close"
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
            Index           =   16
            Left            =   720
            TabIndex        =   18
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            Caption         =   "to return to the GameBase Toolbox."
            Height          =   255
            Index           =   15
            Left            =   1320
            TabIndex        =   17
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPrompt 
            Caption         =   "The Genres list has been cleared!"
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
            Index           =   14
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   4815
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   5
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   46
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame4 
         Caption         =   "Remove Parent Genres"
         Height          =   4815
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   5535
         Begin VB.CommandButton cmdRP1 
            Caption         =   "&Link"
            Height          =   375
            Left            =   2160
            TabIndex        =   54
            Top             =   3600
            Width           =   1095
         End
         Begin VB.ComboBox lstRP2 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   4200
            Width           =   5055
         End
         Begin MSComctlLib.ListView lvRP1 
            Height          =   1695
            Left            =   240
            TabIndex        =   51
            Top             =   1680
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Game Name"
               Object.Width           =   3835
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Current Genre"
               Object.Width           =   4390
            EndProperty
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "When there are no more games left to link, click"
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   55
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   3495
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Genre:"
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   52
            Top             =   3960
            UseMnemonic     =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Games to Re-Categorise:"
            Height          =   255
            Index           =   31
            Left            =   240
            TabIndex        =   50
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmGenres.frx":0442
            Height          =   735
            Index           =   30
            Left            =   240
            TabIndex        =   49
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
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
            Index           =   28
            Left            =   3720
            TabIndex        =   48
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   4
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   41
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Remove Parent Genres"
         Height          =   4815
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   5535
         Begin VB.ListBox lstRP1 
            Height          =   3765
            Left            =   240
            TabIndex        =   43
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
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
            Index           =   27
            Left            =   3480
            TabIndex        =   45
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Select a Parent Genre to remove, then click"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   44
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   3255
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   3
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   35
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Remove Parent Genres"
         Height          =   4815
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   5535
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Click"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   40
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Back"
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
            Index           =   24
            Left            =   720
            TabIndex        =   39
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "to return to the Wizard selector screen."
            Height          =   255
            Index           =   23
            Left            =   1320
            TabIndex        =   38
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "There are no Parent Genres to remove!"
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
            Index           =   22
            Left            =   240
            TabIndex        =   37
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   4815
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   1
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      Begin VB.Frame fraClearGenres 
         Caption         =   "Clear Genres"
         Height          =   4815
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5535
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Are you sure you want to clear the entire Genres list?"
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
            Index           =   13
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "to return to the Wizard selector screen."
            Height          =   255
            Index           =   12
            Left            =   1320
            TabIndex        =   12
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Back"
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
            Index           =   11
            Left            =   720
            TabIndex        =   11
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Click"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   10
            Top             =   3000
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "to clear the Genres list."
            Height          =   255
            Index           =   9
            Left            =   1320
            TabIndex        =   9
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Clear"
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
            Index           =   8
            Left            =   720
            TabIndex        =   8
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "Click"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   7
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPrompt 
            BackStyle       =   0  'Transparent
            Caption         =   "This will remove all Genres from your database and re-categorise all games to [Uncategorized]."
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   6
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   5055
         End
      End
   End
End
Attribute VB_Name = "frmGenres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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



Private db As DAO.Database
Private rs As DAO.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()

    Set db = OpenDatabase(g_strDatabase)

End Sub


Private Sub Form_Unload(Cancel As Integer)

    db.Close
    Set db = Nothing

End Sub


Private Sub lvRP1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Debug.Print ColumnHeader.Width

End Sub


