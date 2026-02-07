VERSION 5.00
Begin VB.Form frmEditMusicianProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New / Edit Musician"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmEditMusicianProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdPhoto 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtGroup 
      Height          =   285
      Left            =   120
      MaxLength       =   25
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      MaxLength       =   128
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblPhoto 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Photo:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nick:"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Group:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmEditMusicianProps"
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

Private m_blnActivated As Boolean
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub

Private Sub cmdClear_Click()

    lblPhoto.Caption = ""

End Sub

Private Sub cmdOK_Click()

    'save props
    With g_udtMusicianProps
        .blnOKClicked = True
        .strGroup = txtGroup.Text
        .strMusician = txtName.Text
        .strNick = txtNick.Text
        .strPhoto = lblPhoto.Caption
    End With

    'exit
    Unload Me

End Sub


Private Sub cmdPhoto_Click()

    If g_xobjFunc.DirExists(g_udtRegPathPhotos.strRealPath) = True Then
        With g_udtFileSelectorProps
            .blnOKClicked = False
            .strCaption = Replace(lblPrompt(3).Caption, "&", "")
            .strFolder = g_udtRegPathPhotos.strRealPath
            .strPattern = "*.bmp;*.gif;*.jpg"
            .strReturnedFile = ""
            .strSelectFile = lblPhoto.Caption
            frmFileSelect.Show vbModal, Me
            If .blnOKClicked = True Then
                lblPhoto.Caption = .strReturnedFile
            End If
        End With
    End If

End Sub

Private Sub Form_Activate()

    If m_blnActivated = False Then
        m_blnActivated = True
        
        'first time set the focus
        With g_udtMusicianProps
            If .strFocus = "Name" Then
                txtName.SetFocus
            ElseIf .strFocus = "Group" Then
                txtGroup.SetFocus
            ElseIf .strFocus = "Nick" Then
                txtNick.SetFocus
            Else
                If cmdPhoto.Enabled = True Then
                    cmdPhoto.SetFocus
                Else
                    cmdCancel.SetFocus
                End If
            End If
        End With
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2535    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4950     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_xobjFunc.GetINIMessage("General Buttons", "Save", "&Save")
    cmdCancel.Caption = g_strButtonCancelText
    cmdClear.Caption = g_strButtonClearText
    If g_udtMusicianProps.blnAdding = True Then
        Caption = GetINIMessage("Musician Properties Window", "AddCaption", "Add Musician")
    Else
        Caption = GetINIMessage("Musician Properties Window", "EditCaption", "Edit Musician")
    End If
    lblPrompt(0).Caption = GetINIMessage("Musician Properties Window", "Name", "&Name") & ":"
    lblPrompt(1).Caption = GetINIMessage("Musician Properties Window", "Group", "&Group") & ":"
    lblPrompt(2).Caption = GetINIMessage("Musician Properties Window", "Nick", "&Nick") & ":"
    lblPrompt(3).Caption = GetINIMessage("Musician Properties Window", "Photo", "&Photo") & ":"

    'set the starting values
    With g_udtMusicianProps
        txtName.Text = .strMusician
        txtGroup.Text = .strGroup
        txtNick.Text = .strNick
        lblPhoto.Caption = .strPhoto
        .blnOKClicked = False
    End With
    cmdPhoto.Enabled = (g_xobjFunc.DirExists(g_udtRegPathPhotos.strRealPath))
    cmdOK.Enabled = (txtName.Text <> "")
    
    m_blnActivated = False

End Sub

Private Sub txtGroup_GotFocus()

    With txtGroup
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtName_Change()

    cmdOK.Enabled = (Len(txtName.Text) > 0)

End Sub

Private Sub txtName_GotFocus()

    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtNick_GotFocus()

    With txtNick
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


