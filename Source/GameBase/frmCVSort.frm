VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCVSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sort Custom Views"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmCVSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ListBox lstPaths 
      Height          =   4935
      ItemData        =   "frmCVSort.frx":08CA
      Left            =   120
      List            =   "frmCVSort.frx":08CC
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin MSComctlLib.ImageList iLstToolbar 
      Left            =   6360
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCVSort.frx":08CE
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCVSort.frx":096A
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCVSort.frx":09FE
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCVSort.frx":0A90
            Key             =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbButtons 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "iLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "up"
            ImageKey        =   "up"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "down"
            ImageKey        =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCVSort"
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

Private Sub SetMoveButtonsState()

    With lstPaths
        tbButtons.Buttons("up").Enabled = (.ListIndex > 0)
        tbButtons.Buttons("down").Enabled = (.ListIndex < .ListCount - 1)
    End With

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdOK_Click()

    Dim lngIdx As Long

    'save back the view ordinals
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ID, Ordinal FROM ViewData", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            For lngIdx = 0 To lstPaths.ListCount - 1
                .FindFirst "ID = " & lstPaths.ItemData(lngIdx)
                If Not .NoMatch Then
                    .Edit
                    !Ordinal = lngIdx
                    .Update
                End If
            Next lngIdx
        End If
    End With
    Set g_rsGames = Nothing

    g_blnViewsWereSorted = True
    Unload Me

End Sub


Private Sub Form_Activate()

    If lstPaths.Enabled = True Then
        lstPaths.SetFocus
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 6015    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6255     'The required scalewidth for the form
    Dim lngIdx As Long

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'read in INI text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    Caption = g_xobjFunc.GetINIMessage("Sort Custom Views Window", "Caption", "Sort Custom Views")

    'load the custom views from the database
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ID, Title FROM ViewData ORDER BY Ordinal", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            lngIdx = 0
            Do Until .EOF
                lstPaths.AddItem !Title, lngIdx
                lstPaths.ItemData(lngIdx) = !ID
                lngIdx = lngIdx + 1
                .MoveNext
            Loop
            .Close
        Else
            'no custom views
            cmdOK.Enabled = False
            lstPaths.Enabled = False
        End If
    End With
    Set g_rsGames = Nothing

    'can only sort if there is more than one custom view
    If lstPaths.ListCount = 1 Then
        cmdOK.Enabled = False
        lstPaths.Enabled = False
    ElseIf lstPaths.ListCount > 1 Then
        lstPaths.ListIndex = 0
    End If
    Call SetMoveButtonsState
    
    g_blnViewsWereSorted = False

End Sub

Private Sub lstPaths_Click()

    Call SetMoveButtonsState

End Sub


Private Sub tbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim lngItemData As Long
    Dim strList As String
    Dim lngListIndex As Long

    With lstPaths

        lngListIndex = .ListIndex
        If lngListIndex = -1 Then Exit Sub
    
        Select Case Button.Key
        
            Case "up"
            
                If lngListIndex = 0 Then Exit Sub
                strList = .List(lngListIndex - 1)
                lngItemData = .ItemData(lngListIndex - 1)
                .List(lngListIndex - 1) = .List(lngListIndex)
                .ItemData(lngListIndex - 1) = .ItemData(lngListIndex)
                .List(lngListIndex) = strList
                .ItemData(lngListIndex) = lngItemData
                .ListIndex = lngListIndex - 1
                .SetFocus

            Case "down"

                If lngListIndex = .ListCount - 1 Then Exit Sub
                strList = .List(lngListIndex + 1)
                lngItemData = .ItemData(lngListIndex + 1)
                .List(lngListIndex + 1) = .List(lngListIndex)
                .ItemData(lngListIndex + 1) = .ItemData(lngListIndex)
                .List(lngListIndex) = strList
                .ItemData(lngListIndex) = lngItemData
                .ListIndex = lngListIndex + 1
                .SetFocus

        End Select

    End With

End Sub


