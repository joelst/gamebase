VERSION 5.00
Begin VB.Form frmGetID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select 'Affected Games' Record..."
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmGetID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstRecords 
      Enabled         =   0   'False
      Height          =   1815
      ItemData        =   "frmGetID.frx":0442
      Left            =   120
      List            =   "frmGetID.frx":0444
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Select a record to link any affected Games to:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmGetID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdOK_Click()

    If lstRecords.Enabled = True And lstRecords.ListIndex > -1 Then
        g_lngMoveToID = lstRecords.ItemData(lstRecords.ListIndex)
        g_strMoveToName = lstRecords.List(lstRecords.ListIndex)
    End If
    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2895    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4230     'The required scalewidth for the form
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fldID As Field
    Dim fldName As Field

    On Error GoTo Err_Handler

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    'read in the list
    Set db = OpenDatabase(g_strRemoveDB)
    Set rs = db.OpenRecordset(g_strDelSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        Set fldID = rs!ID
        Set fldName = rs!Name
        rs.MoveFirst
        With lstRecords
            Do Until rs.EOF
                .AddItem fldName
                .ItemData(.NewIndex) = fldID
                rs.MoveNext
            Loop
        End With
        Set fldID = Nothing
        Set fldName = Nothing
    End If
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    
    'done
    lstRecords.Visible = True
    If lstRecords.ListCount > 0 Then
        lstRecords.Enabled = True
        lstRecords.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
    DoEvents
    g_lngMoveToID = 0
    Exit Sub

Err_Handler:
    Screen.MousePointer = vbDefault
    DoEvents
    g_lngMoveToID = 0
    g_strMoveToName = ""
    
End Sub



Private Sub lstRecords_Click()
    
    If lstRecords.ListIndex > -1 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
End Sub


Private Sub lstRecords_DblClick()

    DoEvents
    If cmdOK.Enabled = True Then
        cmdOK.Value = True
    End If

End Sub


