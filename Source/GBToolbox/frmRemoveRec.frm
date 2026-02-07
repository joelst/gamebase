VERSION 5.00
Begin VB.Form frmRemoveRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Records from Table..."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmRemoveRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete..."
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox lstRecords 
      Enabled         =   0   'False
      Height          =   1815
      ItemData        =   "frmRemoveRec.frx":0442
      Left            =   120
      List            =   "frmRemoveRec.frx":0444
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.ComboBox lstTable 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the &Record to Delete:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the &Table:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRemoveRec"
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

Private ma_strSQL(8) As String
Private db As DAO.Database
Private rs As DAO.Recordset
Private Sub cmdClose_Click()

    Unload Me

End Sub


Private Sub cmdDelete_Click()

    Dim strTable As String
    Dim strField As String
    Dim lngOldIdx As Long

    If lstRecords.ListIndex > -1 And lstTable.ListIndex > -1 Then
        Select Case lstTable.ListIndex
            Case 0
                'publishers
                g_strDelSQL = "SELECT PU_Id as ID, Publisher AS Name FROM Publishers WHERE NOT (PU_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Publisher"
                strTable = "Publishers"
                strField = "PU_Id"
            Case 1
                'musicians
                g_strDelSQL = "SELECT MU_Id as ID, Musician AS Name FROM Musicians WHERE NOT (MU_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Musician"
                strTable = "Musicians"
                strField = "MU_Id"
            Case 2
                'programmers
                g_strDelSQL = "SELECT PR_Id as ID, Programmer AS Name FROM Programmers WHERE NOT (PR_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Programmer"
                strTable = "Programmers"
                strField = "PR_Id"
            Case 3
                'languages
                g_strDelSQL = "SELECT LA_Id as ID, Language AS Name FROM Languages WHERE NOT (LA_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Language"
                strTable = "Languages"
                strField = "LA_Id"
            Case 4
                'crackers
                g_strDelSQL = "SELECT CR_Id as ID, Cracker AS Name FROM Crackers WHERE NOT (CR_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Cracker"
                strTable = "Crackers"
                strField = "CR_Id"
            Case 5
                'artists
                g_strDelSQL = "SELECT AR_Id as ID, Artist AS Name FROM Artists WHERE NOT (AR_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Artist"
                strTable = "Artists"
                strField = "AR_Id"
            Case 6
                'developers
                g_strDelSQL = "SELECT DE_Id as ID, Developer AS Name FROM Developers WHERE NOT (DE_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Developer"
                strTable = "Developers"
                strField = "DE_Id"
            Case 7
                'licenses
                g_strDelSQL = "SELECT LI_Id as ID, License AS Name FROM Licenses WHERE NOT (LI_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY License"
                strTable = "Licenses"
                strField = "LI_Id"
            Case Else
                'rarities
                g_strDelSQL = "SELECT RA_Id as ID, Rarity AS Name FROM Rarities WHERE NOT (RA_Id = " & lstRecords.ItemData(lstRecords.ListIndex) & ") ORDER BY Rarity"
                strTable = "Rarities"
                strField = "RA_Id"
        End Select
        Screen.MousePointer = vbHourglass
        DoEvents
        frmGetID.Show vbModal, Me
        If g_lngMoveToID > 0 Then
            If MsgBox("Delete the record """ & lstRecords.List(lstRecords.ListIndex) & """ from the table """ & lstTable.List(lstTable.ListIndex) & """ and" & vbCrLf & "move all affected games to the record """ & g_strMoveToName & """?" & vbCrLf & vbCrLf & "ARE YOU SURE?   THIS CANNOT BE UNDONE!", vbQuestion Or vbYesNo) = vbYes Then
                'move all games to new id
                db.Execute "UPDATE Games SET " & strField & "=" & g_lngMoveToID & " WHERE " & strField & "=" & lstRecords.ItemData(lstRecords.ListIndex)
                'remove record from db
                db.Execute "DELETE * FROM " & strTable & " WHERE " & strField & "=" & lstRecords.ItemData(lstRecords.ListIndex)
                'remove record from list
                lngOldIdx = lstRecords.ListIndex
                lstRecords.RemoveItem lngOldIdx
                If (lstRecords.ListCount - 1) < lngOldIdx Then
                    lstRecords.ListIndex = lngOldIdx - 1
                Else
                    lstRecords.ListIndex = lngOldIdx
                    cmdDelete.Enabled = True
                End If
                MsgBox "Record Deleted!", vbInformation, "GBToolbox"
                If lstRecords.Enabled = True Then
                    lstRecords.SetFocus
                End If
            End If
        Else
            'cancelled
            If lstRecords.Enabled = True Then
                lstRecords.SetFocus
            End If
        End If
    Else
        cmdDelete.Enabled = False
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3615    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4575     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'fill the table listbox
    lstTable.AddItem "Publishers"
    lstTable.AddItem "Musicians"
    lstTable.AddItem "Programmers"
    lstTable.AddItem "Languages"
    lstTable.AddItem "Crackers"
    lstTable.AddItem "Artists"
    lstTable.AddItem "Developers"
    lstTable.AddItem "Licenses"
    lstTable.AddItem "Rarities"

    ma_strSQL(0) = "SELECT PU_Id AS ID, Publisher AS Name FROM Publishers WHERE NOT (PU_ID = 37 OR PU_ID = 39) ORDER BY Publisher"
    ma_strSQL(1) = "SELECT MU_Id AS ID, Musician AS Name FROM Musicians WHERE NOT (MU_ID = 38 OR MU_ID = 39) ORDER BY Musician"
    ma_strSQL(2) = "SELECT PR_Id AS ID, Programmer AS Name FROM Programmers WHERE NOT (PR_ID = 1) ORDER BY Programmer"
    ma_strSQL(3) = "SELECT LA_Id AS ID, Language AS Name FROM Languages WHERE NOT (LA_ID = 1) ORDER BY Language"
    ma_strSQL(4) = "SELECT CR_Id AS ID, Cracker AS Name FROM Crackers WHERE NOT (CR_ID = 7 OR CR_ID = 8) ORDER BY Cracker"
    ma_strSQL(5) = "SELECT AR_Id AS ID, Artist AS Name FROM Artists WHERE NOT (AR_ID = 1 OR AR_ID = 2) ORDER BY Artist"
    ma_strSQL(6) = "SELECT DE_Id AS ID, Developer AS Name FROM Developers WHERE NOT (DE_ID = 1 OR DE_ID = 2) ORDER BY Developer"
    ma_strSQL(7) = "SELECT LI_Id AS ID, License AS Name FROM Licenses WHERE NOT (LI_ID = 1) ORDER BY License"
    ma_strSQL(8) = "SELECT RA_Id AS ID, Rarity AS Name FROM Rarities WHERE NOT (RA_ID = 1) ORDER BY Rarity"

    Set db = OpenDatabase(g_strRemoveDB)

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set rs = Nothing
    db.Close
    Set db = Nothing

End Sub


Private Sub lstRecords_Click()

    cmdDelete.Enabled = (lstRecords.ListIndex > -1)

End Sub

Private Sub lstRecords_DblClick()

    DoEvents
    If cmdDelete.Enabled = True Then
        cmdDelete.Value = True
    End If
    
End Sub


Private Sub lstTable_Click()

    On Error GoTo Err_Handler

    Dim fldName As Field
    Dim fldID As Field

    lstTable.Enabled = False
    cmdDelete.Enabled = False
    cmdClose.Enabled = False
    MousePointer = vbHourglass
    DoEvents
    lstRecords.Enabled = False
    lstRecords.Visible = False
    lstRecords.Clear
    If lstTable.ListIndex > -1 Then
        Set rs = db.OpenRecordset(ma_strSQL(lstTable.ListIndex), dbOpenSnapshot)
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
    End If
    lstTable.Enabled = True
    lstRecords.Visible = True
    cmdClose.Enabled = True
    If lstRecords.ListCount = 0 Then
        lstRecords.Enabled = False
    Else
        lstRecords.ListIndex = 0
        lstRecords.Enabled = True
        DoEvents
        lstRecords.SetFocus
    End If
    MousePointer = vbDefault
    

    Exit Sub

Err_Handler:
    MousePointer = vbDefault
    MsgBox Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, "GBToolbox"

End Sub


