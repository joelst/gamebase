VERSION 5.00
Begin VB.Form frmExtrasEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extras Manager"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmExtrasEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAll 
      Caption         =   "&Apply to all games in current view (i.e. clear extras)"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Frame fraExtras 
      Caption         =   "Extras"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin GameBase.ctlExtrasSelector ExtrasSelector 
         Height          =   3900
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6879
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmExtrasEditor"
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

Private Type GetExtrasType
    lngEXID As Long
    lngType As Long
    strName As String
    strPath As String
    strData As String
    strRunnableFile As String
End Type

Private m_blnActivatedOnce As Boolean
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim lngIdx As Long
    Dim a_udtExtra() As GetExtrasType
    Dim lngTotalExtras As Long
    Dim blnFound As Boolean
    Dim lngExtrasIdx As Long

    cmdCancel.Enabled = False
    cmdOK.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    g_blnApplyClearExtrasToAll = ((ExtrasSelector.ListCount = 0) And (chkAll.Enabled = True) And (chkAll.Value = vbChecked))
    If g_blnApplyClearExtrasToAll = False Then
        If ExtrasSelector.ListCount = 0 Then
            'delete all extras for current game
            g_dbGamebase.Execute "DELETE * FROM Extras WHERE GA_Id = " & g_udtGame.lngID
            g_dbGamebase.Execute "UPDATE Games SET Extras = False WHERE GA_Id = " & g_udtGame.lngID
        Else
            'grab extras
            lngTotalExtras = ExtrasSelector.ListCount
            ReDim a_udtExtra(lngTotalExtras - 1)
            For lngIdx = 0 To (ExtrasSelector.ListCount - 1)
                Call ExtrasSelector.GetItem( _
                        lngIdx, _
                        a_udtExtra(lngIdx).lngEXID, _
                        a_udtExtra(lngIdx).strName, _
                        a_udtExtra(lngIdx).strPath, _
                        a_udtExtra(lngIdx).lngType, _
                        a_udtExtra(lngIdx).strData, _
                        a_udtExtra(lngIdx).strRunnableFile _
                    )
            Next lngIdx
            'add/remove/edit the extras
            g_dbGamebase.Execute "UPDATE Games SET Extras = True WHERE GA_Id = " & g_udtGame.lngID
            Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Extras WHERE GA_Id = " & g_udtGame.lngID, dbOpenDynaset)
            With g_rsGames
                'remove extras not in the newly retrieved list
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        blnFound = False
                        For lngIdx = 0 To (lngTotalExtras - 1)
                            If !EX_Id = a_udtExtra(lngIdx).lngEXID Then
                                blnFound = True
                                Exit For
                            End If
                        Next lngIdx
                        If Not blnFound Then
                            .Delete
                        End If
                        .MoveNext
                    Loop
                End If
                'now update the table with the new data
                For lngIdx = 0 To (lngTotalExtras - 1)
                    .FindFirst "EX_Id = " & a_udtExtra(lngIdx).lngEXID
                    If .NoMatch Then
                        .AddNew
                    Else
                        .Edit
                    End If
                    !Name = a_udtExtra(lngIdx).strName
                    !Path = a_udtExtra(lngIdx).strPath
                    !Type = a_udtExtra(lngIdx).lngType
                    !Data = a_udtExtra(lngIdx).strData
                    !FileToRun = a_udtExtra(lngIdx).strRunnableFile
                    !GA_Id = g_udtGame.lngID
                    If a_udtExtra(lngIdx).lngType < 3 Then
                        !EA = (FindExtrasFile(lngExtrasIdx, a_udtExtra(lngIdx).strPath) = "")
                    Else
                        !EA = True
                    End If
                    !DisplayOrder = lngIdx
                    .Update
                Next lngIdx
                'all done
                .Close
            End With
            Set g_rsGames = Nothing
        End If
    End If

    'save flag and exit
    g_blnExtrasEditorOKPressed = True
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub ExtrasSelector_DisableAllCheckbox()

    chkAll.Enabled = False
    chkAll.Value = vbUnchecked

End Sub

Private Sub ExtrasSelector_EnableAllCheckbox()

    chkAll.Enabled = True

End Sub

Private Sub Form_Activate()
    
    If m_blnActivatedOnce = False Then
        m_blnActivatedOnce = True
        Call ShowTip(gc_lngTIPS_EDITEXTRAS)
    End If

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 5415    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5040     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set the ini text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Extras Manager", "Caption", "Extras for") & " " & g_udtGame.strName
        fraExtras.Caption = .GetINIMessage("Extras Manager", "Extras", "Extras")
        chkAll.Caption = .GetINIMessage("Extras Manager", "ApplyToAll", "&Apply to all games in current view (i.e. clear extras)")
    End With

    'fill the selector
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Extras WHERE GA_Id = " & g_udtGame.lngID & " ORDER BY DisplayOrder", dbOpenSnapshot)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Call ExtrasSelector.AddItem(!EX_Id, !Name, !Path, !Type, !Data, !FileToRun)
                .MoveNext
            Loop
            chkAll.Enabled = False
        Else
            chkAll.Enabled = True
        End If
        .Close
    End With
    Set g_rsGames = Nothing
    'set the

    'all loaded
    If ExtrasSelector.ListCount > 0 Then
        ExtrasSelector.SetListIndex 0
    End If
    g_blnExtrasEditorOKPressed = False
    g_blnApplyClearExtrasToAll = False
    m_blnActivatedOnce = False
    Screen.MousePointer = vbDefault

End Sub


