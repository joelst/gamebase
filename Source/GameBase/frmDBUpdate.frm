VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DB Update"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmDBUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   735
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating your database..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmDBUpdate"
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

Private m_blnAlreadyActivated As Boolean

Private Type DBUpdateNewTableType
    strTableName As String
    strPrimaryKeyName As String
    strTextFieldName As String
    strPrimaryKeyIndex As String
    strTextFieldIndex As String
    strGamesIndex As String
    strRelation As String
End Type
Private Type DBUpdateLinkTableType
    strKeyName As String
    strTableName As String
    strFieldName As String
    lngFieldSize As Long
    strIndexName As String
End Type
Private Function Update_v20_to_v21() As String

    'Updates from database version 2.0 to version 2.1.
    'Returns empty string on success, or an error message on failure.

    Dim lngTemp As Long
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim rel As Relation
    Dim rsTemp As DAO.Recordset
    Dim fld1s As Field
    Dim fld2s As Field
    Dim fld3s As Field
    Dim fld4s As Field
    Dim fld1d As Field
    Dim fld2d As Field
    Dim fld3d As Field
    Dim fld4d As Field

    On Error GoTo Err_Handler

    '================================
    'Increase size of filename fields
    '================================

    lblStatus = "Increasing Games Table Field Sizes..."
    DoEvents
    
    'remove Games->Filename index
    g_dbGamebase.TableDefs("Games").Indexes.Delete "Filename"
    'copy Games table data into new table
    g_dbGamebase.Execute "SELECT GA_Id, Filename, FileToRun, ScrnShotFilename, SidFilename INTO Temp FROM Games"
    'remove fields from Games table
    g_dbGamebase.TableDefs("Games").Fields.Delete "Filename"
    g_dbGamebase.TableDefs("Games").Fields.Delete "FileToRun"
    g_dbGamebase.TableDefs("Games").Fields.Delete "ScrnShotFilename"
    g_dbGamebase.TableDefs("Games").Fields.Delete "SidFilename"
    'create the fields again as memo fields
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("Filename", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("FileToRun", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("ScrnShotFilename", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("SidFilename", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = Nothing
    'copy the data back again
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Filename, FileToRun, ScrnShotFilename, SidFilename FROM Games ORDER BY GA_Id", dbOpenDynaset)
    Set rsTemp = g_dbGamebase.OpenRecordset("SELECT Filename, FileToRun, ScrnShotFilename, SidFilename FROM Temp ORDER BY GA_Id", dbOpenSnapshot)
    With rsTemp
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount
            g_rsGames.MoveFirst
            Set fld1s = !Filename
            Set fld2s = !FileToRun
            Set fld3s = !ScrnShotFilename
            Set fld4s = !SidFilename
            Set fld1d = g_rsGames!Filename
            Set fld2d = g_rsGames!FileToRun
            Set fld3d = g_rsGames!ScrnShotFilename
            Set fld4d = g_rsGames!SidFilename
            Do Until .EOF
                g_rsGames.Edit
                fld1d = fld1s
                fld2d = fld2s
                fld3d = fld3s
                fld4d = fld4s
                g_rsGames.Update
                pbStatus.Value = .AbsolutePosition
                DoEvents
                .MoveNext
                g_rsGames.MoveNext
            Loop
            .Close
            Set fld1s = Nothing
            Set fld2s = Nothing
            Set fld3s = Nothing
            Set fld4s = Nothing
            Set fld1d = Nothing
            Set fld2d = Nothing
            Set fld3d = Nothing
            Set fld4d = Nothing
        End If
    End With
    g_rsGames.Close
    Set rsTemp = Nothing
    Set g_rsGames = Nothing
    g_dbGamebase.Execute "DROP TABLE Temp"

    lblStatus = "Increasing Music Table Field Sizes..."
    DoEvents
    
    'remove Music->Filename index
    g_dbGamebase.TableDefs("Music").Indexes.Delete "Filename"
    'copy Music table data into new table
    g_dbGamebase.Execute "SELECT GA_Id, Filename INTO Temp FROM Music"
    'remove field from Music table
    g_dbGamebase.TableDefs("Music").Fields.Delete "Filename"
    'create the field again as a memo field
    Set td = g_dbGamebase.TableDefs("Music")
    Set fld = td.CreateField("Filename", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = Nothing
    'copy the data back again
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Filename FROM Music ORDER BY GA_Id", dbOpenDynaset)
    Set rsTemp = g_dbGamebase.OpenRecordset("SELECT Filename FROM Temp ORDER BY GA_Id", dbOpenSnapshot)
    With rsTemp
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount
            g_rsGames.MoveFirst
            Set fld1s = !Filename
            Set fld1d = g_rsGames!Filename
            Do Until .EOF
                g_rsGames.Edit
                fld1d = fld1s
                g_rsGames.Update
                pbStatus.Value = .AbsolutePosition
                DoEvents
                .MoveNext
                g_rsGames.MoveNext
            Loop
            .Close
            Set fld1s = Nothing
            Set fld1d = Nothing
        End If
    End With
    g_rsGames.Close
    Set rsTemp = Nothing
    Set g_rsGames = Nothing
    g_dbGamebase.Execute "DROP TABLE Temp"

    '==================
    'update years table
    '==================

    lblStatus = "Updating 'Years' Table..."
    DoEvents

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Years ORDER BY YE_ID", dbOpenDynaset)
    With g_rsGames
        'verify the table is how we want it
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            If .RecordCount = 122 Then
                .MoveLast
                lngTemp = !YE_ID
                .MoveFirst
                If !YE_ID = 1 And lngTemp = 122 Then
                    'ok! Update the table!!
                    .MoveFirst
                    .Edit
                    !Year = 9991 '????
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9992 '19??
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9993 '197?
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9994 '198?
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9995 '199?
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9996 '20??
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9997 '200?
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9998 '201?
                    .Update
                    .MoveNext
                    .Edit
                    !Year = 9999 '202?
                    .Update
                    .MoveNext
                    pbStatus.Max = 2101
                    pbStatus.Min = 1970
                    pbStatus.Value = 1970
                    DoEvents
                    For lngTemp = 1970 To 2082
                        .Edit
                        !Year = lngTemp
                        .Update
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                        .MoveNext
                    Next lngTemp
                    For lngTemp = 2083 To 2100
                        .AddNew
                        !Year = lngTemp
                        .Update
                        pbStatus.Value = pbStatus.Value + 1
                        DoEvents
                    Next lngTemp
                    .Close
                Else
                    'no recs
                    .Close
                    Set g_rsGames = Nothing
                    Update_v20_to_v21 = "The years table was not the v2.0 standard. Unexpected YE_ID found"
                End If
            Else
                'no recs
                .Close
                Set g_rsGames = Nothing
                Update_v20_to_v21 = "The years table was not the v2.0 standard. Expected records = 122"
            End If
        Else
            'no recs
            .Close
            Set g_rsGames = Nothing
            Update_v20_to_v21 = "The years table was not the v2.0 standard. No records found"
        End If
    End With
    Set g_rsGames = Nothing

    '==============================
    'update games table years field
    '==============================

    lblStatus = "Updating 'Games->Year' field..."
    DoEvents
    g_dbGamebase.Execute "UPDATE Games SET YE_ID = YE_ID + 18 WHERE YE_ID > 1"

    '====================
    'add the crc-32 field
    '====================

    'add field
    lblStatus = "Adding 'Games->CRC' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("CRC", dbText, 32)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    'add index
    Set idx = td.CreateIndex("CRC_Index")
    Set fld = idx.CreateField("CRC")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Games SET CRC = """" WHERE CRC IS NULL"

    '======================
    'add the Filesize field
    '======================

    'add field
    lblStatus = "Adding 'Games->Filesize' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("Filesize", dbLong)
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Games SET Filesize = 0 WHERE Filesize IS NULL"

    '====================
    'add new extras table
    '====================

    lblStatus = "Adding 'Extras' table..."
    DoEvents
    'delete old extras table
    g_dbGamebase.TableDefs.Delete "Extras"
    'create new extras table
    Set td = g_dbGamebase.CreateTableDef("Extras")
    'add EX_Id field
    Set fld = td.CreateField("EX_Id", dbLong)
    fld.Attributes = dbAutoIncrField
    td.Fields.Append fld
    Set fld = Nothing
    'add GA_Id field
    Set fld = td.CreateField("GA_Id", dbLong)
    fld.Required = True
    td.Fields.Append fld
    Set fld = Nothing
    'add DisplayOrder field
    Set fld = td.CreateField("DisplayOrder", dbLong)
    fld.Required = True
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    'add Type field
    Set fld = td.CreateField("Type", dbLong)
    fld.Required = True
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    'add Name field
    Set fld = td.CreateField("Name", dbText, 255)
    fld.Required = True
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    'add Path field
    Set fld = td.CreateField("Path", dbMemo)
    fld.Required = True
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    'add EA field
    Set fld = td.CreateField("EA", dbBoolean)
    fld.DefaultValue = False
    td.Fields.Append fld
    Set fld = Nothing
    'add Data field
    Set fld = td.CreateField("Data", dbMemo)
    fld.Required = False
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    'add the EX_ID index
    Set idx = td.CreateIndex("EX_Id_Index")
    Set fld = idx.CreateField("EX_Id")
    idx.Unique = True
    idx.Primary = True
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    'add GA_Id index
    Set idx = td.CreateIndex("GA_Id_Index")
    Set fld = idx.CreateField("GA_Id")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    'add DisplayOrder index
    Set idx = td.CreateIndex("DisplayOrder_Index")
    Set fld = idx.CreateField("DisplayOrder")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    'add the table to the database
    g_dbGamebase.TableDefs.Append td
    Set td = Nothing
    'set the Games->Extras field to False
    g_dbGamebase.Execute "UPDATE Games SET Extras = False WHERE (1)"

    '========================
    'set the database to v2.1
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 1
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v20_to_v21 = ""
    Exit Function

Err_Handler:
    Update_v20_to_v21 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function Update_v23_to_v24() As String

    '========================
    'set the database to v2.4
    '========================

    On Error GoTo Err_Handler

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 4
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v23_to_v24 = ""
    Exit Function

Err_Handler:
    Update_v23_to_v24 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"


End Function

Private Function Update_v24_to_v25() As String

    Dim td As TableDef
    Dim fld As Field

    On Error GoTo Err_Handler

    '=========================================================================
    'Add "FirstLoadMessage" and "FirstLoadGemusAsk" fields to the Config table
    '=========================================================================

    'add fields
    lblStatus = "Adding 'Config->FirstLoadMessage' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Config")
    Set fld = td.CreateField("FirstLoadMessage", dbMemo)
    fld.DefaultValue = ""
    fld.AllowZeroLength = True
    td.Fields.Append fld
    Set fld = Nothing
    Set fld = td.CreateField("FirstLoadGemusAsk", dbLong)
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing

    '========================================================================
    'Add "V_GameLengthType" field to the Games table
    '========================================================================

    'add field
    lblStatus = "Adding 'Games->V_LengthType' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_LengthType", dbLong)
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Games SET V_LengthType = 1 WHERE V_Length < 0"
    g_dbGamebase.Execute "UPDATE Games SET V_LengthType = 0 WHERE V_Length > -1"
    g_dbGamebase.Execute "UPDATE Games SET V_Length = -V_Length WHERE V_Length < 0"
    
    '=======================================================
    'set the database to v2.5 and init the FirstLoad* fields
    '=======================================================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 5
        !FirstLoadMessage = ""
        !FirstLoadGemusAsk = 0
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v24_to_v25 = ""
    Exit Function

Err_Handler:
    Update_v24_to_v25 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function Update_v25_to_v26() As String

    On Error GoTo Err_Handler
 
    Dim lngTemp As Long
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim rel As Relation
    Dim rsTemp As DAO.Recordset
    Dim a_tables() As DBUpdateNewTableType
    Dim lngTableIdx As Long
    Dim fld1s As Field
    Dim fld2s As Field
    Dim fld3s As Field
    Dim fld4s As Field
    Dim fld1d As Field
    Dim fld2d As Field
    Dim fld3d As Field
    Dim fld4d As Field
    Dim a_link_tables() As DBUpdateLinkTableType
    Dim a_sec0(8) As Long
    Dim a_sec1(11) As Long
    Dim strTemp As String

    '==============================================
    'Second Column Re-Indexing for multiple columns
    '==============================================
    
    lblStatus = "Re-Indexing Second Column data for Multiple columns..."
    DoEvents
    
    'game info
    a_sec0(0) = 0
    a_sec0(1) = 1
    a_sec0(2) = 3
    a_sec0(3) = 4
    a_sec0(4) = 7
    a_sec0(5) = 8
    a_sec0(6) = 9
    a_sec0(7) = 10
    a_sec0(8) = 14

    'version info
    a_sec1(0) = 0
    a_sec1(1) = 1
    a_sec1(2) = 2
    a_sec1(3) = 4
    a_sec1(4) = 13
    a_sec1(5) = 12
    a_sec1(6) = 5
    a_sec1(7) = 11
    a_sec1(8) = 9
    a_sec1(9) = 14
    a_sec1(10) = 15
    a_sec1(11) = 16

    'create ExtraColumns field
    Set td = g_dbGamebase.TableDefs("ViewData")
    Set fld = td.CreateField("ExtraColumns", dbText, 255)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = Nothing

    'read in all views
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ID, SecondColumnMenu, SecondColumn, ExtraColumns FROM ViewData", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                strTemp = ""
                If !SecondColumnMenu = 0 Then
                    'game info
                    If !SecondColumn >= 0 And !SecondColumn < 9 Then
                        strTemp = !SecondColumnMenu & "-" & a_sec0(!SecondColumn)
                    End If
                ElseIf !SecondColumnMenu = 1 Then
                    'version info
                    If !SecondColumn >= 0 And !SecondColumn < 12 Then
                        strTemp = !SecondColumnMenu & "-" & a_sec1(!SecondColumn)
                    End If
                ElseIf !SecondColumnMenu = 2 Then
                    'personal info (no re-indexing needed)
                    If !SecondColumn >= 0 And !SecondColumn < 5 Then
                        strTemp = !SecondColumnMenu & "-" & !SecondColumn
                    End If
                End If
                .Edit
                !ExtraColumns = strTemp
                .Update
                .MoveNext
            Loop
            .Close
        End If
    End With
    Set g_rsGames = Nothing

    'remove SecondColumnMenu, SecondColumn, Col1Width and Col2Width fields
    g_dbGamebase.TableDefs("ViewData").Fields.Delete "SecondColumnMenu"
    g_dbGamebase.TableDefs("ViewData").Fields.Delete "SecondColumn"
    g_dbGamebase.TableDefs("ViewData").Fields.Delete "Col1Width"
    g_dbGamebase.TableDefs("ViewData").Fields.Delete "Col2Width"

    '==================================
    'Increase field size in link tables
    '==================================

    ReDim a_link_tables(11)
    With a_link_tables(0)
        .strTableName = "Crackers"
        .strKeyName = "CR_Id"
        .strFieldName = "Cracker"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(1)
        .strTableName = "Difficulty"
        .strKeyName = "DI_Id"
        .strFieldName = "Difficulty"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(2)
        .strTableName = "Languages"
        .strKeyName = "LA_Id"
        .strFieldName = "Language"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(3)
        .strTableName = "Musicians"
        .strKeyName = "MU_Id"
        .strFieldName = "Musician"
        .lngFieldSize = 255
        .strIndexName = "Name"
    End With
    With a_link_tables(4)
        .strTableName = "Musicians"
        .strKeyName = "MU_Id"
        .strFieldName = "Grp"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(5)
        .strTableName = "Musicians"
        .strKeyName = "MU_Id"
        .strFieldName = "Nick"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(6)
        .strTableName = "Programmers"
        .strKeyName = "PR_Id"
        .strFieldName = "Programmer"
        .lngFieldSize = 255
        .strIndexName = "Programmer"
    End With
    With a_link_tables(7)
        .strTableName = "Publishers"
        .strKeyName = "PU_Id"
        .strFieldName = "Publisher"
        .lngFieldSize = 255
        .strIndexName = "Publisher"
    End With
    With a_link_tables(8)
        .strTableName = "Genres"
        .strKeyName = "GE_Id"
        .strFieldName = "Genre"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(9)
        .strTableName = "PGenres"
        .strKeyName = "PG_Id"
        .strFieldName = "ParentGenre"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(10)
        .strTableName = "Games"
        .strKeyName = "GA_Id"
        .strFieldName = "Comment"
        .lngFieldSize = 255
        .strIndexName = ""
    End With
    With a_link_tables(11)
        .strTableName = "Games"
        .strKeyName = "GA_Id"
        .strFieldName = "V_Comment"
        .lngFieldSize = 255
        .strIndexName = ""
    End With

    For lngTableIdx = 0 To 11

        lblStatus = "Increasing Field Size of " & _
            a_link_tables(lngTableIdx).strTableName & ":" & a_link_tables(lngTableIdx).strFieldName & _
            " to " & a_link_tables(lngTableIdx).lngFieldSize & " ..."
        DoEvents

        'copy Games table data into new table
        g_dbGamebase.Execute "SELECT " & _
            a_link_tables(lngTableIdx).strKeyName & ", " & _
            a_link_tables(lngTableIdx).strFieldName & _
            " INTO TempTable FROM " & a_link_tables(lngTableIdx).strTableName

        'remove index on field?
        If a_link_tables(lngTableIdx).strIndexName <> "" Then
            g_dbGamebase.TableDefs(a_link_tables(lngTableIdx).strTableName).Indexes.Delete a_link_tables(lngTableIdx).strIndexName
        End If

        'remove field from table
        g_dbGamebase.TableDefs(a_link_tables(lngTableIdx).strTableName).Fields.Delete a_link_tables(lngTableIdx).strFieldName

        'create the field again to the specified size
        Set td = g_dbGamebase.TableDefs(a_link_tables(lngTableIdx).strTableName)
        Set fld = td.CreateField(a_link_tables(lngTableIdx).strFieldName, dbText, a_link_tables(lngTableIdx).lngFieldSize)
        fld.Required = False
        fld.DefaultValue = ""
        fld.AllowZeroLength = True
        td.Fields.Append fld
        Set fld = Nothing
        Set idx = td.CreateIndex(a_link_tables(lngTableIdx).strFieldName & "_Index")
        Set fld = idx.CreateField(a_link_tables(lngTableIdx).strFieldName)
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set fld = Nothing
        Set idx = Nothing
    
        'copy the data back again
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT " & a_link_tables(lngTableIdx).strFieldName & " FROM " & a_link_tables(lngTableIdx).strTableName & " ORDER BY " & a_link_tables(lngTableIdx).strKeyName, dbOpenDynaset)
        Set rsTemp = g_dbGamebase.OpenRecordset("SELECT " & a_link_tables(lngTableIdx).strFieldName & " FROM TempTable ORDER BY " & a_link_tables(lngTableIdx).strKeyName, dbOpenSnapshot)
        With rsTemp
            If Not .EOF Then
                .MoveLast
                .MoveFirst
                pbStatus.Min = 0
                pbStatus.Max = .RecordCount
                g_rsGames.MoveFirst
                Set fld1s = .Fields(a_link_tables(lngTableIdx).strFieldName)
                Set fld1d = g_rsGames.Fields(a_link_tables(lngTableIdx).strFieldName)
                Do Until .EOF
                    g_rsGames.Edit
                    fld1d = fld1s
                    g_rsGames.Update
                    pbStatus.Value = .AbsolutePosition
                    DoEvents
                    .MoveNext
                    g_rsGames.MoveNext
                Loop
                .Close
                Set fld1s = Nothing
                Set fld1d = Nothing
            End If
        End With
        g_rsGames.Close
        Set rsTemp = Nothing
        Set g_rsGames = Nothing
        g_dbGamebase.Execute "DROP TABLE TempTable"
    Next lngTableIdx

    '============================================
    'Convert game yes/no fields to yes/no/unknown
    '============================================

    lblStatus = "Converting Yes/No fields to Yes/No/Unknown..."
    DoEvents
    
    'copy Games table data into new table
    g_dbGamebase.Execute "SELECT GA_Id, V_LoadingScreen, V_HighScoreSaver, V_IncludedDocs, V_TrueDriveEmu INTO TempTable FROM Games"
    'remove fields from Games table
    g_dbGamebase.TableDefs("Games").Fields.Delete "V_LoadingScreen"
    g_dbGamebase.TableDefs("Games").Fields.Delete "V_HighScoreSaver"
    g_dbGamebase.TableDefs("Games").Fields.Delete "V_IncludedDocs"
    g_dbGamebase.TableDefs("Games").Fields.Delete "V_TrueDriveEmu"
    'create the fields again as number fields
    
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_LoadingScreen", dbLong)
    fld.Required = False
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_LoadingScreen_Index")
    Set fld = idx.CreateField("V_LoadingScreen")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_HighScoreSaver", dbLong)
    fld.Required = False
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_HighScoreSaver_Index")
    Set fld = idx.CreateField("V_HighScoreSaver")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing

    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_IncludedDocs", dbLong)
    fld.Required = False
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_IncludedDocs_Index")
    Set fld = idx.CreateField("V_IncludedDocs")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing

    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_TrueDriveEmu", dbLong)
    fld.Required = False
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_TrueDriveEmu_Index")
    Set fld = idx.CreateField("V_TrueDriveEmu")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing

    'copy the data back again
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT V_LoadingScreen, V_HighScoreSaver, V_IncludedDocs, V_TrueDriveEmu FROM Games ORDER BY GA_Id", dbOpenDynaset)
    Set rsTemp = g_dbGamebase.OpenRecordset("SELECT V_LoadingScreen, V_HighScoreSaver, V_IncludedDocs, V_TrueDriveEmu FROM TempTable ORDER BY GA_Id", dbOpenSnapshot)
    With rsTemp
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount
            g_rsGames.MoveFirst
            Set fld1s = !V_LoadingScreen
            Set fld2s = !V_HighScoreSaver
            Set fld3s = !V_IncludedDocs
            Set fld4s = !V_TrueDriveEmu
            Set fld1d = g_rsGames!V_LoadingScreen
            Set fld2d = g_rsGames!V_HighScoreSaver
            Set fld3d = g_rsGames!V_IncludedDocs
            Set fld4d = g_rsGames!V_TrueDriveEmu
            Do Until .EOF
                g_rsGames.Edit
                fld1d = Abs(fld1s)
                fld2d = Abs(fld2s)
                fld3d = Abs(fld3s)
                fld4d = Abs(fld4s)
                g_rsGames.Update
                pbStatus.Value = .AbsolutePosition
                DoEvents
                .MoveNext
                g_rsGames.MoveNext
            Loop
            .Close
            Set fld1s = Nothing
            Set fld2s = Nothing
            Set fld3s = Nothing
            Set fld4s = Nothing
            Set fld1d = Nothing
            Set fld2d = Nothing
            Set fld3d = Nothing
            Set fld4d = Nothing
        End If
    End With
    g_rsGames.Close
    Set rsTemp = Nothing
    Set g_rsGames = Nothing
    g_dbGamebase.Execute "DROP TABLE TempTable"

    '===================================================
    'New tables: Artists, Developers, Licenses, Rarities
    '===================================================

    ReDim a_tables(3)
    With a_tables(0)
        .strGamesIndex = "AR_IdIndex"
        .strPrimaryKeyIndex = "AR_Id_Index"
        .strPrimaryKeyName = "AR_Id"
        .strRelation = "AR_Id_Relation"
        .strTableName = "Artists"
        .strTextFieldIndex = "Artist_Index"
        .strTextFieldName = "Artist"
    End With
    With a_tables(1)
        .strGamesIndex = "DE_IdIndex"
        .strPrimaryKeyIndex = "DE_Id_Index"
        .strPrimaryKeyName = "DE_Id"
        .strRelation = "DE_Id_Relation"
        .strTableName = "Developers"
        .strTextFieldIndex = "Developer_Index"
        .strTextFieldName = "Developer"
    End With
    With a_tables(2)
        .strGamesIndex = "LI_IdIndex"
        .strPrimaryKeyIndex = "LI_Id_Index"
        .strPrimaryKeyName = "LI_Id"
        .strRelation = "LI_Id_Relation"
        .strTableName = "Licenses"
        .strTextFieldIndex = "License_Index"
        .strTextFieldName = "License"
    End With
    With a_tables(3)
        .strGamesIndex = "RA_IdIndex"
        .strPrimaryKeyIndex = "RA_Id_Index"
        .strPrimaryKeyName = "RA_Id"
        .strRelation = "RA_Id_Relation"
        .strTableName = "Rarities"
        .strTextFieldIndex = "Rarity_Index"
        .strTextFieldName = "Rarity"
    End With

    'add new tables and link fields
    For lngTableIdx = 0 To 3
        'create table
        lblStatus = "Adding '" & a_tables(lngTableIdx).strTableName & "' table..."
        DoEvents
        Set td = g_dbGamebase.CreateTableDef(a_tables(lngTableIdx).strTableName)
        'add XX_Id field
        Set fld = td.CreateField(a_tables(lngTableIdx).strPrimaryKeyName, dbLong)
        fld.Attributes = dbAutoIncrField
        td.Fields.Append fld
        Set fld = Nothing
        'add text field
        Set fld = td.CreateField(a_tables(lngTableIdx).strTextFieldName, dbText, 255)
        fld.AllowZeroLength = False
        td.Fields.Append fld
        Set fld = Nothing
        'add the XX_ID index
        Set idx = td.CreateIndex(a_tables(lngTableIdx).strPrimaryKeyIndex)
        Set fld = idx.CreateField(a_tables(lngTableIdx).strPrimaryKeyName)
        idx.Unique = True
        idx.Primary = True
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set fld = Nothing
        Set idx = Nothing
        'add the Artist index
        Set idx = td.CreateIndex(a_tables(lngTableIdx).strTextFieldIndex)
        Set fld = idx.CreateField(a_tables(lngTableIdx).strTextFieldName)
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set fld = Nothing
        Set idx = Nothing
        'add the table to the database
        g_dbGamebase.TableDefs.Append td
        Set td = Nothing
        'add the Unknown and None records
        Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT * FROM " & a_tables(lngTableIdx).strTableName, dbOpenDynaset)
        With g_rsGames
            .AddNew
            .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "(Unknown)"
            .Update
            If lngTableIdx = 0 Or lngTableIdx = 1 Then
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "(None)"
                .Update
            ElseIf lngTableIdx = 2 Then
                'licenses
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Charityware"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Commercial"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Coverdisk"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Game & Mini Mag"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Giftware"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Hidden Sub Game"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Interpreted"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Licenseware"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Loading Sub Game"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Open Source Conversion"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "PD/Freeware"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Promotional"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Reverse Engineered Modification"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "ShareWare"
                .Update
            ElseIf lngTableIdx = 3 Then
                'rarities
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Common as mud"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "One version common, other versions rare"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Rare"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Extremely Rare"
                .Update
                .AddNew
                .Fields(a_tables(lngTableIdx).strTextFieldName).Value = "Finished/reviewed but not found/released"
                .Update
            End If
            .Close
        End With
        Set g_rsGames = Nothing
    
        'add the Games->XX_Id field
        lblStatus = "Adding 'Games->" & a_tables(lngTableIdx).strPrimaryKeyName & "' field..."
        DoEvents
        Set td = g_dbGamebase.TableDefs("Games")
        Set fld = td.CreateField(a_tables(lngTableIdx).strPrimaryKeyName, dbLong)
        fld.DefaultValue = 1
        td.Fields.Append fld
        Set fld = Nothing
        'add index
        Set idx = td.CreateIndex(a_tables(lngTableIdx).strGamesIndex)
        Set fld = idx.CreateField(a_tables(lngTableIdx).strPrimaryKeyName)
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set fld = Nothing
        Set idx = Nothing
        Set td = Nothing
        'init field
        g_dbGamebase.Execute "UPDATE Games SET " & a_tables(lngTableIdx).strPrimaryKeyName & " = 1 WHERE (1)"
        'set the relationship: Games->XX_Id ----> [Table]->XX_Id
        Set rel = g_dbGamebase.CreateRelation(a_tables(lngTableIdx).strRelation, a_tables(lngTableIdx).strTableName, "Games")
        rel.Attributes = dbRelationDontEnforce
        rel.Fields.Append rel.CreateField(a_tables(lngTableIdx).strPrimaryKeyName)
        rel.Fields(a_tables(lngTableIdx).strPrimaryKeyName).ForeignName = a_tables(lngTableIdx).strPrimaryKeyName
        g_dbGamebase.Relations.Append rel
        Set rel = Nothing
    Next lngTableIdx

    '======================
    'new games table fields
    '======================

    'WebLink_Name
    lblStatus = "Adding 'Games->WebLink_Name' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("WebLink_Name", dbText, 255)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("WebLink_Name_Index")
    Set fld = idx.CreateField("WebLink_Name")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET WebLink_Name = """" WHERE WebLink_Name IS NULL"

    'WebLink_URL
    lblStatus = "Adding 'Games->WebLink_URL' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("WebLink_URL", dbMemo)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET WebLink_URL = """" WHERE WebLink_URL IS NULL"
    
    'V_WebLink_Name
    lblStatus = "Adding 'Games->V_WebLink_Name' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_WebLink_Name", dbText, 255)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_WebLink_Name_Index")
    Set fld = idx.CreateField("V_WebLink_Name")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET V_WebLink_Name = """" WHERE V_WebLink_Name IS NULL"

    'V_WebLink_URL
    lblStatus = "Adding 'Games->V_WebLink_URL' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_WebLink_URL", dbMemo)
    fld.AllowZeroLength = True
    fld.DefaultValue = """"""
    td.Fields.Append fld
    Set fld = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET V_WebLink_URL = """" WHERE V_WebLink_URL IS NULL"

    'V_TitleScreen
    lblStatus = "Adding 'Games->V_TitleScreen' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_TitleScreen", dbLong)
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_TitleScreen_Index")
    Set fld = idx.CreateField("V_TitleScreen")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET V_TitleScreen = 2 WHERE (1)"

    'V_Playable
    lblStatus = "Adding 'Games->V_Playable' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_Playable", dbLong)
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_Playable_Index")
    Set fld = idx.CreateField("V_Playable")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET V_Playable = 2 WHERE (1)"

    'V_Original
    lblStatus = "Adding 'Games->V_Original' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("V_Original", dbLong)
    fld.DefaultValue = 2
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("V_Original_Index")
    Set fld = idx.CreateField("V_Original")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET V_Original = 2 WHERE (1)"

    'CloneOf
    lblStatus = "Adding 'Games->CloneOf' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("CloneOf", dbLong)
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("CloneOf_Index")
    Set fld = idx.CreateField("CloneOf")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET CloneOf = 0 WHERE (1)"

    '========================
    'set the database to v2.6
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 6
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v25_to_v26 = ""
    Exit Function

Err_Handler:
    Update_v25_to_v26 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function Update_v26_to_v27() As String

    On Error GoTo Err_Handler
 
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index

    '==========
    'NEW FIELDS
    '==========

    'ReviewRating
    lblStatus = "Adding 'Games->ReviewRating' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("ReviewRating", dbLong)
    fld.DefaultValue = -1
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("ReviewRating_Index")
    Set fld = idx.CreateField("ReviewRating")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    g_dbGamebase.Execute "UPDATE Games SET ReviewRating = -1 WHERE (1)"

    '========================
    'set the database to v2.7
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 7
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v26_to_v27 = ""
    Exit Function

Err_Handler:
    Update_v26_to_v27 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"
End Function

Private Function Update_v27_to_v28() As String

    On Error GoTo Err_Handler
 
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim lngOrdinal As Long

    '==========
    'NEW FIELDS
    '==========

    'Config->DatabaseName
    lblStatus = "Adding 'Config->DatabaseName' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Config")
    Set fld = td.CreateField("DatabaseName", dbText, 255)
    fld.DefaultValue = """"""
    fld.AllowZeroLength = True
    fld.Required = False
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Config SET DatabaseName = """" WHERE (1)"

    'Config->GameBaseWindowTitle
    lblStatus = "Adding 'Config->GameBaseWindowTitle' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Config")
    Set fld = td.CreateField("GameBaseWindowTitle", dbText, 255)
    fld.DefaultValue = """"""
    fld.AllowZeroLength = True
    fld.Required = False
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Config SET GameBaseWindowTitle = """" WHERE (1)"

    'Extras->FileToRun
    lblStatus = "Adding 'Extras->FileToRun' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Extras")
    Set fld = td.CreateField("FileToRun", dbMemo)
    fld.DefaultValue = """"""
    fld.AllowZeroLength = True
    fld.Required = False
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Extras SET FileToRun = """" WHERE (1)"

    'ViewData->Ordinal
    lblStatus = "Adding 'ViewData->Ordinal' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("ViewData")
    Set fld = td.CreateField("Ordinal", dbLong)
    fld.DefaultValue = 0
    td.Fields.Append fld
    Set fld = Nothing
    Set idx = td.CreateIndex("Ordinal_Index")
    Set fld = idx.CreateField("Ordinal")
    idx.Unique = False
    idx.Fields.Append fld
    td.Indexes.Append idx
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    'init field
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT ID, Ordinal FROM ViewData ORDER BY ID", dbOpenDynaset)
    With g_rsGames
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            lngOrdinal = 0
            pbStatus.Min = 0
            pbStatus.Max = .RecordCount
            Do Until .EOF
                .Edit
                !Ordinal = lngOrdinal
                .Update
                pbStatus.Value = .AbsolutePosition
                DoEvents
                .MoveNext
                lngOrdinal = lngOrdinal + 1
            Loop
            .Close
        End If
    End With
    Set g_rsGames = Nothing

    '========================
    'set the database to v2.8
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 8
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v27_to_v28 = ""
    Exit Function

Err_Handler:
    Update_v27_to_v28 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function UpdateDatabase() As String

    'Updates the database to the current version.
    'Returns true if there is no error.

    'backup file
    Call CopyFile(g_strGameBaseDBPath, g_strGameBaseDBPath & ".bak", 0&)

    'v2.0 -> v2.1 changes
    If g_sngDBVersion < 2.1 Then
        UpdateDatabase = Update_v20_to_v21()
        g_sngDBVersion = 2.1
    End If
    
    'v2.1 -> v2.2 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.2 Then
            UpdateDatabase = Update_v21_to_v22()
            g_sngDBVersion = 2.2
        End If
    End If

    'v2.2 -> v2.3 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.3 Then
            UpdateDatabase = Update_v22_to_v23()
            g_sngDBVersion = 2.3
        End If
    End If

    'v2.3 -> v2.4 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.4 Then
            UpdateDatabase = Update_v23_to_v24()
            g_sngDBVersion = 2.4
        End If
    End If

    'v2.4 -> v2.5 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.5 Then
            UpdateDatabase = Update_v24_to_v25()
            g_sngDBVersion = 2.5
        End If
    End If

    'v2.5 -> v2.6 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.6 Then
            UpdateDatabase = Update_v25_to_v26()
            g_sngDBVersion = 2.6
        End If
    End If

    'v2.6 -> v2.7 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.7 Then
            UpdateDatabase = Update_v26_to_v27()
            g_sngDBVersion = 2.7
        End If
    End If

    'v2.7 -> v2.8 changes
    If Len(UpdateDatabase) = 0 Then
        If g_sngDBVersion < 2.8 Then
            UpdateDatabase = Update_v27_to_v28()
            g_sngDBVersion = 2.8
        End If
    End If

End Function

Private Function Update_v21_to_v22() As String

    'Updates from database version 2.1 to version 2.2.
    'Returns empty string on success, or an error message on failure.

    Dim lngTemp As Long
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim rel As Relation
    Dim rsTemp As DAO.Recordset
    Dim fld1s As Field
    Dim fld2s As Field
    Dim fld3s As Field
    Dim fld4s As Field
    Dim fld1d As Field
    Dim fld2d As Field
    Dim fld3d As Field
    Dim fld4d As Field

    On Error GoTo Err_Handler

    '===========================================================
    'Allow Publisher/Programmer/Musician fields to be duplicates
    '===========================================================

    lblStatus = "Changing indices to allow duplicates..."
    DoEvents

    With g_dbGamebase
        'first delete'em as we can't just change'em
        .TableDefs("Publishers").Indexes.Delete "Publisher"
        .TableDefs("Programmers").Indexes.Delete "Programmer"
        .TableDefs("Musicians").Indexes.Delete "Name"
        'then add'em
        Set td = .TableDefs("Publishers")
        Set idx = td.CreateIndex("Publisher")
        Set fld = idx.CreateField("Publisher")
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set td = Nothing
        Set td = .TableDefs("Musicians")
        Set idx = td.CreateIndex("Name")
        Set fld = idx.CreateField("Musician")
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set td = Nothing
        Set td = .TableDefs("Programmers")
        Set idx = td.CreateIndex("Programmer")
        Set fld = idx.CreateField("Programmer")
        idx.Unique = False
        idx.Fields.Append fld
        td.Indexes.Append idx
        Set td = Nothing
    End With

    '======================================
    'Add "Version" field to the Games table
    '======================================
    
    'add field
    lblStatus = "Adding 'Games->Version' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("Version", dbLong)
    fld.DefaultValue = 1
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Games SET Version = 1 WHERE (1)"

    '========================
    'set the database to v2.2
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 2
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v21_to_v22 = ""
    Exit Function

Err_Handler:
    Update_v21_to_v22 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Function Update_v22_to_v23() As String

    'Updates from database version 2.2 to version 2.3.
    'Returns empty string on success, or an error message on failure.

    Dim lngTemp As Long
    Dim td As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim rel As Relation
    Dim rsTemp As DAO.Recordset
    Dim a_strNotes() As String
    Dim lngTotalLines As Long
    Dim strNotes As String
    Dim strGemus As String
    Dim strLine As String
    Dim lngEqualsPos As Long
    Dim lngIdx As Long
    Dim lngLen As Long
    Dim strKey As String
    Dim strValue As String
    Dim blnIsPair As Boolean

    On Error GoTo Err_Handler

    '====================================
    'Add "Gemus" field to the Games table
    '====================================
    
    'add field
    lblStatus = "Adding 'Games->Gemus' field..."
    DoEvents
    Set td = g_dbGamebase.TableDefs("Games")
    Set fld = td.CreateField("Gemus", dbMemo)
    fld.DefaultValue = """"""
    fld.AllowZeroLength = True
    fld.Required = False
    td.Fields.Append fld
    Set fld = Nothing
    'init field
    g_dbGamebase.Execute "UPDATE Games SET Gemus = """" WHERE (1)"


    '====================================================================================
    'Copy all gemus key=value pairs from the Games->Notes field to the Games->Gemus field
    '====================================================================================
    
    'what does the user want to do?
    frmNotes2Gemus.Show vbModal, Me
    If frmNotes2Gemus.optGemus(0).Value = True Then
        lngTemp = 0
    ElseIf frmNotes2Gemus.optGemus(1).Value = True Then
        lngTemp = 1
    Else
        lngTemp = 2
    End If
    Unload frmNotes2Gemus
    If lngTemp = 2 Then
    
        'just move all notes to gemus as they only use the notes field for key=value pairs
        lblStatus = "Moving ""key=value"" pairs..."
        DoEvents
        g_dbGamebase.Execute "UPDATE Games SET Gemus = MemoText WHERE (1)"
        g_dbGamebase.Execute "UPDATE Games SET MemoText = """" WHERE (1)"

    ElseIf lngTemp = 1 Then
        'with each notes field, only move key=value pairs
        lblStatus = "Moving ""key=value"" pairs..."
        DoEvents
        Set rsTemp = g_dbGamebase.OpenRecordset("SELECT MemoText, Gemus FROM Games", dbOpenDynaset)
        With rsTemp
            If Not .EOF Then
                .MoveLast
                .MoveFirst
                pbStatus.Min = 0
                pbStatus.Max = .RecordCount
                Do Until .EOF
                    .Edit
                    If !MemoText <> "" Then
                        'get the new strings
                        a_strNotes = Split(!MemoText, vbCrLf)
                        lngTotalLines = UBound(a_strNotes)
                        strGemus = ""
                        strNotes = ""
                        For lngIdx = 0 To lngTotalLines - 1
                            strLine = Trim$(a_strNotes(lngIdx))
                            lngEqualsPos = InStr(strLine, "=")
                            If lngEqualsPos > 0 Then
                                If Left$(strLine, 1) <> "=" Then
                                    lngLen = Len(strLine)
                                    If lngEqualsPos < lngLen Then
                                        'split into key and value strings
                                        strKey = Trim$(Left$(strLine, lngEqualsPos - 1))
                                        strValue = Trim$(Right$(strLine, lngLen - lngEqualsPos))
                                        blnIsPair = True
                                        'is there an embedded space in the key?
                                        If (InStr(strKey, " ") > 0) Then
                                            blnIsPair = (MsgBox("Should I treat this as a Gemus ""key=value"" pair?" & vbCrLf & vbCrLf & strLine, vbQuestion Or vbYesNo Or vbDefaultButton2, "Ambiguous ""key=value"" pair found!") = vbYes)
                                        End If
                                        If blnIsPair Then
                                            'yay - we found one!
                                            strGemus = strGemus & a_strNotes(lngIdx) & vbCrLf
                                        Else
                                            'not a key=value pair
                                            strNotes = strNotes & a_strNotes(lngIdx) & vbCrLf
                                        End If
                                    Else
                                        'last char is the first found "=" so treat as not a key=value pair
                                        strNotes = strNotes & a_strNotes(lngIdx) & vbCrLf
                                    End If
                                Else
                                    'first char is "=" so treat as not a key=value pair
                                    strNotes = strNotes & a_strNotes(lngIdx) & vbCrLf
                                End If
                            Else
                                'no = char
                                strNotes = strNotes & a_strNotes(lngIdx) & vbCrLf
                            End If
                        Next lngIdx
                        'write them back to the table
                        !MemoText = strNotes
                        !Gemus = strGemus
                    End If
                    .Update
                    pbStatus.Value = .AbsolutePosition
                    DoEvents
                    .MoveNext
                Loop
                .Close
            End If
        End With
        Set rsTemp = Nothing
    Else
        'nothing needed here as the user doesn't have any key=value pairs in the notes field
    End If

    '========================
    'set the database to v2.3
    '========================

    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT MajorVersion, MinorVersion FROM Config", dbOpenDynaset)
    With g_rsGames
        .MoveFirst
        .Edit
        !MajorVersion = 2
        !MinorVersion = 3
        .Update
        .Close
    End With
    Set g_rsGames = Nothing

    '=======
    'all ok!
    '=======

    Update_v22_to_v23 = ""
    Exit Function

Err_Handler:
    Update_v22_to_v23 = Err.Description & "  (" & Err.Number & "," & Err.LastDllError & ")"

End Function
Private Sub Form_Activate()

    If m_blnAlreadyActivated = False Then
        m_blnAlreadyActivated = True
        g_strErrorUpdatingDatabase = UpdateDatabase()
        Unload Me
    End If

End Sub
Private Sub Form_Load()

    Const DesignScaleHeight As Single = 735     'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5175     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'assume an error until all completed successfully!
    g_strErrorUpdatingDatabase = "Unexpected Error!"

    'read in the ini file text
    With g_xobjFunc
        Caption = .GetINIMessage("Database Updater", "Caption", "Updating Database") & "..."
    End With

    'Continues in Form_Activate()
    m_blnAlreadyActivated = False

End Sub


