VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsD64Entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1999-2010 James Burrows
' Email: gamebase.frontend@leesti.com
' Homepage: http://www.bu22.com/
'----------------------------------------------
'     GEMUS - Gamebase EMUlator Script
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

Private mEntries As Collection

Public Function Add(ByVal Filename As String, ByVal FileSize As String, ByVal FileType As String) As clsD64Entry

    Dim NewEntry As New clsD64Entry
    
    With NewEntry
        .Filename = Filename
        .FileSizeInBlocks = FileSize
        .FileType = FileType
    End With

    mEntries.Add NewEntry
    Set Add = NewEntry
    Set NewEntry = Nothing

End Function


Public Function Count() As Integer
    Count = mEntries.Count
End Function


Public Function Entry(ByVal Index As Integer) As clsD64Entry
Attribute Entry.VB_UserMemId = 0
    Set Entry = mEntries.item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mEntries.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mEntries = New Collection
End Sub


Private Sub Class_Terminate()
    Set mEntries = Nothing
End Sub


