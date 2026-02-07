VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsD64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private mDiskName As String
Private mDiskID As String
Private mcolEntries As clsD64Entries

Private Type d64BAMType
    TrackOfFirstDirBlock As Byte
    SectorOfFirstDirBlock As Byte
    DiskDOSVersionType As Byte
    NotUsed As Byte
    BAMTrackEntries(139) As Byte
    DiskNameChars(15) As Byte
    PaddedWithAO(1) As Byte
    DiskID(1) As Byte
    UsuallyAO As Byte
    DOSType(1) As Byte
    FilledWithAO(3) As Byte
    UsedWithExtended40Tracks(84) As Byte
End Type
Private Type d64DIRType
    TrackOfNextDirBlock As Byte
    SectorOfNextDirBlock As Byte
    FileType As Byte
    TrackOfFirstBlockOfFile As Byte
    SectorOfFirstBlockOfFile As Byte
    FileNameChars(15) As Byte
    RELFileTrackLocation As Byte
    RELFileSectorLocation As Byte
    RELFileRecordLength As Byte
    GeosUnused(5) As Byte
    FileSizeInBlocksLO As Byte
    FileSizeInBlocksHI As Byte
End Type
Private D64DirEntry As d64DIRType
Private D64BAM As d64BAMType
'filetype constants
Private Const cSCRATCHED = 0
Private Const cSEQClosed = 1
Private Const cPRGClosed = 2
Private Const cUSRClosed = 3
Private Const cRELClosed = 4
Private Const cDEL = 128
Private Const cSEQ = 129
Private Const cPRG = 130
Private Const cUSR = 131
Private Const cREL = 132
Private Const cDELSaved = 160
Private Const cSEQSaved = 161
Private Const cPRGSaved = 162
Private Const cUSRSaved = 163
Private Const cRELSaved = 164
Private Const cDELLocked = 192
Private Const cSEQLocked = 193
Private Const cPRGLocked = 194
Private Const cUSRLocked = 195
Private Const cRELLocked = 196
Public Property Get DiskID() As String
    DiskID = mDiskID
End Property


Public Property Get DiskName() As String
    DiskName = mDiskName
End Property


Public Property Get Entries() As clsD64Entries
    Set Entries = mcolEntries
End Property

Private Function GetName(Chars() As Byte) As String
    
    Dim Char As String
    Dim a, Length As Integer
    Dim strWork As String
    
    'Convert byte array to string
    GetName = ""
    For a = 0 To 15
        Char = Chr(Chars(a))
        GetName = GetName & Char
    Next a
    
    'Trim trailing 0xA0 values
    strWork = RtrimA0(GetName)
    GetName = ""
    
    Length = Len(strWork)
    If Length > 0 Then
        Length = Length - 1
    Else
        GetName = ""
        Exit Function
    End If

    'Convert any other non-trailing 0xA0 values to spaces
    For a = 0 To Length
        Char = Mid$(strWork, a + 1, 1)
        If Asc(Char) = 160 Then
            Char = " "
        End If
        GetName = GetName & Char
    Next a
        
End Function

Private Sub DisplayDIREntry(Entry As d64DIRType)
    
    Dim EntryName As String
    Dim EntryType As String
    Dim EntrySize As String
    Dim OKToAdd As Boolean
    Dim Sizey As String
    
    OKToAdd = True
    
    Select Case D64DirEntry.FileType
        Case cSCRATCHED
            OKToAdd = False
        Case cDEL
            EntryType = "DEL"
        Case cSEQ
            EntryType = "SEQ"
        Case cPRG
            EntryType = "PRG"
        Case cUSR
            EntryType = "USR"
        Case cREL
            EntryType = "REL"
        Case cDELLocked
            EntryType = "DEL<"
        Case cSEQLocked
            EntryType = "SEQ<"
        Case cPRGLocked
            EntryType = "PRG<"
        Case cUSRLocked
            EntryType = "USR<"
        Case cRELLocked
            EntryType = "REL<"
        Case cSEQClosed
            EntryType = "*SEQ"
        Case cPRGClosed
            EntryType = "*PRG"
        Case cUSRClosed
            EntryType = "*USR"
        Case cRELClosed
            EntryType = "*REL"
        Case cDELSaved
            EntryType = "@0:DEL"
        Case cSEQSaved
            EntryType = "@0:SEQ"
        Case cPRGSaved
            EntryType = "@0:PRG"
        Case cUSRSaved
            EntryType = "@0:USR"
        Case cRELSaved
            EntryType = "@0:REL"
        Case Else
            EntryType = "?"
    End Select
    
    If OKToAdd = True Then
        Sizey = Space$(5)
        With D64DirEntry
            EntryName = GetName(.FileNameChars())
            EntrySize = Trim$(.FileSizeInBlocksHI * 256 + .FileSizeInBlocksLO)
        End With
        EntrySize = Left$(Sizey, 5 - Len(EntrySize)) & EntrySize
        mcolEntries.Add EntryName, EntrySize, EntryType
    End If
    
End Sub
Private Function GetNextBlock(ByVal Track As Byte, ByVal Block As Byte) As Long
    
    Dim WorkingLong As Long
    Select Case Track
        Case Is = 0
            WorkingLong = 0
        Case Is < 19
            WorkingLong = ((Track - 1) * 21&) * 256&
            WorkingLong = WorkingLong + (Block * 256&)
        Case Is < 26
            WorkingLong = 357& * 256&
            WorkingLong = WorkingLong + (((Track - 19) * 21&) * 256&)
            WorkingLong = WorkingLong + (Block * 256)
        Case Is < 32
            WorkingLong = 490& * 256&
            WorkingLong = WorkingLong + (((Track - 26) * 21&) * 256&)
            WorkingLong = WorkingLong + (Block * 256&)
        Case Is < 41
            WorkingLong = 598& * 256&
            WorkingLong = WorkingLong + (((Track - 32) * 21&) * 256&)
            WorkingLong = WorkingLong + (Block * 256&)
        Case Else
            WorkingLong = 0
    End Select
    WorkingLong = WorkingLong + 1 'get #x,[starts at 1 not 0],d64entry
    GetNextBlock = WorkingLong
    
End Function
Public Function Load(ByVal strFilename As String) As String

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim StartAddress As Long
    Dim Block As Integer
    Dim Entry As Integer
    Dim NextTrack As Long
    Dim NextSector As Long
    Dim tempDiskID As String
    Dim a As Integer
    Dim Char As String
    Dim strMessFile As String
    
    f = FreeFile()
    Open strFilename For Binary As #f
    Get #f, 91393, D64BAM
    With D64BAM
        mDiskName = GetName(.DiskNameChars())
        mDiskID = Chr$(.DiskID(0)) & Chr$(.DiskID(1)) & Chr$(.UsuallyAO) & Chr$(.DOSType(0)) & Chr$(.DOSType(1))
    End With
    tempDiskID = ""
    For a = 1 To 5
        Char = Mid$(mDiskID, a, 1)
        If Asc(Char) = 0 Or Asc(Char) = 160 Then
            tempDiskID = tempDiskID & " "
        Else
            tempDiskID = tempDiskID & Char
        End If
    Next a
    mDiskID = tempDiskID
    'now get the dir entries
    StartAddress = 91649
    For Block = 1 To 18 ' 18 blocks
        For Entry = 0 To 7 '8 entries per block
        Get #f, StartAddress + (Entry * 32), D64DirEntry
        DisplayDIREntry D64DirEntry
        If Entry = 0 Then
            With D64DirEntry
                NextTrack = .TrackOfNextDirBlock
                NextSector = .SectorOfNextDirBlock
            End With
        ElseIf Entry = 7 Then
            StartAddress = GetNextBlock(NextTrack, NextSector)
        End If
        Next Entry
        If NextTrack = 0 Then
            Exit For ' no more blocks
        End If
    Next Block

    
    Close #f
    Load = ""
    Exit Function
    
Err_Handler:
    strMessFile = App.Path
    If Right$(strMessFile, 1) <> "\" Then strMessFile = strMessFile & "\"
    strMessFile = strMessFile & "GemusText.ini"
    Load = GetINIMessageFromFile("Running Games", "Status_NonConform", "There was an error reading from the file or the file does not conform to the standard.", strMessFile)
  
End Function


Private Function RtrimA0(ByVal strName As String) As String
    
    Dim a As Integer
    
    RtrimA0 = ""
    For a = Len(strName) To 1 Step -1
        If Asc(Mid$(strName, a, 1)) <> 160 Then
            Exit For
        End If
    Next a
    RtrimA0 = Left$(strName, a)
    
    
End Function

Private Sub Class_Initialize()
    Set mcolEntries = New clsD64Entries
End Sub


Private Sub Class_Terminate()
    Set mcolEntries = Nothing
End Sub


