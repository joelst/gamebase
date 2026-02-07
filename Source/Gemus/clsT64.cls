VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsT64"
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

Private mcolEntries As New clsT64Entries

Private Type T64Type
    Signature(31) As Byte
    Version(1) As Byte
    EntriesCount(1) As Byte
    EntriesUsed(1) As Byte
    NotUsed(1) As Byte
    TapeImageName(23) As Byte
End Type
Private Header As T64Type

Private Type EntryType
    EntryData(15) As Byte
    C64Filename(15) As Byte
End Type
Private DirEntry As EntryType

Private mSignature As String
Private mVersion As String
Private mEntriesCount As Long
Private mEntriesUsed As Long
Private mTapeImageName As String
Public Property Get Entries() As clsT64Entries
    Set Entries = mcolEntries
End Property


Public Property Get EntriesCount() As Long
    EntriesCount = mEntriesCount
End Property


Public Property Get EntriesUsed() As Long
    EntriesUsed = mEntriesUsed
End Property

Private Function Get2ByteAddress(ByVal HiByte As Byte, ByVal LoByte As Byte) As Long
    
    Dim Addr& 'long for storing the address
    Dim lHiByte&, lLoByte& 'longs for working the conversion
    
    lLoByte& = CLng(LoByte)
    lHiByte& = CLng(HiByte)
    
    Addr& = lLoByte& + lHiByte& * 256&
    Get2ByteAddress = Addr&
    
End Function
Private Sub Get2ByteValues(ByVal Address As Long, HiByte As Byte, LoByte As Byte)
    
    HiByte = Int(Address / 256)
    LoByte = Address - (256& * HiByte)

End Sub
Private Sub Get4ByteValues(ByVal Address As Double, Byte1Hi As Byte, Byte2Hi As Byte, Byte3Lo As Byte, Byte4Lo As Byte)

    Dim Addr# 'decimal for storing the address
    Dim dByte4Lo#, dByte3Lo#, dByte2Hi#, dByte1Hi# 'decimals for working the conversion
    
    Addr# = Address
    dByte1Hi# = Int(Addr# / 4294967296#)
    Addr# = Addr# - (4294967296# * dByte1Hi#)
    dByte2Hi# = Int(Addr# / 65536#)
    Addr# = Addr# - (65536# * dByte2Hi#)
    dByte3Lo# = Int(Addr# / 256#)
    dByte4Lo# = Addr# - (256# * dByte3Lo#)
    
    Byte1Hi = dByte1Hi#
    Byte2Hi = dByte2Hi#
    Byte3Lo = dByte3Lo#
    Byte4Lo = dByte4Lo#
    
End Sub
Private Function Get4ByteAddress(ByVal Byte1Hi As Byte, ByVal Byte2Hi As Byte, ByVal Byte3Lo As Byte, ByVal Byte4Lo As Byte) As Double
    
    
    Dim Addr# 'decimal for storing the address
    Dim dByte4Lo#, dByte3Lo#, dByte2Hi#, dByte1Hi# 'decimals for working the conversion
    
    dByte4Lo# = CDec(Byte4Lo)
    dByte3Lo# = CDec(Byte3Lo)
    dByte2Hi# = CDec(Byte2Hi)
    dByte1Hi# = CDec(Byte1Hi)

    Addr# = dByte4Lo# + (dByte3Lo# * 256#) + (dByte2Hi# * 65536#) + (dByte1Hi# * 4294967296#)
    Get4ByteAddress = Addr#
    
End Function
Public Function Load(ByVal strFilename As String) As String

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim a As Integer
    Dim b As Integer
    Dim TempC64Filename As String
    Dim strMessFile As String
    
    f = FreeFile()
    Open strFilename For Binary As #f
    Get #f, , Header
    
    mSignature = ""
    For a = 0 To 31
        mSignature = mSignature & Chr$(Header.Signature(a))
    Next a
    mSignature = Trim(mSignature)
    
    mVersion = "$" & Format$(Hex$(Get2ByteAddress(Header.Version(1), Header.Version(0))), "0000")
    mEntriesCount = Get2ByteAddress(Header.EntriesCount(1), Header.EntriesCount(0))
    mEntriesUsed = Get2ByteAddress(Header.EntriesUsed(1), Header.EntriesUsed(0))

    mTapeImageName = ""
    For a = 0 To 23
        mTapeImageName = mTapeImageName & Chr$(Header.TapeImageName(a))
    Next a
    mTapeImageName = Trim(mTapeImageName)
    
    For a = 1 To mEntriesUsed
        Get #f, , DirEntry
        TempC64Filename = ""
        For b = 0 To 15
            TempC64Filename = TempC64Filename & Chr$(DirEntry.C64Filename(b))
        Next b
        mcolEntries.Add Trim(TempC64Filename)
    Next a
    
    Close #f
    Load = ""
    Exit Function
    
Err_Handler:
    strMessFile = App.Path
    If Right$(strMessFile, 1) <> "\" Then strMessFile = strMessFile & "\"
    strMessFile = strMessFile & "GemusText.ini"
    Load = GetINIMessageFromFile("Running Games", "Status_NonConform", "There was an error reading from the file or the file does not conform to the standard.", strMessFile)

End Function

Private Function TrimCharFromRight(ByVal strFullString As String, ByVal strTrimChar As String) As String
    
    Dim Char As Long
    
    For Char = Len(strFullString) To 1 Step -1
        If Mid$(strFullString, Char, 1) <> strTrimChar Then
            Exit For
        End If
    Next Char
    TrimCharFromRight = Left$(strFullString, Char)
    
End Function
Public Property Get Signature() As String
    Signature = mSignature
End Property


Public Property Get TapeImageName() As String
    TapeImageName = mTapeImageName
End Property

Public Property Get Version() As String
    Version = mVersion
End Property


Private Sub Class_Terminate()
    Set mcolEntries = Nothing
End Sub


