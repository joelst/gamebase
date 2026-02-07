VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRT"
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

'  Bytes:$0000-000F - 16-byte cartridge signature  "C64  CARTRIDGE"  (padded
'                     with space characters)
'         0010-0013 - File header  length  ($00000040,  in  high/low  format,
'                     calculated from offset $0000). Default value is $40
'         0014-0015 - Cartridge version (high/low, presently 01.00)
'         0016-0017 - Cartridge hardware type ($0000)
'                       0 - Normal cartridge
'                       1 - Action Replay
'                       2 - KCS Power Cartridge
'                       3 - Final Cartridge III
'                       4 - Simons Basic
'                       5 - Ocean type 1 (128K and 256K)*
'                       6 - Expert Cartridge
'                       7 - Fun Play
'                       8 - Super Games
'                       9 - Atomic Power
'                      10 - Epyx Fastload
'                      11 - Westermann
'                      12 - Rex
'                      13 - Final Cartridge I
'              0018 - Cartridge port EXROM line status
'                      0 - inactive
'                      1 - active
'              0019 - Cartridge port GAME line status
'                      0 - inactive
'                      1 - active
'         001A-001F - Reserved for future use
'         0020-003F - 32-byte cartridge  name  "CCSMON"  (padded  with  null
'                     characters)
'         0040-xxxx - Cartridge contents (called CHIP PACKETS, as there  can
'                     be more than one  per  CRT  file).  See  below  for  a
'                     breakdown of the CHIP format.
Private Type TAPType
    Signature(15) As Byte
    HeaderLen(3) As Byte
    Version(1) As Byte
    HardwareType(1) As Byte
    ExromStatus As Byte
    GameStatus As Byte
    Reserved(5) As Byte
    CartName(31) As Byte
End Type
Private Header As TAPType

Private mFilesize As String
Private mSignature As String
Private mVersion As String
Private mHardwareType As Long
Private mExromStatus As Boolean
Private mGameStatus As Boolean
Private mCartName As String

Public Function Load(ByVal strFilename As String) As String

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim a As Integer
    Dim strMessFile As String

    f = FreeFile()
    Open strFilename For Binary As #f
    mFilesize = Format$((LOF(f) / 1024), "0.00")
    Get #f, , Header
    Close #f

    mSignature = ""
    For a = 0 To 15
        mSignature = mSignature & Chr$(Header.Signature(a))
    Next a
    mSignature = TrimForDisplay(mSignature)
    mSignature = Trim$(mSignature)

    mVersion = "$" & Format$(Hex$(Header.Version(0)), "00") & "." & Format$(Hex$(Header.Version(1)), "00")
    
    mHardwareType = CLng((CLng(Header.HardwareType(0)) * 256&) + Header.HardwareType(1))
    
    mExromStatus = Not (Header.ExromStatus = 0)
    mGameStatus = Not (Header.GameStatus = 0)
    
    mCartName = ""
    For a = 0 To 31
        mCartName = mCartName & Chr$(Header.CartName(a))
    Next a
    mCartName = TrimForDisplay(mCartName)
    
    Load = ""
    Exit Function
    
Err_Handler:
    strMessFile = App.Path
    If Right$(strMessFile, 1) <> "\" Then strMessFile = strMessFile & "\"
    strMessFile = strMessFile & "GemusText.ini"
    Load = GetINIMessageFromFile("Running Games", "Status_NonConform", "There was an error reading from the file or the file does not conform to the standard.", strMessFile)

End Function

Private Function TrimForDisplay(ByVal strFullString As String) As String
    
    Dim Char As Long
    Dim Letter As String
    
    'strip nulls
    For Char = Len(strFullString) To 1 Step -1
        If Mid$(strFullString, Char, 1) <> Chr$(0) Then
            Exit For
        End If
    Next Char
    strFullString = Left$(strFullString, Char)
    
    'replace nulls
    TrimForDisplay = ""
    For Char = 1 To Len(strFullString)
        Letter = Mid$(strFullString, Char, 1)
        If Letter = Chr$(0) Then
            TrimForDisplay = TrimForDisplay & " "
        Else
            TrimForDisplay = TrimForDisplay & Letter
        End If
    Next Char
        
End Function

Public Property Get Version() As String
    Version = mVersion
End Property
Public Property Get Signature() As String
    Signature = mSignature
End Property
Public Property Get FileSize() As String
    FileSize = mFilesize
End Property
Public Property Get HardwareType() As Long
    HardwareType = mHardwareType
End Property
Public Property Get ExromLine() As Boolean
    ExromLine = mExromStatus
End Property
Public Property Get GameLine() As Boolean
    GameLine = mGameStatus
End Property
Public Property Get CartName() As String
    CartName = mCartName
End Property

