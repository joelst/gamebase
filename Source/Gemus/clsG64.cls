VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsG64"
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

'  Bytes: $0000-0007: File signature "GCR-1541"
'               0008: G64 version (presently only $00 defined)
'               0009: Number of tracks in image (usually $54, decimal 84)
'          000A-000B: Size of each stored track in bytes (usually  7928,  or
'                     $1EF8 in LO/HI format.
Private Type G64Type
    Signature(7) As Byte
    Version As Byte
    NoTracks As Byte
    TrackSize(1) As Byte
End Type
Private Header As G64Type

Private mSignature As String
Private mVersion As String
Private mNoTracks As String
Private mTrackSize As String

Public Function Load(ByVal strFilename As String) As String

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim a As Integer
    Dim strMessFile As String
    
    f = FreeFile()
    Open strFilename For Binary As #f
    Get #f, , Header
    Close #f
    
    mSignature = ""
    For a = 0 To 7
        mSignature = mSignature & Chr$(Header.Signature(a))
    Next a
    mSignature = TrimForDisplay(mSignature)

    mVersion = "$" & Format$(Hex$(Header.Version), "00")

    mNoTracks = Trim$(Header.NoTracks)
    
    mTrackSize = Trim$(CLng(Header.TrackSize(1) * 256&) + Header.TrackSize(0))

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


Public Property Get NumTracks() As String
    NumTracks = mNoTracks
End Property

Public Property Get TrackSize() As String
    TrackSize = mTrackSize
End Property

Public Property Get Signature() As String
    Signature = mSignature
End Property



