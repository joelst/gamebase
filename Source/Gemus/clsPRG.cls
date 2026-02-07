VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRG"
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

Private mFilesize As String
Private mLoadAddress As String
Private mSYSAddress As String

Public Function Load(ByVal strFilename As String) As String

    On Error GoTo Err_Handler

    Dim f As Integer
    Dim SYSAdd&
    Dim LoByte As Byte
    Dim HiByte As Byte
    Dim strMessFile As String
    
    f = FreeFile()
    Open strFilename For Binary As #f
    mFilesize = Format$(LOF(f) / 1024, "#####0.00")
    Get #f, , LoByte
    Get #f, , HiByte
    Close #f

    mLoadAddress = "$" & Format$(Hex$(HiByte), "00") & Format$(Hex$(LoByte), "00")
    SYSAdd = (HiByte * 256) + LoByte
    mSYSAddress = "(SYS " & Trim$(SYSAdd) & ")"

    Load = ""
    Exit Function
    
Err_Handler:
    strMessFile = App.Path
    If Right$(strMessFile, 1) <> "\" Then strMessFile = strMessFile & "\"
    strMessFile = strMessFile & "GemusText.ini"
    Load = GetINIMessageFromFile("Running Games", "Status_NonConform", "There was an error reading from the file or the file does not conform to the standard.", strMessFile)

End Function
Public Property Get SYSAddress() As String
    SYSAddress = mSYSAddress
End Property



Public Property Get FileSize() As String
    FileSize = mFilesize
End Property


Public Property Get LoadAddress() As String
    LoadAddress = mLoadAddress
End Property


