Attribute VB_Name = "modGlobal"
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
'       Borg Assimilation Console (BAC)
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

'external objects
Public g_xobjStartupClock As GBFuncs.clsClock
Public g_xobjFunc As GBFuncs.clsFunctions
Public g_xobjRegistry As GBFuncs.clsRegistry

'databases
Public g_dbMaster As DAO.Database
Public g_dbDiff As DAO.Database

'database files
Public g_strMaster As String
Public g_strDiff As String
Public g_blnMasterOpen As Boolean
Public g_blnDiffOpen As Boolean

'constants
Public Const gc_sngREQ_DB_VERSION = 2.1

'winapi
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LVM_FIRST = &H1000
Public Const LVM_DELETEALLITEMS = LVM_FIRST + 9

Public Function ListItemExists(lvGames_in As ListView, strKey_in As String) As Boolean

    Dim LItem As ListItem

    On Error GoTo Err_Handler
    
    Set LItem = lvGames_in.ListItems(strKey_in)
    Set LItem = Nothing
    
    ListItemExists = True
    Exit Function

Err_Handler:
    ListItemExists = False

End Function


