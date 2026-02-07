Attribute VB_Name = "modMain"
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
' HVSC2GB64 - Tool to update GB64 database with
' the latest version of the HVSC
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

Public Const gc_strDELIM As String = vbTab
Public g_strAPPVERSION As String
Public Const gc_strAPPNAME As String = "GB64's HVSC Update Wizard"
Public Const gc_strCOPYRIGHT As String = "(c) 2003 James Burrows"
Public Const gc_sngREQ_DB_VERSION As Single = 2.5
Public Const gc_strREQ_DB_VERSION As String = "2.5"
Public g_sngDBVersion As Single
Public g_lngDBMajorVersion As Long
Public g_lngDBMinorVersion As Long

Public Type MusicianPropsType
    strMusician As String
    lngMUID As Long
    strGroup As String
    strNick As String
    strPhoto As String
End Type
Public g_xobjFunc As GBFuncs.clsFunctions
Public g_xobjReg As GBFuncs.clsRegistry
Public g_strAppPathWithBackSlash As String
Public g_udtMusicianProps As MusicianPropsType
Public g_strPhotoPath As String
Public g_strDBFile As String
Public g_dbGB64 As DAO.Database
Public g_strHVSCPath As String
Public g_strHVSCPathWBS As String
Public g_strSidplayerExe As String
Public Const gc_lngEARLIESTYEAR As Long = 1970
Public Const gc_lngYEARLOOKUPMAX As Long = 140
Public ga_strYearLookup(1 To gc_lngYEARLOOKUPMAX) As String
Public g_strCache(999) As String
Public g_lngTotalCacheStrings As Long
Public g_lngCurrentCacheString As Long
Public g_lngCurrentCacheItem As Long
Public g_lngTotalSids As Long
Public g_lngLenHVSCPathWBS As Long
Public g_blnCacheError As Boolean
Public Type CacheArrayType
    lngID As Long
    strPathFile As String
    strPath As String
    strFile As String
    strName As String
    strAuthor As String
    strCopyright As String
    strSongs As String
    strDefSong As String
    blnFlag As Boolean
End Type
Public ga_udtCache() As CacheArrayType
Public g_lngCacheTotal As Long

Public Type FileSelectorType
    blnOKClicked As Boolean
    strSelectFile As String
    strReturnedFile As String
    strPattern As String
    strFolder As String
    strCaption As String
End Type
Public g_udtFileSelectorProps As FileSelectorType

Public Type MusicianPropertiesType
    blnOKClicked As Boolean
    blnGNPChanged As Boolean
    strMusician As String
    strGroup As String
    strNick As String
    strPhoto As String
    blnAdding As Boolean
    strFocus As String
End Type
Public g_udtMusicianProps2 As MusicianPropertiesType

Public Type RecordSelectorType
    blnOKClicked As Boolean
    lngID As Long
    strSQL As String
    strReturnedName As String
    strGroup As String
    strNick As String
    strPhoto As String
    blnGNPChanged As Boolean
End Type
Public g_udtRecordSelector As RecordSelectorType
