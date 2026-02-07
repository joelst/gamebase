Attribute VB_Name = "modLanguage"
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



'views engine
Public g_strViewNone As String
Public g_strCustomViewWindowTitle As String
Public g_strCVFilterWindowTitle As String
Public g_strCV_Contains As String
Public g_strCV_NotContains As String
Public g_strCV_StartsWith As String
Public g_strCV_EndsWith As String
Public g_strCV_Filled As String
Public g_strCV_Empty As String
Public g_strCV_Exists As String
Public g_strCV_Missing As String
Public g_strCV_IsAfter As String
Public g_strCV_IsBefore As String
Public g_strCV_MoreThan As String
Public g_strCV_LessThan As String
Public g_strCV_Equal As String
Public g_strCV_NotEqual As String
Public g_strCVFD_GameFilename As String
Public g_strCVFD_ScrnShotFilename As String
Public g_strCVFD_SidFilename As String
Public g_strCVFilter_Field As String
Public g_strCVFilter_Operator As String
Public g_strCVFilter_FilterData As String
Public ga_QuickFilterNames(11) As String

'general text
Public g_strSidNone As String
Public g_strNever As String
Public g_strPleaseWait As String
Public g_strUnknown As String
Public g_strNone As String
Public g_strNotEntered As String
Public g_strNoneSelected As String
Public g_strMarked As String

'general buttons
Public g_strButtonOKText As String
Public g_strButtonCancelText As String
Public g_strButtonCloseText As String
Public g_strButtonClearText As String
Public g_strButtonBrowseText As String
Public g_strButtonDefaultText As String
Public g_strButtonHelpText As String

'field names
Public g_strField_Name As String
Public g_strField_Year As String
Public g_strField_Comment As String
Public g_strField_GameFilename As String
Public g_strField_ScreenShotFilename As String
Public g_strField_SidFilename As String
Public g_strField_Musician As String
Public g_strField_Genre As String
Public g_strField_ParentGenre As String
Public g_strField_Publisher As String
Public g_strField_Difficulty As String
Public g_strField_Cracker As String
Public g_strField_Programmer As String
Public g_strField_Language As String
Public g_strField_DateLastPlayed As String
Public g_strField_TimesPlayed As String
Public g_strField_HighScore As String
Public g_strField_Favourite As String
Public g_strField_Extras As String
Public g_strField_ExtrasName As String
Public g_strField_Classic As String
Public g_strField_Rating As String
Public g_strField_LoadingScreen As String
Public g_strField_HighScoreSaver As String
Public g_strField_IncludedDocs As String
Public g_strField_PALNTSC As String
Public g_strField_TrueDriveEmulation As String
Public g_strField_GameLength As String
Public g_strField_GameLengthType As String
Public g_strField_Trainers As String
Public g_strField_NumPlayers As String
Public g_strField_Players_Simultaneous As String
Public g_strField_VersionComment As String
Public g_strField_GameVer As String
Public g_strField_Prequel As String
Public g_strField_Sequel As String
Public g_strField_Related As String
Public g_strField_CloneOf As String
Public g_strField_Notes As String
Public g_strField_MusicianGroup As String
Public g_strField_MusicianNick As String
Public g_strField_MusicianPhoto As String
Public g_strField_Control As String
Public g_strField_Adult As String
Public g_strField_License As String
Public g_strField_Artist As String
Public g_strField_Rarity As String
Public g_strField_Developer As String
Public g_strField_WebLinkName As String
Public g_strField_WebLinkURL As String
Public g_strField_VWebLinkName As String
Public g_strField_VWebLinkURL As String
Public g_strField_TitleScreen As String
Public g_strField_Playable As String
Public g_strField_Original As String
Public g_strField_ReviewRating As String
