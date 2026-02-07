VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecCol"
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
'   GBSql:  Database Functions used in GB
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

'------------------------------------------------------------------------------
'This class sets up the second column information for the menus in gamebase.
'To move the second column fields around between their parent menus you can do
'it here.  Also you can add fields here to the menus and delete them.  If you
'want to add a new field that needs specific handling to display in the listview
'in gamebase, you'll have to add another value to the SecondColTypeEnum to
'handle it.
'
'Note: When adding second column selection menus and fields, you must also add
'the relevant text messages to the Text.ini file.  For each Menu ARRAY you
'create here, you must also create the corresponding menu in frmMain's menu
'editor.
'------------------------------------------------------------------------------

'different ways to handle second-column field
Private Enum SecondColTypeEnum
    e_SecColType_Normal
    e_SecColType_YesNo
    e_SecColType_Publisher
    e_SecColType_Year
    e_SecColType_Trainers
    e_SecColType_Players
    e_SecColType_Genres
    e_SecColType_Rating
    e_SecColType_PALNTSC
    e_SecColType_Control
    e_SecColType_Length
    e_SecColType_ReviewRating
End Enum
'each field has these properties
Private Type SecondColumnType
    strMenuText As String
    strColumnTitle As String
    strFields As String
    strMusicFields As String
    eType As SecondColTypeEnum
    blnGameOnlyTable As Boolean
End Type
'can have multiple sets of second-column fields in the menus in gamebase64
Private Type SecondColumnArrayType
    a_udtColumnItems() As SecondColumnType
    lngMaxColumns As Long
    strMenuText As String
End Type
'local storage for the column arrays
Private ma_udtArrays() As SecondColumnArrayType
'total second column (menu) arrays
Private Const mc_lngMAXMENUARRAYS As Long = 2

'menu-item splitter bar
Private Const mc_strMENUBAR As String = "-"

'reading from ini files
Private m_xobjFunc As GBFuncs.clsFunctions

Private m_strIniFile As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Sub SetIniFile(strIniFile As String)

    m_strIniFile = strIniFile

End Sub

Private Function GetINIMessage(strSection_in As String, strKey_in As String, strDefault_in As String) As String

    '------------------------------------------------------------------------
    'Reads a text item from the chosen GameBase Text.ini file
    '
    'Inputs:    strSection_in:      The INI file section
    '           strKey_in:          The INI file key
    '           strDefault_in:      The default string to return if not found
    '
    'Returns:   The found string in the INI file or the default string passed in!
    '------------------------------------------------------------------------

    Dim lngRes As Long                  'return Value
    Dim strReturned As String * 2096    'max length of ini file entry
    
    'get the required message
    lngRes = GetPrivateProfileString(strSection_in, strKey_in, strDefault_in, strReturned, 2096, m_strIniFile)
    GetINIMessage = Left$(strReturned, lngRes)

End Function

Public Sub GetSecondColField(ByVal lngMenuArrayIdx_in As Long, ByVal lngFieldArrayIdx_in As Long, strMenuText_out As String, strColumnTitle_out As String, strFields_out As String, strMusicFields_out As String, lngType_out As Long, blnGameOnlyTable_out As Boolean)

    '--------------------------------------------------------------------------------
    'Returns the data from the current array index of the GameInfo array.
    '
    'Inputs:    lngMenuArrayIdx_in      the index of the array to get info from
    '           lngFieldArrayIdx_in     the index of the field in the menu array
    '
    'Outputs:   strMenuText_out:        buffer to receive the Menu Text
    '           strColumnTitle_out:     buffer to receive the Column Title Text
    '           strFields_out:          buffer to receive the SQL fields
    '           strMusicFields_out:     buffer to receive the music SQL fields
    '           lngType_out:            buffer to receive the Column Type
    '           blnGameOnlyTable_out:   buffer to receive the GameOnly Table info
    '--------------------------------------------------------------------------------

    'copy data across
    With ma_udtArrays(lngMenuArrayIdx_in).a_udtColumnItems(lngFieldArrayIdx_in)
        strMenuText_out = .strMenuText
        strColumnTitle_out = .strColumnTitle
        strFields_out = .strFields
        strMusicFields_out = .strMusicFields
        lngType_out = .eType
        blnGameOnlyTable_out = .blnGameOnlyTable
    End With

End Sub
Public Property Get MaxArrays() As Long
    
    '-----------------------------------
    'return the ubound of the main array
    'i.e. how many s.c. menus we have...
    '-----------------------------------
    
    MaxArrays = mc_lngMAXMENUARRAYS
    
End Property
Public Property Get MaxFields(ByVal lngArrayIdx_in As Long) As Long
    
    '------------------------------------------------
    'return the ubound of the requested menu's fields
    '------------------------------------------------
    
    MaxFields = ma_udtArrays(lngArrayIdx_in).lngMaxColumns
    
End Property
Public Property Get MenuText(ByVal lngArrayIdx_in As Long) As String
    
    '------------------------------------------------
    'return the Menu Text for the second column menu
    '------------------------------------------------
    
    MenuText = ma_udtArrays(lngArrayIdx_in).strMenuText
    
End Property

Private Function RemoveAmpersand(strMenuText_in As String) As String

    RemoveAmpersand = Replace(strMenuText_in, "&&", "~~~~")
    RemoveAmpersand = Replace(RemoveAmpersand, "&", "")
    RemoveAmpersand = Replace(RemoveAmpersand, "~~~~", "&")

End Function

Public Sub InitColumns()

    'redim the main array of second-column menus
    ReDim ma_udtArrays(mc_lngMAXMENUARRAYS)

    'set the maximum ubounds for each array and redim it.
    'The max arrays redimensions here must not exceed mc_lngMAXMENUARRAYS
    With ma_udtArrays(0)
        .lngMaxColumns = 15
        ReDim .a_udtColumnItems(.lngMaxColumns)
    End With
    With ma_udtArrays(1)
        .lngMaxColumns = 20
        ReDim .a_udtColumnItems(.lngMaxColumns)
    End With
    With ma_udtArrays(2)
        .lngMaxColumns = 5
        ReDim .a_udtColumnItems(.lngMaxColumns)
    End With
        
    'Setup menu array 0 (GameInfo)
    With ma_udtArrays(0)
        .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_GameInfo", "Game Info")
        With .a_udtColumnItems(0)
            .eType = e_SecColType_Publisher
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Publisher", "Software House/Publisher")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Publishers.Publisher AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(1)
            .eType = e_SecColType_Year
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Year", "Year")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.YE_Id AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(2)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Developer", "Developer")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Developers.Developer AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(3)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Programmer", "Programmer")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Programmers.Programmer AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(4)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Musician", "Musician")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Musicians.Musician AS ColX"
            .strMusicFields = "Musicians.Musician AS ColX"
            .blnGameOnlyTable = False
        End With
        With .a_udtColumnItems(5)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Artist", "Artist")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Artists.Artist AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(6)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_License", "License")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Licenses.License AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(7)
            .eType = e_SecColType_Genres
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Genre", "Genre")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Genres.Genre AS ColX, PGenres.PG_Id as ColXExtra"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(8)
            .eType = e_SecColType_Players
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Players", "No. of Players")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.PlayersFrom AS ColX, Games.PlayersTo AS ColXExtra, Games.PlayersSim AS ColXExtra2"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(9)
            .eType = e_SecColType_Control
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Control", "Control")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.Control as ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(10)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Language", "Language")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Languages.Language AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(11)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Rarity", "Rarity")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Rarities.Rarity AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(12)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_WebLinkName", "More Info Name")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.WebLink_Name AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(13)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_WebLinkURL", "More Info URL")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.WebLink_URL AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(14)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Comment", "Comment")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.Comment AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
    End With

    'Setup the VersionInfo Column
    With ma_udtArrays(1)
        .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_VersionInfo", "Version Info")
        With .a_udtColumnItems(0)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_GameFilename", "Game Filename")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.Filename AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(1)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_CrackingGroup", "Cracking Group")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Crackers.Cracker AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(2)
            .eType = e_SecColType_Trainers
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Trainers", "No. of Trainers")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_Trainers AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(3)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Playable", "Playable")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_Playable AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(4)
            .eType = e_SecColType_Length
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_GameLength", "Game Length")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_Length AS ColX, Games.V_LengthType AS ColXExtra"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(5)
            .eType = e_SecColType_PALNTSC
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_NTSC", "PAL/NTSC")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_PalNTSC AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(6)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_VWebLinkName", "Download Name")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_WebLink_Name AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(7)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_VWebLinkURL", "Download URL")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_WebLink_URL AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(8)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Original", "Original")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_Original AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(9)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_LoadingScreen", "Loading Screen")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_LoadingScreen AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(10)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_TitleScreen", "Title Screen")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_TitleScreen AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With

        With .a_udtColumnItems(11)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_IncludedDocs", "Included Docs")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_IncludedDocs AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(12)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_HighScoreSaver", "High Score Saver")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_HighScoreSaver AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(13)
            .eType = e_SecColType_YesNo
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_TrueDriveEmu", "True Drive Emulation")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_TrueDriveEmu AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(14)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_VersionComment", "Version Comment")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.V_Comment AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(15)
            .eType = e_SecColType_Normal
            .strMenuText = mc_strMENUBAR
            .strColumnTitle = ""
            .strFields = ""
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(16)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_SidFilename", "Music Filename")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.SidFilename AS ColX"
            .strMusicFields = "Music.Filename AS ColX"
            .blnGameOnlyTable = False
        End With
        With .a_udtColumnItems(17)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_PictureFilename", "Picture Filename")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.ScrnShotFilename AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(18)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_MusicianNick", "Musician Nickname")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Musicians.Nick AS ColX"
            .strMusicFields = "Musicians.Nick AS ColX"
            .blnGameOnlyTable = False
        End With
        With .a_udtColumnItems(19)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_MusicianGroup", "Musician Group")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Musicians.Grp AS ColX"
            .strMusicFields = "Musicians.Grp AS ColX"
            .blnGameOnlyTable = False
        End With
        With .a_udtColumnItems(20)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_MusicianPhotoFilename", "Musician Photo Filename")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Musicians.Photo AS ColX"
            .strMusicFields = "Musicians.Photo AS ColX"
            .blnGameOnlyTable = False
        End With
    End With

    'Setup the StatsInfo Column
    With ma_udtArrays(2)
        .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Personal", "Personal Info")
        With .a_udtColumnItems(0)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_HighScore", "High Score")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.HighScore AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(1)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Difficulty", "Difficulty Level")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Difficulty.Difficulty AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(2)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_TimesPlayed", "Times Played")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.TimesPlayed AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(3)
            .eType = e_SecColType_Normal
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_LastPlayed", "Date Last Played")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.DateLastPlayed AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(4)
            .eType = e_SecColType_Rating
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_Rating", "Rating")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.Rating as ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
        With .a_udtColumnItems(5)
            .eType = e_SecColType_ReviewRating
            .strMenuText = GetINIMessage("MainMenus", "Menu_View_SC_ReviewRating", "Review Rating")
            .strColumnTitle = RemoveAmpersand(.strMenuText)
            .strFields = "Games.ReviewRating AS ColX"
            .strMusicFields = ""
            .blnGameOnlyTable = True
        End With
    End With


End Sub
Private Sub Class_Initialize()

    '-----------------------------------
    'Initialise the Second Column Arrays
    '-----------------------------------
    
    Set m_xobjFunc = New GBFuncs.clsFunctions
    Call InitColumns
    
End Sub


Private Sub Class_Terminate()
    
    Set m_xobjFunc = Nothing
    
End Sub


