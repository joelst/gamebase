Attribute VB_Name = "modGlobals"
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

'set this to True when running GB from VB, False when compiling for the public
Public Const g_blnInVBIDE As Boolean = False
Public Const gc_strAPP_PATH_WBS_WHEN_IN_VBIDE = "C:\GameBase\Build\"

Public Const g_blnBETA As Boolean = False
Public Const gc_strRELEASEDATE As String = "23 November 2010"


Public g_strAppPathWithBackslash As String
Public g_strMessagesFile As String
Public g_strGameBasePath As String
Public g_strGameBasePathWBS As String
Public g_strGameBaseName As String
Public g_strGameBaseDBPath As String
Public g_strPathFile As String
Public g_strEmuFile As String
Public g_strMusicFile As String
Public g_strSettingsFile As String
Public g_strGameBaseDotINI As String
Public g_blnAbort As Boolean

Public g_blnEasterEgg As Boolean
Public g_blnIsWavOnStart As Boolean
Public g_blnIsFirstItemDisplayed As Boolean
Public g_blnLoadImageExtrasAsScreenshots As Boolean

Public g_blnResizingWindow As Boolean
Public g_blnDebugMode As Boolean
Public g_blnSubclassed As Boolean
Public g_blnCatSubclassed As Boolean
Public g_blnPictureSelectorSubclassed As Boolean
Public g_blnGameSelectorSubclassed As Boolean
Public g_blnExtrasSubclassed As Boolean

Public g_xobjFunc As GBFuncs.clsFunctions
Public g_xobjRegistry As GBFuncs.clsRegistry
Public g_xobjClock As GBFuncs.clsClock
Public g_xobjGEMUS As Gemus.clsGemus
Public g_xobjStartupClock As GBFuncs.clsClock

Public g_xobjStartupWavClock As GBFuncs.clsClock
Public g_lngStartupWavSecondsToWait As Long
Public g_lngStartupSplashSecondsToWait As Long

'Public g_xobjPNG As cPNG  'removed

'Declare external C functions contained in pnglib.dll
Declare Function png_load Lib "pnglib.dll" (ByVal Filename As String) As Long
Declare Function png_draw Lib "pnglib.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function png_width Lib "pnglib.dll" () As Long
Declare Function png_height Lib "pnglib.dll" () As Long

Public g_dbGamebase As Database
Public g_rsGames As Recordset
Public g_strAPPVERSION As String
Public g_strWINDOWCAPTION As String
Public g_strDatabaseName As String

Public Const gc_strAPPNAME As String = "GameBase"
Public Const gc_strAUTHOR As String = "James Burrows"

Public Const gc_sngREQ_DB_VERSION As Single = 2.8
Public g_sngDBVersion As Single
Public g_lngDBMajorVersion As Long
Public g_lngDBMinorVersion As Long
Public g_strErrorUpdatingDatabase As String
Public g_lngOfficialDBUpdate As Long

Public Const gc_lngINVALIDWINDOWSIZE As Long = 100
Public Const gc_lngINVALIDWINDOWPOS As Long = -1
Public Const gc_lngNOLISTINDEX As Long = -1

Public Const gc_lngSCRNSHOTWIDTH As Long = 320
Public Const gc_lngSCRNSHOTHEIGHT As Long = 200
Public Const gc_lngSCRNSHOTWIDTHx2 As Long = 640
Public Const gc_lngSCRNSHOTHEIGHTx2 As Long = 400
Public Const gc_lngSCRNSHOTWIDTHHALF As Long = 160
Public Const gc_lngSCRNSHOTHEIGHTHALF As Long = 100

Public Const gc_lngEARLIESTYEAR As Long = 1970
Public Const gc_lngYEARLOOKUPMAX As Long = 140
Public Const gc_lngGAMENAMELOOKUPMAX As Long = 12
Public Const gc_lngRATINGLOOKUPMAX As Long = 6
Public Const gc_lngCONTROLLOOKUPMAX As Long = 8
Public Const gc_lngMAXMEMOTEXTLENGTH As Long = 65526
Public Const gc_lngMAXSCRIPTLENGTH As Long = 65526
Public Const gc_lngPLAYERSUNKNOWN As Long = -1
Public Const gc_lngTRAINERSUNKNOWN As Long = -1
Public Const gc_strEXTRASDELIM = "|~|"

Public g_lngFullSizePictureWidthPixels As Long
Public g_lngFullSizePictureHeightPixels As Long
Public g_lngSmallPictureWidthPixels As Long
Public g_lngSmallPictureHeightPixels As Long
Public g_lngZoomedPictureWidthPixels As Long
Public g_lngZoomedPictureHeightPixels As Long

Public g_sngScrnShotExtraTwipsX As Single
Public g_sngScrnShotExtraTwipsY As Single
Public g_lngScrnShotExtraPixelsX As Long
Public g_lngScrnShotExtraPixelsY As Long

Public g_sngCopyrightGap As Single

'key/value setting

Public Enum UpdateKVEnum
    e_Save_Current = 0
    e_Save_All = 1
    e_Add_All_Overwrite = 2
    e_Add_All_Leave = 3
    e_Remove_All_Matching_Keys = 4
    e_Remove_All_Matching_Exactly = 5
End Enum
Public g_UpdateKVTypeEnum As UpdateKVEnum
Public g_strUpdateKVText As String

'sid selector
Public g_strSidSelectorFilename As String
Public g_strSidSelectorSidName As String
Public g_blnApplySidFileToAll As Boolean
Public g_blnSidSelectorNewItem As Boolean

'game selector
Public g_strGameSelectorFilename As String
Public g_strGameSelectorIndex As Long
Public g_strGameSelectorDefaultFile As String
Public g_strGameSelectorImageName As String
Public g_blnGameSelectorNewGame As Boolean
Public g_lngGameSelectorFileSize As Long
Public g_strGameSelectorCRC32 As String
Public g_blnGameSelectorOKPressed As Boolean
Public g_lngGameSelectorVersion As Long
Public g_blnGameSelectorApplyAllVersion As Boolean
Public g_blnApplyGameFileToAll As Boolean

'screenshot selector
Public g_strScrnShotSelectorFilename As String
Public g_strScrnShotSelectorOpeningPath As String
Public g_strScrnShotSelectorDBFilename As String
Public g_blnScrnShotSelectorOKPressed As Boolean
Public g_blnApplyScrnShotFileToAll As Boolean
Public g_lngPictureBackColor As Long

'removing games/tunes
Public g_blnRemoveItemConfirmed As Boolean

'details tab positioning
Public Type DetailsTabPositionType
    sngLabelTwipsX As Single
    sngFieldTwipsX As Single
    sngTwipsY As Single
End Type
Public ga_udtDetailsTabPositionsLookup() As DetailsTabPositionType
Public g_lngTotalDetailsTabPositions As Long
Public ga_lngDetailsTabPositions() As Long

'extras manager
Public g_blnExtrasEditorOKPressed As Boolean
Public g_blnApplyClearExtrasToAll As Boolean
Public Type ExtrasPropsType
    strName As String
    lngType As Long
    lngPathIdx As Long
    strFile As String
    strInitPath As String
    lngGemusOption As Long
    lngGF_Control As Long
    lngGF_Players As Long
    lngGF_PalNTSC As Long
    blnGF_TDE As Boolean
    strGF_Comment As String
    strGF_VComment As String
    strGF_Notes As String
    lngGF_ImageIndex As Long
    strRunnableFile As String
End Type
Public g_udtExtrasProps As ExtrasPropsType

'verify games
Public g_blnVerifyDone As Boolean

'inputbox window
Public Type InputBoxWindowType
    blnAllowEmptyReturn As Boolean
    blnOKClicked As Boolean
    lngMaxLength As Long
    strCancelButtonText As String
    strCaption As String
    strInvalidChars As String
    strOKButtonText As String
    strPrompt As String
    strReturnedText As String
    strStartingText As String
End Type
Public g_udtInputBox As InputBoxWindowType

Public g_blnImportPersonalInfoSuccess As Boolean

'first load window
Public Type FirstLoadInfoType
    blnGemus As Boolean
    strMessage As String
End Type
Public g_udtFirstLoad As FirstLoadInfoType

'inputbox window with checkbox
Public Type InputBoxWindowType2
    blnAllowEmptyReturn As Boolean
    blnOKClicked As Boolean
    lngCheckBoxStartingValue As Long
    lngCheckBoxReturnedValue As Long
    lngMaxLength As Long
    strCancelButtonText As String
    strCaption As String
    strInvalidChars As String
    strOKButtonText As String
    strPrompt As String
    strReturnedText As String
    strStartingText As String
    strCheckBoxText As String
End Type
Public g_udtInputBox2 As InputBoxWindowType2

'for edit game info window
Public Type EditGameInfoType
    blnTableEdited As Boolean
    blnEditAll As Boolean
    blnOKClicked As Boolean
    blnUpdateYear As Boolean
    blnUpdatePublisher As Boolean
    blnUpdateProgrammer As Boolean
    blnUpdateMusician As Boolean
    blnUpdateLanguage As Boolean
    blnUpdateArtist As Boolean
    blnUpdateDeveloper As Boolean
    blnUpdateLicense As Boolean
    blnUpdateRarity As Boolean
    blnUpdateGenre As Boolean
    blnUpdateControl As Boolean
    blnUpdatePlayers As Boolean
    blnUpdateComment As Boolean
    blnUpdateWebLink As Boolean
    blnUpdateReviewRating As Boolean
    lngYearID As Long
    lngPublisherID As Long
    lngMusicianID As Long
    lngProgrammerID As Long
    lngLanguageID As Long
    lngArtistID As Long
    lngDeveloperID As Long
    lngLicenseID As Long
    lngRarityID As Long
    lngGenreID As Long
    strComment As String
    strWebLinkName As String
    strWebLinkURL As String
    lngPlayersFrom As Long
    lngPlayersTo As Long
    lngControlID As Long
    blnSimPlay As Boolean
    lngReviewRating As Long
End Type
Public g_udtEditGameInfo As EditGameInfoType
Public Enum EditGameInfoFocusEnum
    e_InfoFocus_Year
    e_InfoFocus_Publisher
    e_InfoFocus_Programmer
    e_InfoFocus_Musician
    e_InfoFocus_Language
    e_InfoFocus_Genre
    e_InfoFocus_Control
    e_infoFocus_Players
    e_InfoFocus_Comment
    e_InfoFocus_Artist
    e_InfoFocus_Developer
    e_InfoFocus_License
    e_InfoFocus_Rarity
    e_InfoFocus_WebLink
    e_InfoFocus_ReviewRating
End Enum
Public g_udtEditGameInfoFocus As EditGameInfoFocusEnum

'for edit version info window
Public Type EditVersionInfoType
    blnTableEdited As Boolean
    blnEditAll As Boolean
    blnOKClicked As Boolean
    blnUpdateCracker As Boolean
    blnUpdateTrainers As Boolean
    blnUpdateHSS As Boolean
    blnUpdateLength As Boolean
    blnUpdateNTSC As Boolean
    blnUpdateTDE As Boolean
    blnUpdateLoadingScreen As Boolean
    blnUpdateDocs As Boolean
    blnUpdateVComment As Boolean
    blnUpdateWebLink As Boolean
    blnUpdatePlayable As Boolean
    blnUpdateOriginal As Boolean
    blnUpdateTitleScreen As Boolean
    lngCrackerID As Long
    lngTrainers As Long
    lngHSS As Long
    lngGameLength As Long
    lngGameLengthType As Long
    lngNTSC As Long
    lngTDE As Long
    lngLoadingScreen As Long
    lngDocs As Long
    lngPlayable As Long
    lngOriginal As Long
    lngTitleScreen As Long
    strVComment As String
    strWebLinkName As String
    strWebLinkURL As String
End Type
Public g_udtEditVersionInfo As EditVersionInfoType
Public Enum EditVersionInfoFocusEnum
    e_VersionFocus_Cracker
    e_VersionFocus_Trainers
    e_VersionFocus_HSS
    e_VersionFocus_Length
    e_VersionFocus_NTSC
    e_VersionFocus_TDE
    e_VersionFocus_LoadingScreen
    e_VersionFocus_Docs
    e_VersionFocus_VComment
    e_VersionFocus_WebLink
    e_VersionFocus_TitleScreen
    e_VersionFocus_Playable
    e_VersionFocus_Original
End Enum
Public g_udtEditVersionInfoFocus As EditVersionInfoFocusEnum

'for edit personal info window
Public Type EditPersonalInfoType
    blnEditAll As Boolean
    blnOKClicked As Boolean
    blnUpdateHighScore As Boolean
    blnUpdateDiff As Boolean
    blnUpdateRating As Boolean
    blnUpdateTP As Boolean
    blnUpdateLP As Boolean
    blnUpdateAdult As Boolean
    blnUpdateFav As Boolean
    blnUpdateNotes As Boolean
    strHighScore As String
    lngDiffID As Long
    lngRating As Long
    blnClassic As Boolean
    blnTP As Boolean
    blnLP As Boolean
    blnAdult As Boolean
    blnFav As Boolean
    strNotes As String
End Type
Public g_udtEditPersonalInfo As EditPersonalInfoType
Public Enum EditPersonalInfoFocusEnum
    e_PersonalFocus_HighScore
    e_PersonalFocus_Difficulty
    e_PersonalFocus_Rating
    e_PersonalFocus_TP
    e_PersonalFocus_LP
    e_PersonalFocus_Adult
    e_PersonalFocus_Favourite
    e_PersonalFocus_Memo
End Enum
Public g_udtEditPersonalInfoFocus As EditPersonalInfoFocusEnum

'for edit Music Tune info window
Public Type EditMusicInfoType
    blnTableEdited As Boolean
    blnEditAll As Boolean
    blnOKClicked As Boolean
    blnUpdateMusician As Boolean
    blnUpdateAdult As Boolean
    blnUpdateFav As Boolean
    lngMusicianID As Long
    blnAdult As Boolean
    blnFav As Boolean
End Type
Public g_udtEditMusicInfo As EditMusicInfoType
Public Enum EditMusicInfoFocusEnum
    e_MusicFocus_Musician
    e_MusicFocus_Adult
    e_MusicFocus_Fav
End Enum
Public g_udtEditMusicInfoFocus As EditMusicInfoFocusEnum

'record selector
Public Enum RecordSelectorEnum
    e_RecordSelector_Publisher
    e_RecordSelector_Programmer
    e_RecordSelector_Musician
    e_RecordSelector_Language
    e_RecordSelector_Cracker
    e_RecordSelector_Artist
    e_RecordSelector_Developer
    e_RecordSelector_License
    e_RecordSelector_Rarity
End Enum
Public Type RecordSelectorType
    eRecordSelectorType As RecordSelectorEnum
    blnOKClicked As Boolean
    lngID As Long
    strSQL As String
    strReturnedName As String
End Type
Public g_udtRecordSelector As RecordSelectorType

'edit musician properties window
Public Type MusicianPropertiesType
    blnOKClicked As Boolean
    strMusician As String
    strGroup As String
    strNick As String
    strPhoto As String
    blnAdding As Boolean
    strFocus As String
End Type
Public g_udtMusicianProps As MusicianPropertiesType

Public Type FileSelectorType
    blnOKClicked As Boolean
    strSelectFile As String
    strReturnedFile As String
    strPattern As String
    strFolder As String
    strCaption As String
End Type
Public g_udtFileSelectorProps As FileSelectorType

'for speed
Public ga_strLengthTypeLookupTable(19) As String
Public ga_strYearLookupTable(1 To gc_lngYEARLOOKUPMAX) As String
Public ga_strPalNTSCLookupTable(3) As String
Public ga_strYesNoLookupTable(1) As String
Public ga_strYesNoUnknownLookupTable(2) As String
Public ga_strPlayerSimLookupTable(1) As String
Public ga_strRatingLookupTable(gc_lngRATINGLOOKUPMAX) As String
Public ga_strControlLookupTable(gc_lngCONTROLLOOKUPMAX) As String
Public ga_strGenreParentLookupTable() As String
Public Type GameNameLookupTableType
    strCommaText As String
    lngLenCommaText As Long
    strReplaceText As String
    lngLenReplaceText As Long
End Type
Public ga_udtGameNameLookupTable(gc_lngGAMENAMELOOKUPMAX) As GameNameLookupTableType

'field lengths
Public Enum EnumFieldLengths
    e_FldLen_Comment = 100
    e_FldLen_VersionComment = 100
End Enum

'snug picture fit
Public Type SnugFitType
    lngSourceW As Long
    lngSourceH As Long
    lngDestW As Long
    lngDestH As Long
    lngBlitX As Long
    lngBlitY As Long
    lngBlitW As Long
    lngBlitH As Long
End Type

'what is selected?
Public Enum EnumSelectedType
    e_SelType_IsGame
    e_SelType_IsSid
    e_SelType_IsNone
End Enum
Public g_eWhatIsSelected As EnumSelectedType

'Game data
Public Type GameDataType
    lngID As Long
    blnAdult As Boolean
    blnFav As Boolean
    lngDiffID As Long
    lngGenreID As Long
    lngPlayersFrom As Long
    lngPlayersTo As Long
    blnPlayersSim As Boolean
    blnGameFileExists As Boolean
    lngRating As Long
    strName As String
    strPublisher As String
    strYear As String
    strProgrammer As String
    strMusician As String
    strGenre As String
    strPlayers As String
    strLanguage As String
    strCracker As String
    lngTrainers As Long
    strTrainers As String
    strHighScoreSaver As String
    lngHSS As Long
    strGameLength As String
    lngGameLength As Long
    lngGameLengthType As Long
    strTrueDriveEmulation As String
    lngTrueDriveEmulation As Long
    strNTSC As String
    lngNTSC As Long
    strLoadingScreen As String
    lngLoadingScreen As Long
    strDocs As String
    lngDocs As Long
    strComment As String
    strVersionComment As String
    strHighScore As String
    strDifficulty As String
    strRating As String
    strTimesPlayed As String
    strDateLastPlayed As String
    blnMemoLoaded As Boolean
    strMemo As String
    lngPrequel As Long
    lngSequel As Long
    lngRelated As Long
    lngCloneOf As Long
    strPrequel As String
    strSequel As String
    strRelated As String
    strCloneOf As String
    blnClassic As Boolean
    strGameFilename As String
    strFileToRun As String
    lngGameFilenameIndex As Long
    strScreenshotFilename As String
    lngControl As Long
    strControl As String
    lngYearID As Long
    lngPublisherID As Long
    lngProgrammerID As Long
    lngLanguageID As Long
    lngCrackerID As Long
    blnSimPlay As Boolean
    blnLoaded_0 As Boolean
    blnLoaded_1 As Boolean
    blnLoaded_2 As Boolean
    blnLoaded_3 As Boolean
    blnLoaded_Extras As Boolean
    blnHasExtras As Boolean
    lngVer As Long
    strPairs As String
    lngDeveloperID As Long
    lngArtistID As Long
    lngLicenseID As Long
    lngRarityID As Long
    strRarity As String
    strLicense As String
    strDeveloper As String
    strArtist As String
    strWebLinkName As String
    strWebLinkURL As String
    blnHasWebLink As Boolean
    strVWebLinkName As String
    strVWebLinkURL As String
    blnHasVWebLink As Boolean
    lngTitleScreen As Long
    strTitleScreen As String
    lngPlayable As Long
    strPlayable As String
    lngOriginal As Long
    strOriginal As String
    lngReviewRating As Long
End Type
Public g_udtGame As GameDataType
'Sid data
Public Type SidDataType
    blnGotInfo As Boolean
    blnGameTable As Boolean
    strName As String
    strFilename As String
    lngID As Long
    lngMU_Id As Long
    strMusician As String
    strNick As String
    strGroup As String
    strPhoto As String
    strSFName As String
    strSFAuthor As String
    strSFCopyright As String
    lngSFTotalSongs As Long
    lngSFDefaultSong As Long
    blnFileExists As Boolean
    blnIsPSID As Boolean
    blnFav As Boolean
    blnAdult As Boolean
End Type
Public g_udtSid As SidDataType

'forms loaded flags
Public g_blnLoadedfrmBigScreenshot As Boolean
Public g_blnLoadedfrmCategories As Boolean
Public g_blnLoadedfrmNotes As Boolean
Public g_blnLoadedfrmSocket As Boolean

'option flags/variables
Public g_blnReZip As Boolean
Public g_blnRezipAsk As Boolean
Public ge_RezipAdditional As Gemus.EnumRezipAdditionalFiles
Public g_blnZipNotify As Boolean
Public g_blnBossKeyOn As Boolean
Public g_blnSetPicZoomForeground As Boolean
Public g_blnMamePicStyle As Boolean
Public g_blnOpenFilter64 As Boolean
Public g_blnPromptForGamePathDevice As Boolean
Public Enum DoubleClickActionEnum
    e_DoubleClick_PlayGame
    e_DoubleClick_PlaySid
End Enum
Public g_eDoubleClickAction As DoubleClickActionEnum

Public g_blnSkipVersionCheck As Boolean
Public g_blnOptionsShowGameID As Boolean
Public g_blnOptionsReClickGame As Boolean
Public g_blnOptionsReLoadList As Boolean
Public g_blnChangedEmulatorSettings As Boolean
Public g_blnChangedSidplayerSettings As Boolean
Public g_strBrowseForFolder As String
Public g_blnShutDownSidPlayersOnExit As Boolean
Public g_lngRecentStore As Long
Public g_lngRecentUpdate As Long
Public g_strRecentEmpty As String
Public g_blnRecentGameSid As Boolean
Public Const gc_lngRECENT_GAMES As Long = 0
Public Const gc_lngRECENT_SIDS As Long = 1
Public Const gc_lngRECENT_BOTH As Long = 2
Public Const gc_lngRECENT_SELECTED As Long = 0
Public Const gc_lngRECENT_PLAYED As Long = 1

Public Enum PathLocationEnum
    e_PathLocation_Removable
    e_PathLocation_FixedHD
    e_PathLocation_CD
    e_PathLocation_Invalid
End Enum
Public Type PathsType
    strWithBackSlash As String
    strRealPath As String
    ePathType As PathLocationEnum
    blnIsSet As Boolean
End Type
Public g_udtRegPathSid As PathsType
Public ga_udtRegPathGame() As PathsType
Public ga_udtRegPathScrnShot() As PathsType
Public ga_udtRegPathExtras() As PathsType
Public g_udtPathGBGame As PathsType
Public g_lngTotalGamePaths As Long
Public g_lngTotalScrnShotPaths As Long
Public g_lngTotalExtrasPaths As Long
Public g_udtRegPathPhotos As PathsType
Public g_blnValidGamePathExists As Boolean

'sid players
Public Type SidPlayerType
    strName As String
    strFullPathAndFilename As String
    strFullPathOnly As String
    strFullPathOnlyWithBackSlash As String
    strFullFilenameOnly As String
    strShortPathAndFilename As String
    strShortPathOnly As String
    strShortPathOnlyWithBackSlash As String
    strShortFilenameOnly As String
    strCommandLine As String
    blnUseShortFilenames As Boolean
    ePathType As PathLocationEnum
    lngTotalTypes As Long
    a_strTypes() As String
End Type
Public ga_udtSidPlayers() As SidPlayerType
Public g_lngTotalSidPlayers As Long
Public ga_lngSidPlayerHandles() As Long
Public g_lngTotalSidPlayerHandles As Long
Public Type SidPlayerPropsType
    strName As String
    strFullPath As String
    strCommandLine As String
    blnUseShort As Boolean
    a_strTypes() As String
    lngTotalTypes As Long
End Type
Public g_udtSidPlayerProps As SidPlayerPropsType

'for tips
Public Const gc_lngMAXTIPS As Long = 8
Public Const gc_lngTIPS_RUNGAME = 0
Public Const gc_lngTIPS_PLAYSID = 1
Public Const gc_lngTIPS_DETAILSTABS = 2
Public Const gc_lngTIPS_SCREENSHOTS = 3
Public Const gc_lngTIPS_SCREENSHOTS2 = 4
Public Const gc_lngTIPS_GAMELINKS = 5
Public Const gc_lngTIPS_TESTGAMEFLAG = 6
Public Const gc_lngTIPS_PICTURESTYLE = 7
Public Const gc_lngTIPS_EDITEXTRAS = 8
Public Type GBTipType
    blnShowIt As Boolean
    strINIAndRegKey As String
End Type
Public g_udtSelectedTip As GBTipType
Public g_strSelectedTipTitle As String
Public g_strSelectedTipText As String
Public g_udtMainTips(gc_lngMAXTIPS) As GBTipType

'game link selector
Public g_strMarkedGame As String
Public g_strGameLinkField As String
Public g_blnGameLinkSetFlag As Boolean
'Public g_blnMaintainGameLinks As Boolean
Public g_strDefaultGameLinkField As String

'if no g_udtGame paths are valid this equals False
Public g_blnAGamePathSet As Boolean

'for g_udtGame selector
Public g_strLastGSPath As String

Public g_blnOKClicked As Boolean

Public Function CreateDefaultGBGameFolder() As Boolean

    On Error GoTo Err_Handler

    Call MkDir(GetDefaultGBGameFolder())
    CreateDefaultGBGameFolder = g_xobjFunc.DirExists(GetDefaultGBGameFolder())
    Exit Function

Err_Handler:
    CreateDefaultGBGameFolder = False

End Function
Public Sub Debug_MsgBox(strMessage As String)

    MsgBox strMessage

End Sub


Public Function GetDefaultGBGameFolder() As String

    GetDefaultGBGameFolder = Left$(g_strAppPathWithBackslash, 3) & "GBGame"

End Function

Public Function GetReviewRatingImgIdx(ByVal lngReviewRating_in As Long)

    GetReviewRatingImgIdx = Int(lngReviewRating_in / 2)
    If lngReviewRating_in Mod 2 = 1 Then
        GetReviewRatingImgIdx = GetReviewRatingImgIdx + 1
    End If

End Function

Public Function ParseExtraColumns(a_udtExtraColumns() As ExtraColumnType, strExtraColumns_in As String)

    On Error GoTo Err_Handler

    Dim a_strColumns() As String
    Dim a_strColumn() As String
    Dim lngIdx As Long
    Dim lngUBoundColumns As Long
    Dim lngMaxColumns As Long
    Const cMAX0 As Long = 14
    Const cMAX1 As Long = 16
    Const cMAX2 As Long = 5
    Dim lngType As Long
    Dim lngCol As Long
    Dim blnOK As Boolean
    
    ReDim a_udtExtraColumns(0)
    ParseExtraColumns = 0
    
    '"0-1,2-3,1-10,..."
    strExtraColumns_in = Trim$(strExtraColumns_in)

    If strExtraColumns_in = "" Then Exit Function

    a_strColumns = Split(strExtraColumns_in, ",")
    lngUBoundColumns = UBound(a_strColumns)
    lngMaxColumns = 0
    For lngIdx = 0 To lngUBoundColumns
        a_strColumn = Split(a_strColumns(lngIdx), "-")
        If UBound(a_strColumn) = 1 Then
            blnOK = False
            lngType = Val(a_strColumn(0))
            lngCol = Val(a_strColumn(1))
            If lngType = 0 Then
                If lngCol >= 0 Or lngCol <= cMAX0 Then
                    blnOK = True
                End If
            ElseIf lngType = 1 Then
                If (lngCol >= 0 Or lngCol <= cMAX1) And lngCol <> 15 Then
                    blnOK = True
                End If
            ElseIf lngType = 2 Then
                If lngCol >= 0 Or lngCol <= cMAX2 Then
                    blnOK = True
                End If
            End If
            If blnOK Then
                'add this column to the structure
                ReDim Preserve a_udtExtraColumns(lngMaxColumns)
                a_udtExtraColumns(lngMaxColumns).lngTypeIndex = lngType
                a_udtExtraColumns(lngMaxColumns).strColumnName = ga_udtExtraColumns(lngType).a_udtColumnItems(lngCol).strColumnTitle
                a_udtExtraColumns(lngMaxColumns).lngFieldIndex = lngCol
                a_udtExtraColumns(lngMaxColumns).lngColumnHeaderWidth = gc_lngDEFAULTCOLUMNWIDTH
                lngMaxColumns = lngMaxColumns + 1
            End If
        End If
    Next lngIdx

    ParseExtraColumns = lngMaxColumns
    Exit Function

Err_Handler:
    ReDim a_udtExtraColumns(0)
    ParseExtraColumns = 0

End Function
Public Function XMLEscape(strUnescaped_in As String) As String

    'WildCard Replacer

    XMLEscape = Replace(strUnescaped_in, "&", "&amp;")
    XMLEscape = Replace(XMLEscape, """", "&quot;")
    XMLEscape = Replace(XMLEscape, "<", "&lt;")
    XMLEscape = Replace(XMLEscape, ">", "&gt;")

End Function
Public Function DecMess(strEnc_in As String) As String

    'simple decrypt string.. implemented to protect against hex-editor morons

    Dim l As Long
    Dim i As Long
    Dim s As String
    Dim c As Long

    l = Len(strEnc_in)
    s = ""

    For i = l To 1 Step -1
        c = Asc(Mid$(strEnc_in, i, 1))
        s = s & Chr$(c + 1)
    Next i

    DecMess = s

End Function
Public Function EncMess(strEnc_in As String) As String

    'simple encrypt string.. implemented to protect against hex-editor morons

    Dim l As Long
    Dim i As Long
    Dim s As String
    Dim c As Long

    l = Len(strEnc_in)
    s = ""

    For i = 1 To l
        c = Asc(Mid$(strEnc_in, i, 1))
        s = Chr$(c - 1) & s
    Next i

    EncMess = s

End Function

Public Function GetAppropriatePlayer(strMusicFile_in As String) As Long

    'returns index of player or -1 if no matching filetype

    Dim lngPlayerIdx As Long
    Dim strType As String
    Dim lngFTIdx As Long

    GetAppropriatePlayer = -1

    If g_lngTotalSidPlayers > 0 Then
        'get the type
        strType = UCase$(Right$(strMusicFile_in, Len(strMusicFile_in) - InStrRev(strMusicFile_in, ".")))
        'find a player with a matching type
        For lngPlayerIdx = 0 To (g_lngTotalSidPlayers - 1)
            With ga_udtSidPlayers(lngPlayerIdx)
                For lngFTIdx = 0 To (.lngTotalTypes - 1)
                    If UCase$(.a_strTypes(lngFTIdx)) = strType Then
                        GetAppropriatePlayer = lngPlayerIdx
                        Exit Function
                    ElseIf .a_strTypes(lngFTIdx) = "*" Then
                        GetAppropriatePlayer = lngPlayerIdx
                        Exit Function
                    End If
                Next lngFTIdx
            End With
        Next lngPlayerIdx
    End If

End Function


Public Sub GetSnugFitImageDimensions(udtSnug_inout As SnugFitType)

    'gets the blit dimensions for fitting and centering a larger or
    'smaller source image into a destination area of a specified size
    'dest size ratio must be 1.6 (w/h)

    Dim sngRatio As Single

    With udtSnug_inout
        If (.lngSourceH > .lngDestH) Or (.lngSourceW > .lngDestW) Then
            sngRatio = CSng(.lngSourceW / .lngSourceH)
            If sngRatio < 1.6 Then
                'make height 100%
                .lngBlitY = 0
                .lngBlitH = .lngDestH
                sngRatio = .lngDestH / .lngSourceH
                .lngBlitW = .lngSourceW * sngRatio
                .lngBlitX = (.lngDestW / 2) - (.lngBlitW / 2)
            ElseIf sngRatio = 1.6 Then
                'make both 100%
                .lngBlitX = 0
                .lngBlitY = 0
                .lngBlitH = .lngDestH
                .lngBlitW = .lngDestW
            Else
                'make width 100%
                .lngBlitX = 0
                .lngBlitW = .lngDestW
                sngRatio = .lngDestW / .lngSourceW
                .lngBlitH = .lngSourceH * sngRatio
                .lngBlitY = (.lngDestH / 2) - (.lngBlitH / 2)
            End If
        Else
            'image small enough to fit, just center it
            .lngBlitX = (.lngDestW / 2) - (.lngSourceW / 2)
            .lngBlitY = (.lngDestH / 2) - (.lngSourceH / 2)
            .lngBlitW = .lngSourceW
            .lngBlitH = .lngSourceH
        End If
    End With
    

End Sub

Public Function Muddle(ByVal strText_in As String) As String

    Dim lngIdx As Long
    Dim strOut As String
    Dim strChar As String
    Dim strAscVal As String

    strOut = ""
    'reverse, taking length of each ascii value then each ascii value
    For lngIdx = Len(strText_in) To 1 Step -1
        strChar = Mid$(strText_in, lngIdx, 1)
        strAscVal = Trim$(Asc(strChar))
        strOut = strOut & Format$(Len(strAscVal), "000") & strAscVal
    Next lngIdx
    'add one to each number, looping to 0 if at 9
    For lngIdx = 1 To Len(strOut)
        strChar = Mid$(strOut, lngIdx, 1)
        If strChar = "9" Then
            Mid$(strOut, lngIdx, 1) = "0"
        Else
            Mid$(strOut, lngIdx, 1) = Trim$(Val(strChar) + 1)
        End If
    Next lngIdx
    'insert random numbers at odd intervals
    Muddle = ""
    Randomize
    For lngIdx = 1 To Len(strOut)
        Muddle = Muddle & Trim$(Int((9 + 1) * Rnd))
        Randomize Int((30000 + 1) * Rnd)
        Muddle = Muddle & Mid$(strOut, lngIdx, 1)
    Next lngIdx
    'reverse again
    strOut = ""
    For lngIdx = 1 To Len(Muddle)
        strOut = Mid$(Muddle, lngIdx, 1) & strOut
    Next lngIdx
    Muddle = strOut

End Function
Public Function UnMuddle(ByVal strText_in As String) As String

    Dim lngIdx As Long
    Dim strOut As String
    Dim strChar As String
    Dim strAscVal As String
    Dim lngVal As Long

    If Len(strText_in) > 3 Then
        'reverse
        strOut = ""
        For lngIdx = 1 To Len(strText_in)
            strOut = Mid$(strText_in, lngIdx, 1) & strOut
        Next lngIdx
        'remove every other random char
        strText_in = ""
        For lngIdx = 1 To Len(strOut)
            If lngIdx Mod 2 = 0 Then
                strText_in = strText_in & Mid$(strOut, lngIdx, 1)
            End If
        Next lngIdx
        'subtract one from each number, looping to 9 if at 0
        For lngIdx = 1 To Len(strText_in)
            strChar = Mid$(strText_in, lngIdx, 1)
            If strChar = "0" Then
                Mid$(strText_in, lngIdx, 1) = "9"
            Else
                Mid$(strText_in, lngIdx, 1) = Trim$(Val(strChar) - 1)
            End If
        Next lngIdx
        'create the string by taking the length of the ascii value followed by the ascii value
        strOut = ""
        lngIdx = 1
        Do Until lngIdx > Len(strText_in)
            lngVal = Val(Mid$(strText_in, lngIdx, 3))
            lngIdx = lngIdx + 3
            If lngVal = 0 Then
                UnMuddle = ""
                Exit Function
            End If
            strChar = Chr$(Val(Mid$(strText_in, lngIdx, lngVal)))
            strOut = strChar & strOut
            lngIdx = lngIdx + lngVal
        Loop
        UnMuddle = strOut
    Else
        UnMuddle = ""
    End If

End Function

Public Sub SetSidFileTypes(strFileTypes_in As String, a_strTypes_out() As String, lngTotal_out As Long)

    Dim a_strTemp() As String
    Dim lngFTIdx As Long

    lngTotal_out = 0
    ReDim a_strTypes_out(0)

    If strFileTypes_in <> "" Then
        a_strTemp = Split(strFileTypes_in, ";")
        For lngFTIdx = 0 To UBound(a_strTemp)
            a_strTemp(lngFTIdx) = Trim$(a_strTemp(lngFTIdx))
            If ValidateFileType(a_strTemp(lngFTIdx)) = True Then
                ReDim Preserve a_strTypes_out(lngTotal_out)
                a_strTypes_out(lngTotal_out) = UCase$(a_strTemp(lngFTIdx))
                lngTotal_out = lngTotal_out + 1
            End If
        Next lngFTIdx
    End If

End Sub
Private Function ValidateFileType(strFT_inout As String) As Boolean

    If strFT_inout = "*.*" Then
        strFT_inout = "*"
        ValidateFileType = True
        Exit Function
    ElseIf strFT_inout <> "" Then
        If Left$(strFT_inout, 1) = "." Then
            'strip dot
            If Len(strFT_inout) > 1 Then
                strFT_inout = Right$(strFT_inout, Len(strFT_inout) - 1)
                ValidateFileType = True
                Exit Function
            End If
        Else
            'no dot
            ValidateFileType = True
            Exit Function
        End If
    End If
    ValidateFileType = False

End Function
Public Function GetTipText(ByVal lngTipIdx_in As Long) As String

    '-----------------------------------------------
    'Gets the Tip Text for the passed in tip
    '
    'Inputs:    lngTipIdx_in:   the index of the tip
    '
    'Returns:   The Text of the tip
    '-----------------------------------------------
    
    'get the text or use default
    Select Case lngTipIdx_in
    Case gc_lngTIPS_RUNGAME
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Left-clicking the Play Game button will run the game in the first emulator that supports it's file-type.~BL~Right-clicking the button will display a list of available emulators to choose from!")
    Case gc_lngTIPS_PLAYSID
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Left-clicking the Play Music button will play the music file in the first player that supports it's file-type.~BL~Right-clicking the button will display a list of available players to choose from!")
    Case gc_lngTIPS_DETAILSTABS
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "You can use the function keys F5 to F9 to switch between the Details Tabs!")
    Case gc_lngTIPS_SCREENSHOTS
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Double Clicking on the left-hand side of a Picture will zoom in.~BL~Clicking on the right-hand side of Pictures will scroll through them (you can also press F11/F12 to scroll).~BL~Right-click a Picture for more options.")
    Case gc_lngTIPS_SCREENSHOTS2
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Use Drag/Drop to place picture files into your Pictures Path.~BL~Holding SHIFT or CTRL whilst dragging will Copy the file.  Holding down neither of these keys will Move the file.~BL~Click the Picture Selector's Help button for more info.")
    Case gc_lngTIPS_GAMELINKS
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "You can setup links to a game's Precursor, Sequel or Related game by first 'Marking' a game with CTRL+C, then clicking on the game you wish to link it to and pressing CTRL+V.~BL~You can use the 'Related Game' link to setup circular links between multiple games.~BL~E.g.  Game1 --> Game2 --> Game3 --> Game1")
    Case gc_lngTIPS_TESTGAMEFLAG
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Some settings, such as whether a game is PAL or NTSC compliant, cannot be determined simply by looking at the game file.~BL~Clicking the small checkbox by the 'Test' button will allow you to enter these settings before GEMUS runs the game in the specified emulator.")
    Case gc_lngTIPS_EDITEXTRAS
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "Game Extras are additional files that you can associate with a game!~BL~You can setup these files as 'NORMAL' Extras (i.e. they run with their default program) or you can setup additional game-files for a game by specifying that an Extra should use 'GEMUS' to run it.")
    Case gc_lngTIPS_PICTURESTYLE
        GetTipText = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey, "There are 2 'styles' that GameBase can use for displaying pictures: -~BL~The 'GameBase' style (e.g. 'pic.gif', 'pic_1.gif', 'pic_2.gif' etc. until no more can be found in your Picture Paths).~BL~The 'MAME' style (e.g. 'pic.png' is searched for in each Picture Path and displayed if found).")
    'add more here as needed
    End Select
    GetTipText = Replace(GetTipText, "~BL~", vbCrLf & vbCrLf)

End Function

Public Sub LoadTips()

    '--------------------------------------------
    'loads the main tips info into the tips array
    '--------------------------------------------
    
    Dim lngTipIdx As Long       'loop index
    
    'set the ini/reg-key property
    g_udtMainTips(gc_lngTIPS_RUNGAME).strINIAndRegKey = "Main_RunningGames"
    g_udtMainTips(gc_lngTIPS_PLAYSID).strINIAndRegKey = "Main_PlayingSids"
    g_udtMainTips(gc_lngTIPS_DETAILSTABS).strINIAndRegKey = "Main_DetailsTabs"
    g_udtMainTips(gc_lngTIPS_SCREENSHOTS).strINIAndRegKey = "Main_Pictures"
    g_udtMainTips(gc_lngTIPS_SCREENSHOTS2).strINIAndRegKey = "Picture_Selector"
    g_udtMainTips(gc_lngTIPS_GAMELINKS).strINIAndRegKey = "Linking_Games"
    g_udtMainTips(gc_lngTIPS_TESTGAMEFLAG).strINIAndRegKey = "TestGame_Flag"
    g_udtMainTips(gc_lngTIPS_EDITEXTRAS).strINIAndRegKey = "Edit_Extras"
    g_udtMainTips(gc_lngTIPS_PICTURESTYLE).strINIAndRegKey = "Picture_Style"
    'add more here as needed

    'retrieve the Show-It property from the registry and set it in the array
    For lngTipIdx = 0 To gc_lngMAXTIPS
        With g_udtMainTips(lngTipIdx)
            g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Tips", .strINIAndRegKey, .blnShowIt, False, True
        End With
    Next lngTipIdx

End Sub
Public Function GetTipTitle(ByVal lngTipIdx_in As Long) As String

    '-----------------------------------------------
    'Gets the Tip Title for the passed in tip
    '
    'Inputs:    lngTipIdx_in:   the index of the tip
    '
    'Returns:   The Title of the tip
    '-----------------------------------------------
    
    'get the title or use default
    Select Case lngTipIdx_in
    Case gc_lngTIPS_RUNGAME
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Running Games") & "..."
    Case gc_lngTIPS_PLAYSID
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Playing Sids") & "..."
    Case gc_lngTIPS_DETAILSTABS
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "The Details Tabs") & "..."
    Case gc_lngTIPS_SCREENSHOTS
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Game Pictures") & "..."
    Case gc_lngTIPS_SCREENSHOTS2
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "The Picture Selector") & "..."
    Case gc_lngTIPS_GAMELINKS
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Setting up Game Links") & "..."
    Case gc_lngTIPS_TESTGAMEFLAG
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Testing Games") & "..."
    Case gc_lngTIPS_EDITEXTRAS
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "Game Extras") & "..."
    Case gc_lngTIPS_PICTURESTYLE
        GetTipTitle = g_xobjFunc.GetINIMessage("Tips", g_udtMainTips(lngTipIdx_in).strINIAndRegKey & "_Title", "GB/MAME Picture Style") & "..."
    'add more here as needed
    End Select
    
End Function
Public Sub ShowTip(ByVal lngTipIdx_in As Long)

    '----------------------------------------------------------
    'Tests if a tip can be displayed and if so it calls frmTip
    'to display it.
    '
    'Inputs:    lngTipIdx_in:  The index of the tip to display
    '
    'Remarks:   Any saving to the registry for the tip is done
    '           in frmTip (for changing it's ShowIt value). The
    '       selected tip's blnShowIt parameter is changed
    '       in frmTip if the user selected not to show the
    '       tip again.  frmTip also handles if the user
    '       selects not to show ANY tips again.
    '----------------------------------------------------------

    With g_udtMainTips(lngTipIdx_in)
        'only process if we want to show the tip
        If .blnShowIt Then
        
            'copy to the global SelectedTip variables used in frmTip
            g_udtSelectedTip.blnShowIt = .blnShowIt
            g_udtSelectedTip.strINIAndRegKey = .strINIAndRegKey
            g_strSelectedTipTitle = GetTipTitle(lngTipIdx_in)
            g_strSelectedTipText = GetTipText(lngTipIdx_in)
            
            'show the tip and save the .blnShotIt property to the registry if needed
            frmTip.Show vbModal
            
            'update the actual tip's blnShowIt parameter
            .blnShowIt = g_udtSelectedTip.blnShowIt
        End If
    End With

End Sub
Public Sub AddToSidPlayerHandles(ByVal lngHandle_in As Long)

    If lngHandle_in <> 0 Then
        g_lngTotalSidPlayerHandles = g_lngTotalSidPlayerHandles + 1
        ReDim Preserve ga_lngSidPlayerHandles(g_lngTotalSidPlayerHandles - 1)
        ga_lngSidPlayerHandles(g_lngTotalSidPlayerHandles - 1) = lngHandle_in
    End If

End Sub

Public Function FindGameFile(lngGPIdx_out As Long, strFilename_in As String) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalGamePaths = 0 Then
        FindGameFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoGamePaths", "You have not set up any Game Paths!")
    ElseIf g_blnValidGamePathExists = False Then
        FindGameFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidGamePath", "There are no valid Game Paths!")
    Else
        'test each path
        strNotFound = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!")
        For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
            With ga_udtRegPathGame(lngPathIdx)
                If g_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                    'found file
                    lngGPIdx_out = lngPathIdx
                    FindGameFile = ""
                    Exit Function
                Else
                    'file not found in current game path
                    If (.ePathType = e_PathLocation_CD) Or (.ePathType = e_PathLocation_Removable) Then
                        If g_blnPromptForGamePathDevice = True Then
                            'path is CD or removable drive... let user insert media
                            Do
                                mbRes = MsgBox(strNotFound & vbCrLf & vbCrLf & "'" & .strWithBackSlash & strFilename_in & "'", vbExclamation Or vbAbortRetryIgnore, gc_strAPPNAME)
                                If mbRes = vbAbort Then
                                    FindGameFile = "~~ABORTED~~"
                                    Exit Function
                                ElseIf mbRes = vbIgnore Then
                                    Exit Do
                                Else
                                    If g_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                                        'found file
                                        lngGPIdx_out = lngPathIdx
                                        FindGameFile = ""
                                        Exit Function
                                    End If
                                End If
                            Loop
                        End If
                    End If
                End If
            End With
        Next lngPathIdx
        'if here then the game wasn't found
        FindGameFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInGamePaths", "The file was not found in any of your Game Paths!")
    End If

End Function
Public Function FindScrnShotFile(lngSSIdx_out As Long, strFilename_in As String) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalScrnShotPaths = 0 Then
        FindScrnShotFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoPicturePaths", "You have not set up any Picture Paths!")
    Else
        'test each path
        For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            With ga_udtRegPathScrnShot(lngPathIdx)
                If g_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                    'found file
                    lngSSIdx_out = lngPathIdx
                    FindScrnShotFile = ""
                    Exit Function
                End If
            End With
        Next lngPathIdx
        'if here then the picture wasn't found
        FindScrnShotFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInPicturePaths", "The file was not found in any of your Picture Paths!")
    End If

End Function

Public Function FindExtrasFile(lngExtrasIdx_out As Long, strFilename_in As String) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalExtrasPaths = 0 Then
        FindExtrasFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoExtrasPaths", "You have not set up any Extras Paths!")
    Else
        'test each path
        For lngPathIdx = 0 To (g_lngTotalExtrasPaths - 1)
            With ga_udtRegPathExtras(lngPathIdx)
                If g_xobjFunc.FileExists(.strWithBackSlash & strFilename_in, False, "") = True Then
                    'found file
                    lngExtrasIdx_out = lngPathIdx
                    FindExtrasFile = ""
                    Exit Function
                End If
            End With
        Next lngPathIdx
        'if here then the extra wasn't found
        FindExtrasFile = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotInExtrasPaths", "The file was not found in any of your Extras Paths!")
    End If

End Function
Public Function GetFileBoxPatternStringFromArray(a_strTypes_in() As String) As String

    'Creates a Pattern string for the FileBox Control from the
    'passed in array of File Extensions

    Dim lngMaxArray As Long
    Dim lngArrayIdx As Long

    lngMaxArray = UBound(a_strTypes_in)
    GetFileBoxPatternStringFromArray = ""

    For lngArrayIdx = 0 To lngMaxArray
        If a_strTypes_in(lngArrayIdx) = "*" Then
            GetFileBoxPatternStringFromArray = "*.*;"
            Exit Function
        End If
        GetFileBoxPatternStringFromArray = GetFileBoxPatternStringFromArray & "*." & a_strTypes_in(lngArrayIdx) & ";"
    Next lngArrayIdx
    If Right$(GetFileBoxPatternStringFromArray, 1) = ";" Then
        GetFileBoxPatternStringFromArray = Left$(GetFileBoxPatternStringFromArray, Len(GetFileBoxPatternStringFromArray) - 1)
    End If

End Function
Public Function GetSingleMusicTypesArray(a_strTypes_out() As String) As Long

    'Gets all the File Types supported by music players, and removes duplicates
    'Returns the number of unique file types in the destination array

    Dim lngPlayerIdx As Long
    Dim lngArrayIdx As Long
    Dim lngTypesIdx As Long
    Dim blnFound As Boolean
    Dim strType As String

    GetSingleMusicTypesArray = 0
    ReDim a_strTypes_out(0)

    For lngPlayerIdx = 0 To (g_lngTotalSidPlayers - 1)
        With ga_udtSidPlayers(lngPlayerIdx)
            If .lngTotalTypes > 0 Then
                'loop through each item in the types array
                For lngTypesIdx = 0 To (.lngTotalTypes - 1)
                    'store if not already in array
                    blnFound = False
                    strType = UCase$(.a_strTypes(lngTypesIdx))
                    For lngArrayIdx = 0 To UBound(a_strTypes_out)
                        If UCase$(a_strTypes_out(lngArrayIdx)) = strType Then
                            blnFound = True
                            Exit For
                        End If
                    Next lngArrayIdx
                    If blnFound = False Then
                        ReDim Preserve a_strTypes_out(GetSingleMusicTypesArray)
                        a_strTypes_out(GetSingleMusicTypesArray) = strType
                        GetSingleMusicTypesArray = GetSingleMusicTypesArray + 1
                        If strType = "*" Then
                            Exit Function
                        End If
                    End If
                Next lngTypesIdx
            End If
        End With
    Next lngPlayerIdx

End Function
Public Function LoadGBImageFromFile(imgBox As Image, strFilename_in As String) As Boolean

    'loads the specified file into the specified imagebox

    On Error GoTo Err_Handler

    If g_xobjFunc.FileExists(strFilename_in, False, gc_strAPPNAME) = True Then
        imgBox.Picture = LoadPicture(strFilename_in)
        LoadGBImageFromFile = True
    Else
        LoadGBImageFromFile = False
    End If

    Exit Function

Err_Handler:
    LoadGBImageFromFile = False

End Function
Public Function GetValidGamePath() As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalGamePaths = 0 Then
        GetValidGamePath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoGamePaths", "You have not set up any Game Paths!")
    ElseIf g_blnValidGamePathExists = False Then
        GetValidGamePath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidGamePath", "There are no valid Game Paths!")
    Else
        'test each path
        For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
            With ga_udtRegPathGame(lngPathIdx)
                If g_xobjFunc.DirExists(.strRealPath) = True Then
                    'found path
                    GetValidGamePath = ""
                    Exit Function
                Else
                    'current path not found
                    If (.ePathType = e_PathLocation_CD) Or (.ePathType = e_PathLocation_Removable) Then
                        If g_blnPromptForGamePathDevice = True Then
                            'path is CD or removable drive... let user insert media
                            strNotFound = g_xobjFunc.GetINIMessage("Error Messages", "Error_PathNotFound", "The path was not found!")
                            Do
                                mbRes = MsgBox(strNotFound & vbCrLf & vbCrLf & "'" & .strRealPath & "'", vbExclamation Or vbAbortRetryIgnore, gc_strAPPNAME)
                                If mbRes = vbAbort Then
                                    GetValidGamePath = "~~ABORTED~~"
                                    Exit Function
                                ElseIf mbRes = vbIgnore Then
                                    Exit Do
                                Else
                                    If g_xobjFunc.DirExists(.strRealPath) = True Then
                                        'found path
                                        GetValidGamePath = ""
                                        Exit Function
                                    End If
                                End If
                            Loop
                        End If
                    End If
                End If
            End With
        Next lngPathIdx
        'if here then a path wasn't found
        GetValidGamePath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidGamePath", "There are no valid Game Paths!")
    End If

End Function
Public Function GetValidScrnShotPath(lngSSIdx_out As Long) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalScrnShotPaths = 0 Then
        GetValidScrnShotPath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoPicturePaths", "You have not set up any Picture Paths!")
    Else
        'test each path
        For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            With ga_udtRegPathScrnShot(lngPathIdx)
                If g_xobjFunc.DirExists(.strRealPath) = True Then
                    'found path
                    GetValidScrnShotPath = ""
                    lngSSIdx_out = lngPathIdx
                    Exit Function
                End If
            End With
        Next lngPathIdx
        'if here then a path wasn't found
        GetValidScrnShotPath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidPicturePath", "There are no valid Picture Paths!")
    End If

End Function
Public Function GetValidExtrasPath(lngEXIdx_out As Long) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult

    If g_lngTotalExtrasPaths = 0 Then
        GetValidExtrasPath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoExtrasPaths", "You have not set up any Extras Paths!")
    Else
        'test each path
        For lngPathIdx = 0 To (g_lngTotalExtrasPaths - 1)
            With ga_udtRegPathExtras(lngPathIdx)
                If g_xobjFunc.DirExists(.strRealPath) = True Then
                    'found path
                    GetValidExtrasPath = ""
                    lngEXIdx_out = lngPathIdx
                    Exit Function
                End If
            End With
        Next lngPathIdx
        'if here then a path wasn't found
        GetValidExtrasPath = g_xobjFunc.GetINIMessage("Error Messages", "Error_NoValidExtrasPath", "There are no valid Extras Paths!")
    End If

End Function
Public Function GetMatchingExtrasPath(strPath_in As String, lngEXIdx_out As Long) As String
    
    Dim lngPathIdx As Long      'loop index
    Dim strNotFound As String   'for error message
    Dim mbRes As VbMsgBoxResult
    Dim strPathUpper As String

    If g_lngTotalExtrasPaths = 0 Then
        GetMatchingExtrasPath = "Oops!"
    Else
        If g_xobjFunc.DirExists(strPath_in) = False Then
            'path to match doesn't exist!
            GetMatchingExtrasPath = "Oops!"
        Else
            'test each path
            strPathUpper = UCase$(strPath_in)
            If Right$(strPathUpper, 2) <> "\" Then
                strPathUpper = strPathUpper & "\"
            End If
            For lngPathIdx = 0 To (g_lngTotalExtrasPaths - 1)
                With ga_udtRegPathExtras(lngPathIdx)
                    If g_xobjFunc.DirExists(.strRealPath) = True Then
                        'match the path
                        If Left$(strPathUpper, Len(.strWithBackSlash)) = UCase$(.strWithBackSlash) Then
                            GetMatchingExtrasPath = ""
                            lngEXIdx_out = lngPathIdx
                            Exit Function
                        End If
                    End If
                End With
            Next lngPathIdx
            'if here then a path wasn't found
            GetMatchingExtrasPath = "Oops!"
        End If
    End If

End Function
Public Function GetExtrasDataString() As String

    'returns the string to save in the Extras table in the Data field
    'the string is created from various items in the g_udtExtrasProps structure

    With g_udtExtrasProps
        If .lngType = 0 Then
            'no data as not using GEMUS
            GetExtrasDataString = ""
        Else
            If .lngGemusOption = 0 Then
                'use defaults
                GetExtrasDataString = "0 " & gc_strEXTRASDELIM & .lngGF_ImageIndex
            ElseIf .lngGemusOption = 1 Then
                'use current
                GetExtrasDataString = "1" & gc_strEXTRASDELIM & .lngGF_ImageIndex
            Else
                'use specific
                GetExtrasDataString = "2" & gc_strEXTRASDELIM & .lngGF_ImageIndex & gc_strEXTRASDELIM & IIf(.blnGF_TDE, "1", "0") & gc_strEXTRASDELIM & .lngGF_Control & gc_strEXTRASDELIM & .lngGF_PalNTSC & gc_strEXTRASDELIM & .lngGF_Players & gc_strEXTRASDELIM & .strGF_Comment & gc_strEXTRASDELIM & .strGF_Notes & gc_strEXTRASDELIM & .strGF_VComment
            End If
        End If
    End With

End Function
Public Function GetExtrasPropsFromDataString(strData_in As String) As Long

    'Returns the items in the g_udtExtrasProps structure from
    'the data string stored in the Data field of the Extras table.
    'Returns the Gemus Type (use default, use current or use specific)

    Dim a_udtFields As ManualGEMUSFields
    Dim lngMaxItems As Long

    Dim a_strData() As String

    Call g_xobjGEMUS.GetManualFieldsDefaults(a_udtFields, g_strGameBaseName)

    If strData_in = "" Then
        With g_udtExtrasProps
            .blnGF_TDE = a_udtFields.blnTDE
            .lngGF_Control = a_udtFields.lngControl
            .lngGF_ImageIndex = 0
            .lngGF_PalNTSC = a_udtFields.lngPalNTSC
            .lngGF_Players = a_udtFields.lngPlayers
            .strGF_Comment = a_udtFields.strComment
            .strGF_Notes = a_udtFields.strNotes
            .strGF_VComment = a_udtFields.strVComment
        End With
        GetExtrasPropsFromDataString = 0
    Else
        a_strData() = Split(strData_in, gc_strEXTRASDELIM)
        lngMaxItems = UBound(a_strData)
        If lngMaxItems < 1 Then
            'oops - no data!
             With g_udtExtrasProps
                .blnGF_TDE = a_udtFields.blnTDE
                .lngGF_Control = a_udtFields.lngControl
                .lngGF_ImageIndex = 0
                .lngGF_PalNTSC = a_udtFields.lngPalNTSC
                .lngGF_Players = a_udtFields.lngPlayers
                .strGF_Comment = a_udtFields.strComment
                .strGF_Notes = a_udtFields.strNotes
                .strGF_VComment = a_udtFields.strVComment
            End With
            GetExtrasPropsFromDataString = 0
        ElseIf lngMaxItems < 2 Then
            g_udtExtrasProps.lngGF_ImageIndex = Val(a_strData(1))
            If a_strData(0) = "1" Then
                'use current
                With g_udtExtrasProps
                    .blnGF_TDE = a_udtFields.blnTDE
                    .lngGF_Control = g_udtGame.lngControl
                    .lngGF_PalNTSC = g_udtGame.lngNTSC
                    .lngGF_Players = g_udtGame.lngPlayersFrom
                    .strGF_Comment = g_udtGame.strComment
                    .strGF_Notes = g_udtGame.strPairs
                    .strGF_VComment = g_udtGame.strVersionComment
                End With
                GetExtrasPropsFromDataString = 1
            Else
                'use default
                With g_udtExtrasProps
                    .blnGF_TDE = a_udtFields.blnTDE
                    .lngGF_Control = a_udtFields.lngControl
                    .lngGF_PalNTSC = a_udtFields.lngPalNTSC
                    .lngGF_Players = a_udtFields.lngPlayers
                    .strGF_Comment = a_udtFields.strComment
                    .strGF_Notes = a_udtFields.strNotes
                    .strGF_VComment = a_udtFields.strVComment
                End With
                GetExtrasPropsFromDataString = 0
            End If
        Else
            'use specific!
            g_udtExtrasProps.lngGF_ImageIndex = Val(a_strData(1))
            With g_udtExtrasProps
                If lngMaxItems > 1 Then
                    .blnGF_TDE = (a_strData(2) = "1")
                Else
                    .blnGF_TDE = a_udtFields.blnTDE
                End If
                If lngMaxItems > 2 Then
                    .lngGF_Control = a_strData(3)
                Else
                    .lngGF_Control = a_udtFields.lngControl
                End If
                If lngMaxItems > 3 Then
                    .lngGF_PalNTSC = a_strData(4)
                Else
                    .lngGF_PalNTSC = a_udtFields.lngPalNTSC
                End If
                If lngMaxItems > 4 Then
                    .lngGF_Players = a_strData(5)
                Else
                    .lngGF_Players = a_udtFields.lngPlayers
                End If
                If lngMaxItems > 5 Then
                    .strGF_Comment = a_strData(6)
                Else
                    .strGF_Comment = ""
                End If
                If lngMaxItems > 6 Then
                    .strGF_Notes = a_strData(7)
                Else
                    .strGF_Notes = ""
                End If
                If lngMaxItems > 7 Then
                    .strGF_VComment = a_strData(8)
                Else
                    .strGF_VComment = ""
                End If
            End With
            GetExtrasPropsFromDataString = 2
        End If
    End If

End Function
Public Function GetExistingGamePath() As Long

    '------------------------------------------------------------------------------
    'returns -1 if no game path exists or the index to the first existing game path
    '------------------------------------------------------------------------------

    Dim lngPathIdx As Long  'loop index

    GetExistingGamePath = -1
    For lngPathIdx = 0 To (g_lngTotalGamePaths - 1)
        If g_xobjFunc.DirExists(ga_udtRegPathGame(lngPathIdx).strRealPath) = True Then
            GetExistingGamePath = lngPathIdx
            Exit Function
        End If
    Next lngPathIdx

End Function




Public Sub GetPathInfo(strPath_in As String, udtPath_inout As PathsType)

    '-------------------------------------------------
    'Fills the PathsType structure passed in.
    '
    'Inputs:    strPath_in:     The path to test
    '
    'Modify:    udtPath_inout:  The structure to fill.
    '-------------------------------------------------

    Dim strValidateLocal As String
    Dim iLen As Integer
    
    'has to have at least 3 letters ("c:\")
    With udtPath_inout
        .strRealPath = strPath_in
        iLen = Len(.strRealPath)
       
        If iLen > 2 Then
            'convert root "\" to current drive root
            If Left$(.strRealPath, 1) = "\" Then
                .strRealPath = Left$(g_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 1))
                iLen = Len(.strRealPath)
            End If
            
            'convert relative path to absolute
            If Mid$(.strRealPath, 2, 1) <> ":" Then
                .strRealPath = g_strAppPathWithBackslash & .strRealPath
                iLen = Len(.strRealPath)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If g_xobjFunc.DirExists(.strRealPath) = True Then
                    .strRealPath = CurDir()
                End If
            End If
                    
            'if specified path does not exist, try to validate on the current drive
            If g_xobjFunc.DirExists(.strRealPath) = False Then
                strValidateLocal = Left$(g_strAppPathWithBackslash, 3) & Right(.strRealPath, (iLen - 3))
                If g_xobjFunc.DirExists(strValidateLocal) = True Then
                    .strRealPath = strValidateLocal
                End If
            End If
        
            'set path and path with backslash
            .strWithBackSlash = .strRealPath
            If Right$(.strWithBackSlash, 1) <> "\" Then .strWithBackSlash = .strWithBackSlash & "\"
            
            'test drive type
            .ePathType = GetTypeOfDrive(Left$(.strRealPath, 3))
            If .ePathType = e_PathLocation_FixedHD Then
                'fixed drive - make sure path exists
                If g_xobjFunc.DirExists(.strRealPath) = False Then
                    .ePathType = e_PathLocation_Invalid
                End If
            End If
        Else
            .ePathType = e_PathLocation_Invalid
        End If
    End With
    
End Sub
Public Function GetSidCommandLine(strCLPTemplate_in As String, ByVal strSidFile_in As String, blnUseShort_in As Boolean) As String

    '-----------------------------------------------------------------
    'gets the actual command line to use when shelling a sid player
    '
    'Inputs:    strCLPTemplate_in:  The template command line
    '           strSidFile_in:      The full sid filename to replace
    '           blnUseShort_in:     whether to use short 8.3 filenames
    '
    'Returns:   The command line parameter string to shell
    '-----------------------------------------------------------------

    If Not blnUseShort_in Then
        If InStr(strSidFile_in, " ") > 0 Then
            'put quotes around because of spaces
            GetSidCommandLine = Replace(strCLPTemplate_in, "%musicfile%", """" & strSidFile_in & """", , , vbTextCompare)
        Else
            'no spaces - full filename
            GetSidCommandLine = Replace(strCLPTemplate_in, "%musicfile%", strSidFile_in, , , vbTextCompare)
        End If
    Else
        '8.3 filename
        GetSidCommandLine = Replace(strCLPTemplate_in, "%musicfile%", g_xobjFunc.Long2Short(strSidFile_in), , , vbTextCompare)
    End If

End Function

Public Sub GetSidPlayerInfo(strCLP_in As String, ByVal blnUseShort_in As Boolean, strName_in As String, strPath_in As String, udtSidPlayer_inout As SidPlayerType, a_strTypes_in() As String, ByVal lngTotalTypes_in As Long)

    '--------------------------------------------
    'Fills the SidPlayerType structure passed in.
    '--------------------------------------------

    Dim strValidateLocal As String
    Dim iLen As Integer

    'has to have at least 3 letters ("c:\")
    With udtSidPlayer_inout
        .strName = strName_in
        .blnUseShortFilenames = blnUseShort_in
        .strCommandLine = strCLP_in
        .strFullPathAndFilename = strPath_in
        .lngTotalTypes = lngTotalTypes_in
        .a_strTypes() = a_strTypes_in()
        iLen = Len(.strFullPathAndFilename)
        
        If iLen > 2 Then
            Call g_xobjFunc.GetFileAndPath(.strFullPathAndFilename, .strFullPathOnly, .strFullFilenameOnly)
            iLen = Len(.strFullPathOnly)
            
            'convert root "\" to current drive root
            If Left$(.strFullPathOnly, 1) = "\" Then
                .strFullPathOnly = Left$(g_strAppPathWithBackslash, 3) & Right(.strFullPathOnly, (iLen - 1))
                iLen = Len(.strFullPathOnly)
            End If
        
            'convert relative path to absolute
            If Mid$(.strFullPathOnly, 2, 1) <> ":" Then
                .strFullPathOnly = g_strAppPathWithBackslash & .strFullPathOnly
                iLen = Len(.strFullPathOnly)
                'let filesystem convert relative path to qualified directory name (remove "..\", etc.)
                If g_xobjFunc.DirExists(.strFullPathOnly) = True Then
                    .strFullPathOnly = CurDir()
                End If
            End If
            
            'if specified path does not exist, try to validate on the current drive
            If g_xobjFunc.DirExists(.strFullPathOnly) = False Then
                strValidateLocal = Left$(g_strAppPathWithBackslash, 3) & Right(.strFullPathOnly, (iLen - 3))
                If g_xobjFunc.DirExists(strValidateLocal) = True Then
                    .strFullPathOnly = strValidateLocal
                End If
            End If
            
            .strFullPathAndFilename = .strFullPathOnly & "\" & .strFullFilenameOnly
        
            'test drive type
            .ePathType = GetTypeOfDrive(Left$(.strFullPathAndFilename, 3))
            If .ePathType <> e_PathLocation_Invalid Then
                'check if path exists
                If g_xobjFunc.FileExists(.strFullPathAndFilename, False, "") = False Then
                    If .ePathType = e_PathLocation_FixedHD Then
                        'file doesn't exist on HD
                        .ePathType = e_PathLocation_Invalid
                    End If
                Else
                    'file found - get info
                    .strShortPathAndFilename = g_xobjFunc.Long2Short(.strFullPathAndFilename)
                    Call g_xobjFunc.GetFileAndPath(.strFullPathAndFilename, .strFullPathOnly, .strFullFilenameOnly)
                    Call g_xobjFunc.GetFileAndPath(.strShortPathAndFilename, .strShortPathOnly, .strShortFilenameOnly)
                    .strFullPathOnlyWithBackSlash = .strFullPathOnly
                    .strShortPathOnlyWithBackSlash = .strShortPathOnly
                    If Right$(.strFullPathOnlyWithBackSlash, 1) <> "\" Then .strFullPathOnlyWithBackSlash = .strFullPathOnlyWithBackSlash & "\"
                    If Right$(.strShortPathOnlyWithBackSlash, 1) <> "\" Then .strShortPathOnlyWithBackSlash = .strShortPathOnlyWithBackSlash & "\"
                End If
            End If
        Else
            'invalid path\filename string
            .ePathType = e_PathLocation_Invalid
        End If
    End With

End Sub
Public Sub PlaySidWithSpecificPlayer(ByVal lngPlayerIdx_in As Long)

    '-----------------------------------------------------
    'plays the selected sid tune with the specified player
    '
    'Inputs:    lngPlayerIdx_in:    The player to use
    '-----------------------------------------------------

    Dim dblProcessID As Double  'process ID of shelled player
    Dim lngTotalPlayed As Long  'for stats
    Dim lngHandle As Long

    On Error GoTo Err_Handler

    If g_lngTotalSidPlayers > (lngPlayerIdx_in - 1) Then
        With ga_udtSidPlayers(lngPlayerIdx_in)
            If g_xobjFunc.FileExists(.strFullPathAndFilename, False, "") Then
                'now shell the player
                If Not .blnUseShortFilenames Then
                    dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, False), vbNormalFocus)
                Else
                    dblProcessID = Shell(.strShortPathAndFilename & " " & GetSidCommandLine(.strCommandLine, g_udtRegPathSid.strWithBackSlash & g_udtSid.strFilename, True), vbNormalFocus)
                End If
                'store the process handle
                lngHandle = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                Call AddToSidPlayerHandles(lngHandle)
                'modify stats
                g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed, False, 0
                g_xobjRegistry.SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed + 1, False
                'add to the recent list?
                If (g_eWhatIsSelected = e_SelType_IsGame) And (g_blnRecentGameSid = True) Then
                    If (g_lngRecentUpdate = gc_lngRECENT_PLAYED) And (g_lngRecentStore <> gc_lngRECENT_SIDS) Then
                        Call frmMain.AddToRecentList("G" & g_udtGame.lngID, g_udtGame.strName)
                    End If
                ElseIf g_eWhatIsSelected = e_SelType_IsSid Then
                    If (g_lngRecentUpdate = gc_lngRECENT_PLAYED) And (g_lngRecentStore <> gc_lngRECENT_GAMES) Then
                        Call frmMain.AddToRecentList("M" & g_udtSid.lngID, g_udtSid.strName)
                    End If
                End If
            Else
                MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_MusicPlayerNotFound", "The Music Player you selected was not found!"), vbExclamation, gc_strAPPNAME
            End If
        End With
    End If
    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbExclamation, gc_strAPPNAME

End Sub

Public Sub ResetTips(ByVal blnShowThem_in As Boolean)

    '-------------------------------------------------
    'Resets the registry values for the tips to either
    'never show them or to show all of them again
    '
    'Inputs:    blnShowThem_in:     Flag specifying
    '                               whether to show or
    '                               not show the tips
    '-------------------------------------------------
    
    Dim lngTipIdx As Long       'loop index
    
    'reset the tips
    With g_xobjRegistry
        'main tips
        For lngTipIdx = 0 To gc_lngMAXTIPS
            g_udtMainTips(lngTipIdx).blnShowIt = blnShowThem_in
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Tips", g_udtMainTips(lngTipIdx).strINIAndRegKey, blnShowThem_in, False
        Next lngTipIdx
    End With

End Sub
Public Sub ReleaseSidPlayers(ByVal blnTerminate_in As Boolean)

    '------------------------------------------------------
    'closes all sidplayer instances shelled by this program
    'and releases the handles to them
    '
    'Inputs:    blnTerminate_in:    terminate flag
    '------------------------------------------------------

    Dim lngHandleIdx As Long    'loop index

    If g_lngTotalSidPlayerHandles > 0 Then
        'test each player
        For lngHandleIdx = 0 To (g_lngTotalSidPlayerHandles - 1)
            If ga_lngSidPlayerHandles(lngHandleIdx) <> 0 Then
                'close it
                If blnTerminate_in Then
                    Call TerminateProcess(ga_lngSidPlayerHandles(lngHandleIdx), 0&)
                End If
                Call CloseHandle(ga_lngSidPlayerHandles(lngHandleIdx))
            End If
        Next lngHandleIdx
    End If
    g_lngTotalSidPlayerHandles = 0
    ReDim ga_lngSidPlayerHandles(0)

End Sub


Public Sub UnloadAllForms()

    On Error GoTo Err_Handler

    If g_blnLoadedfrmBigScreenshot Then Unload frmBigScreenshot
    If g_blnLoadedfrmCategories Then Unload frmCategories
    If g_blnLoadedfrmNotes Then Unload frmNotes
    If g_blnLoadedfrmSocket Then Unload frmSocket
    g_blnLoadedfrmBigScreenshot = False
    g_blnLoadedfrmCategories = False
    g_blnLoadedfrmCategories = False
    g_blnLoadedfrmSocket = False
    Exit Sub

Err_Handler:
    MsgBox Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")"
    End

End Sub


Public Function ValidFilename(strFilename_in As String) As Boolean

    ValidFilename = _
    (InStr(strFilename_in, "\") = 0) And _
    (InStr(strFilename_in, "/") = 0) And _
    (InStr(strFilename_in, ":") = 0) And _
    (InStr(strFilename_in, "*") = 0) And _
    (InStr(strFilename_in, "?") = 0) And _
    (InStr(strFilename_in, """") = 0) And _
    (InStr(strFilename_in, "<") = 0) And _
    (InStr(strFilename_in, ">") = 0) And _
    (InStr(strFilename_in, "|") = 0) And _
    (InStr(strFilename_in, ".") > 1)

End Function


