VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScrnShotSelector 
   Caption         =   "Picture Selector"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmScrnShotSelector.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9015
   Begin VB.FileListBox FileBox 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   4335
      Begin VB.CommandButton cmdARF 
         Caption         =   "&Auto Rename"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1380
         Width           =   1935
      End
      Begin VB.CommandButton cmdBMP2PNG 
         Caption         =   "BMP-->&PNG"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   120
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   136
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CommandButton cmdBMP2GIF 
         Caption         =   "BMP-->&GIF"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label lblPreview 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   4095
      End
   End
   Begin VB.Frame fraPicturePaths 
      Caption         =   "Picture Paths"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdRefreshPics 
         Caption         =   "R&efresh"
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "&Apply to all games in current view"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   4095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   3960
         Width           =   1095
      End
      Begin VB.DirListBox DirPics 
         Height          =   990
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox lstPaths 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin MSComctlLib.ListView lvPics 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1191
         EndProperty
      End
      Begin VB.Label lblLinkPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Picture:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label lblLink 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         UseMnemonic     =   0   'False
         Width           =   4095
      End
   End
   Begin VB.Frame fraFindPictures 
      Caption         =   "Find Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdRefreshFind 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView lvFind 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1191
         EndProperty
      End
      Begin VB.DirListBox DirFind 
         Height          =   990
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4095
      End
      Begin VB.DriveListBox DriveFind 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   21
      Top             =   5160
      Width           =   1095
   End
   Begin VB.PictureBox picFull 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5760
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmScrnShotSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private m_blnByCode As Boolean
Private m_strIncompatibleScreenshot As String
Private m_strEndsInX As String
Private m_strLastPath As String
Private m_strLastFindPath As String
Private m_lngFocusIdx As Long
Private m_blnDoneActivate As Boolean
Private m_strPreviewFilename As String
Private m_strBMP2GIFPath As String
Private m_strBMP2PNGPath As String
Private m_strOldFilename As String
Private m_strDraggedFile As String
Private m_strARF As String
Private m_blnARF As Boolean

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmpFileHeader As BITMAPFILEHEADER
        bmiHeader As BITMAPINFOHEADER
End Type
Private Const BI_RGB = 0&

Private Function GetValidAutoRenameFilename(strGameName_in As String, strPicName_out As String) As Boolean

    Dim lngIdx As Long
    Dim strChar As String
    Dim lngLen As Long

    'remove all invalid characters
    strPicName_out = ""
    For lngIdx = 1 To Len(strGameName_in)
        strChar = Mid$(strGameName_in, lngIdx, 1)
        If InStr("\/:*?""<>|", strChar) = 0 Then
            strPicName_out = strPicName_out & strChar
        End If
    Next lngIdx

    'replace spaces with underscores
    strPicName_out = Replace(strPicName_out, " ", "_")

    'ok to use?
    lngLen = Len(strPicName_out)
    If lngLen > 0 Then
        strPicName_out = Left$(strPicName_out, 57)
        GetValidAutoRenameFilename = True
    Else
        GetValidAutoRenameFilename = False
    End If

End Function
Private Function CheckBitmap(strScrnShot_in As String, ByVal blnLimit256_in As Boolean) As Boolean

    '--------------------------------------------------------------------
    'checks we have a 4-bit, 8-bit or 16-bit uncompressed bitmap
    '
    'Inputs:    strScrnShot_in:     The path\filename of the BMP to check
    '
    'Outputs:   blnLimit256_in:     Do we accept only 256 cols max?
    '
    'Returns:   True or False if the bmp is accepted
    '--------------------------------------------------------------------

    Dim intFileNum As Integer           'file number
    Dim udtBMPHeader As BITMAPINFO      'bitmap header structure

    On Error GoTo Err_Handler
    
    'read the BMP header
    intFileNum = FreeFile()
    Open strScrnShot_in For Binary As #intFileNum
    Get #intFileNum, , udtBMPHeader

    'check the bitmap is a valid windows bitmap
    If udtBMPHeader.bmpFileHeader.bfType <> 19778 Then
        'BMP is not a valid windows BMP
        CheckBitmap = False
        Close #intFileNum
        Exit Function
    End If
    With udtBMPHeader.bmiHeader
        'check the bitmap is not compressed
        If .biCompression <> BI_RGB Then
            'BMP is compressed
            CheckBitmap = False
            Close #intFileNum
            Exit Function
        End If
        'check we have a 16 or 256 colour bitmap (4-bit or 8 bit) if necessary
        If blnLimit256_in = True Then
            If (.biBitCount <> 4) And (.biBitCount <> 8) Then
                'BMP wrong bpp
                Close #intFileNum
                CheckBitmap = False
                Exit Function
            End If
        End If
    End With
    
    Close #intFileNum
    CheckBitmap = True
    Exit Function

Err_Handler:
    CheckBitmap = False

End Function





Private Sub ClickLVPics(Item As ListItem)

    Dim strTest As String
    Dim lngPos As Long
    Dim blnTest As Boolean

    'are we allowed to select this image?
    blnTest = True
    strTest = Left$(Item.Text, Len(Item.Text) - 4)
    lngPos = InStrRev(strTest, "_")
    If lngPos > 1 And lngPos < Len(strTest) Then
        blnTest = False
        strTest = Right$(strTest, Len(strTest) - lngPos)
        For lngPos = 1 To Len(strTest)
            If InStr("1234567890", Mid$(strTest, lngPos, 1)) = 0 Then
                blnTest = True
                Exit For
            End If
        Next lngPos
    End If
    strTest = DirPics.Path
    If Right$(strTest, 1) <> "\" Then
        strTest = strTest & "\"
    End If
    strTest = strTest & Item.Text
    'If blnTest = True Then
        lngPos = Len(lstPaths.List(lstPaths.ListIndex))
        If Right$(lstPaths.List(lstPaths.ListIndex), 1) <> "\" Then
            lngPos = lngPos + 1
        End If
        lblLink.FontBold = False
        lblLink.Caption = Right$(strTest, Len(strTest) - lngPos)
        cmdOK.Enabled = True
    'Else
    '    lblLink.FontBold = True
    '    lblLink.Caption = m_strEndsInX
    '    cmdOK.Enabled = False
    'End If

    cmdARF.Enabled = m_blnARF

    'show the screenshot
    Call ShowPreview(strTest)

End Sub
Private Sub ClickLVFind(Item As ListItem)

    Dim strTest As String

    strTest = DirFind.Path
    If Right$(strTest, 1) <> "\" Then
        strTest = strTest & "\"
    End If
    strTest = strTest & Item.Text
    
    cmdARF.Enabled = False
    
    Call ShowPreview(strTest)

End Sub
Private Function CopyReplaceFile(strSourceFile_in As String, strDestDir_in As String) As Boolean

    'copies strSourceFile_in into the folder strDestDir_in (same filename)

    Dim strDestFile As String
    Dim strFile As String
    Dim strPath As String

    On Error GoTo NoCopy

    'get dest filename
    Call g_xobjFunc.GetFileAndPath(strSourceFile_in, strPath, strFile)
    strDestFile = strDestDir_in
    If Right$(strDestFile, 1) <> "\" Then strDestFile = strDestFile & "\"
    strDestFile = strDestFile & strFile
    'exists in the dest folder?
    If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
        If MsgBox(g_xobjFunc.GetINIMessage("Picture Selector", "DragDropReplaceFile", "The file you are dropping here exists in this folder.  Do you want to overwrite it?"), vbYesNoCancel Or vbQuestion Or vbDefaultButton2, gc_strAPPNAME) = vbYes Then
            'want to overwrite - kill existing filename
            If DeleteListViewFile(strDestFile) = True Then
                'existing file deleted - just copy
                Call FileCopy(strSourceFile_in, strDestFile)
                CopyReplaceFile = True
            Else
                'error deleting existing file
                CopyReplaceFile = False
            End If
        Else
            'no wish to overwrite
            CopyReplaceFile = False
        End If
    Else
        'just copy
        Call FileCopy(strSourceFile_in, strDestFile)
        CopyReplaceFile = True
    End If

    Exit Function

NoCopy:
    MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "CouldNotCopyMoveFile", "Could not copy or move the file.") & "  " & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME
    CopyReplaceFile = False

End Function

Private Function DeleteListViewFile(strFileToDelete_in As String) As Boolean

    On Error GoTo NoDelete

    Call Kill(strFileToDelete_in)
    DeleteListViewFile = True
    Exit Function

NoDelete:
    MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "CouldNotDeleteFile", "Could not delete file.") & "  " & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME
    DeleteListViewFile = False

End Function

Private Sub RefreshListsAfterConversion(strOldFilename_in As String, strNewFilename_in As String)

    Dim strUpper As String
    Dim lvToCheck As ListView
    Dim LItem As ListItem
    Dim blnPics As Boolean
    Dim strExt As String

    strExt = UCase$(Right$(strNewFilename_in, 3))

    'which listview has the file selected?
    blnPics = False
    strUpper = UCase$(strOldFilename_in)
    Set lvToCheck = Nothing
    If Not (lvFind.SelectedItem Is Nothing) Then
        If UCase$(lvFind.SelectedItem.Text) = strUpper Then
            lvFind.SelectedItem.Text = strNewFilename_in
            lvFind.SelectedItem.SubItems(1) = strExt
            Call ClickLVFind(lvFind.SelectedItem)
            Set lvToCheck = lvPics
            blnPics = True
        End If
    ElseIf Not (lvPics.SelectedItem Is Nothing) Then
        If UCase$(lvPics.SelectedItem.Text) = strUpper Then
            lvPics.SelectedItem.Text = strNewFilename_in
            lvPics.SelectedItem.SubItems(1) = strExt
            Call ClickLVPics(lvPics.SelectedItem)
            Set lvToCheck = lvFind
        End If
    End If

    'check the other listview
    If Not (lvToCheck Is Nothing) Then
        If UCase$(DirFind.Path) = UCase$(DirPics.Path) Then
            For Each LItem In lvToCheck.ListItems
                If UCase$(LItem.Text) = strUpper Then
                    'found it, replace and quit the loop
                    LItem.Text = strNewFilename_in
                    LItem.SubItems(1) = strExt
                    If LItem = lvToCheck.SelectedItem Then
                        If blnPics = True Then
                            Call ClickLVPics(lvPics.SelectedItem)
                        End If
                    End If
                    Exit For
                End If
            Next LItem
        End If
        Set lvToCheck = Nothing
    End If

End Sub

Private Sub RefreshListView(lvFiles_in As ListView, strFolder_in As String)

    Dim lngIdx As Long
    Dim lngMaxFiles As Long
    Dim LItem As ListItem

    're-fill the passed-in listview with files in that folder
    Screen.MousePointer = vbHourglass
    DoEvents
    With FileBox
        .Pattern = "*.gif;*.jpg;*.bmp;*.png"
        .Path = strFolder_in
        .Refresh
        DoEvents
    End With
    With lvFiles_in
        .Enabled = False
        .Visible = False
        .ListItems.Clear
        lngMaxFiles = FileBox.ListCount - 1
        With .ListItems
            For lngIdx = 0 To lngMaxFiles
                Set LItem = .Add(, , FileBox.List(lngIdx))
                LItem.SubItems(1) = UCase$(Right$(LItem.Text, 3))
            Next lngIdx
        End With
        .Enabled = True
        .Visible = True
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub RemoveFileFromListViews(strFilename_in As String, strPath_in As String)

    'removes a file from both listviews if it exists in them

    Dim LItem As ListItem
    Dim strUpper As String

    strUpper = UCase$(strFilename_in)
    If UCase$(strPath_in) = UCase$(g_xobjFunc.Long2Short(DirFind.Path)) Then
        For Each LItem In lvFind.ListItems
            If UCase$(LItem.Text) = strUpper Then
                lvFind.ListItems.Remove LItem.Index
                Exit For
            End If
        Next LItem
    End If
    If UCase$(strPath_in) = UCase$(g_xobjFunc.Long2Short(DirPics.Path)) Then
        For Each LItem In lvPics.ListItems
            If UCase$(LItem.Text) = strUpper Then
                lvPics.ListItems.Remove LItem.Index
                Exit For
            End If
        Next LItem
    End If
    
End Sub

Private Function RenameListViewFile(strOldFile_in As String, strNewFile_in As String) As Boolean

    Dim strPath As String
    Dim strFile As String

    On Error GoTo Err_Handler

    'compare with old filename
    If UCase$(strNewFile_in) = UCase$(strOldFile_in) Then
        'filenames match - rename if different case
        If strNewFile_in <> strOldFile_in Then
            Name (strOldFile_in) As (strNewFile_in)
            RenameListViewFile = True
        Else
            RenameListViewFile = False
        End If
    Else
        Call g_xobjFunc.GetFileAndPath(strNewFile_in, strPath, strFile)
        If (Right$(UCase$(strNewFile_in), 4) = Right$(UCase$(strOldFile_in), 4)) And (ValidFilename(strFile) = True) Then
            If g_xobjFunc.FileExists(strNewFile_in, False, gc_strAPPNAME) = False Then
                Name (strOldFile_in) As (strNewFile_in)
                RenameListViewFile = True
            Else
                'error - file already exists
                MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "CannotRename", "Cannot rename. File already exists!"), vbExclamation, gc_strAPPNAME
                RenameListViewFile = False
            End If
        Else
            MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "InvalidFilename", "Invalid Filename!"), vbExclamation, gc_strAPPNAME
            RenameListViewFile = False
        End If
    End If

    Exit Function

Err_Handler:
    MsgBox "Could not rename file.  " & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME
    RenameListViewFile = False

End Function

Private Function SelectFileInListView(lvFiles_in As ListView, strFile_in As String) As Boolean

    Dim LItem As ListItem
    Dim strFile As String

    'find and select the requested file in the specified listview

    strFile = UCase$(strFile_in)
    With lvFiles_in
        For Each LItem In .ListItems
            If strFile = UCase$(LItem.Text) Then
                'found it!
                LItem.Selected = True
                Set .SelectedItem = LItem
                LItem.EnsureVisible
                lvPics_ItemClick lvPics.SelectedItem
                SelectFileInListView = True
                Exit Function
            End If
        Next LItem
    End With

    'not in list
    SelectFileInListView = False

End Function

Private Sub ShowPreview(strFilename_in As String)

    Dim strType As String
    Dim udtSnug As SnugFitType

    On Error GoTo Bad_Image

    'get the image type
    strType = UCase$(Right$(strFilename_in, 3))

    'load into picFull at original size
    Select Case strType
        Case "PNG"
            picFull.Cls
            picFull.Picture = LoadPicture
            picFull.AutoSize = False
            png_load strFilename_in
            picFull.Width = png_width * Screen.TwipsPerPixelX
            picFull.Height = png_height * Screen.TwipsPerPixelY
            png_draw picFull.hDC, 0, 0
            picFull.Picture = picFull.Image
            cmdBMP2GIF.Enabled = False
        Case Else
            picFull.AutoSize = True
            picFull.Picture = LoadPicture(strFilename_in)
    End Select
    
    'calculate the new size and blit to picPreview
    With udtSnug
        .lngDestH = picPreview.ScaleHeight
        .lngDestW = picPreview.ScaleWidth
        .lngSourceH = picFull.ScaleHeight
        .lngSourceW = picFull.ScaleWidth
        Call GetSnugFitImageDimensions(udtSnug)
        picPreview.Picture = LoadPicture
        picPreview.PaintPicture picFull.Picture, .lngBlitX, .lngBlitY, .lngBlitW, .lngBlitH, 0, 0, picFull.ScaleWidth, picFull.ScaleHeight
    End With
    picPreview.Visible = True
    m_strPreviewFilename = strFilename_in
    lblPreview.Caption = EllipseText(strFilename_in, lblPreview.Width)

    'tool buttons
    If strType = "BMP" Then
        cmdBMP2GIF.Enabled = (m_strBMP2GIFPath <> "") And (CheckBitmap(strFilename_in, True))
        cmdBMP2PNG.Enabled = (m_strBMP2PNGPath <> "") And (CheckBitmap(strFilename_in, False))
    Else
        cmdBMP2GIF.Enabled = False
        cmdBMP2PNG.Enabled = False
    End If

    Exit Sub

Bad_Image:
    picPreview.Visible = False
    m_strPreviewFilename = ""
    lblPreview.Caption = ""
    cmdBMP2PNG.Enabled = False
    cmdBMP2GIF.Enabled = False
    cmdARF.Enabled = False
    cmdOK.Enabled = False

End Sub


Private Function CannotWriteToPath(strPathWBS_in As String) As Boolean

    Dim intFileNum As Integer
    Dim strFilename As String

    On Error GoTo NoWrite
    
    CannotWriteToPath = False
    strFilename = strPathWBS_in & "gbtestwrite.fil"
    If g_xobjFunc.FileExists(strFilename, False, "") Then
        Kill strFilename
    End If
    intFileNum = FreeFile()
    Open strFilename For Output As #intFileNum
    Print #intFileNum, "."
    Close #intFileNum
    Kill strFilename
    Exit Function
    
NoWrite:
    CannotWriteToPath = True

End Function


Private Sub cmdARF_Click()

    'AUTO RENAME FILENAME

    Dim strNewFilename As String
    Dim LItem As ListItem
    Dim strUpper As String
    Dim strOldFilename As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathWBS As String
    Dim lngIdx As Long
    Dim strExt As String
    Dim blnFound As Boolean
    Dim strNewPath As String
    Dim strNewFile As String
    Dim ITM As ListItem

    'initial checks
    strOldFilename = m_strPreviewFilename
    If m_blnARF = False Then
        cmdARF.Enabled = False
        Exit Sub
    ElseIf g_xobjFunc.FileExists(strOldFilename, False, gc_strAPPNAME) = False Then
        cmdARF.Enabled = False
        Exit Sub
    ElseIf Len(m_strARF) = 0 Then
        cmdARF.Enabled = False
        Exit Sub
    End If
    cmdARF.Enabled = False
    DoEvents
    
    'get filename extension
    strExt = Right$(strOldFilename, 3)

    'get the full new filename
    Call g_xobjFunc.GetFileAndPath(strOldFilename, strPath, strFile)
    strUpper = UCase$(strFile)

    Set ITM = lvPics.SelectedItem
    If ITM Is Nothing Then
        cmdARF.Enabled = False
        Exit Sub
    Else
        If UCase$(ITM.Text) <> strUpper Then
            cmdARF.Enabled = False
            Exit Sub
        End If
    End If

    strPathWBS = strPath
    If Right$(strPathWBS, 1) <> "\" Then strPathWBS = strPathWBS & "\"
    strNewFilename = strPathWBS & m_strARF
    blnFound = False
    If g_xobjFunc.FileExists(strNewFilename & "." & strExt, False, "") = False Then
        'rename to first screenshot
        strNewFilename = strNewFilename & "." & strExt
        blnFound = True
    Else
        'find first free screenshot filename
        MousePointer = vbHourglass
        DoEvents
        For lngIdx = 1 To 99
            If g_xobjFunc.FileExists(strNewFilename & "_" & Trim$(lngIdx) & "." & strExt, False, "") = False Then
                'rename to this extra screenshot
                strNewFilename = strNewFilename & "_" & Trim$(lngIdx) & "." & strExt
                blnFound = True
                Exit For
            End If
        Next lngIdx
    End If

    If blnFound = True Then
        'try to rename it
        If RenameListViewFile(strOldFilename, strNewFilename) = True Then
            MousePointer = vbHourglass
            DoEvents
            'renamed - rename the file in the lists
            Call g_xobjFunc.GetFileAndPath(strNewFilename, strNewPath, strNewFile)
            lvPics.SelectedItem.Text = strNewFile
            If UCase$(DirFind.Path) = UCase$(strPath) Then
                For Each LItem In lvFind.ListItems
                    If UCase$(LItem.Text) = strUpper Then
                        LItem.Text = strNewFile
                        Exit For
                    End If
                Next LItem
            End If
            'update the preview label
            m_strPreviewFilename = strNewFilename
            lblPreview.Caption = EllipseText(strNewFilename, lblPreview.Width)
            're-select the renamed file in the right listview window
            lvPics_ItemClick lvPics.SelectedItem
        End If
    End If
    
    MousePointer = vbDefault

End Sub

Private Sub cmdBMP2GIF_Click()

    Dim strCommandLine As String
    Dim strPath As String
    Dim strFile As String
    Dim dblProcessID As Double
    Dim hProcess As Long
    Dim lngQuitTime&
    Dim strDestFile As String
    Dim strNewFile As String

    MousePointer = vbHourglass
    cmdBMP2GIF.Enabled = False
    DoEvents

    If m_strPreviewFilename <> "" And m_strBMP2GIFPath <> "" Then
        If g_xobjFunc.FileExists(m_strPreviewFilename, False, gc_strAPPNAME) = True Then
            If g_xobjFunc.FileExists(m_strBMP2GIFPath, False, gc_strAPPNAME) = True Then
                'picture file and command line tool exist
                strDestFile = g_xobjFunc.Long2Short(m_strPreviewFilename)
                strDestFile = Left$(strDestFile, Len(strDestFile) - 3) & "gif"
                If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                    If MsgBox(g_xobjFunc.GetINIMessage("Picture Selector", "ConvertReplaceFile", "Do you want to overwrite the following file?") & vbCrLf & vbCrLf & "'" & strDestFile & "'", vbYesNoCancel Or vbQuestion Or vbDefaultButton2, gc_strAPPNAME) <> vbYes Then
                        MousePointer = vbDefault
                        cmdBMP2GIF.Enabled = True
                        Exit Sub
                    Else
                        If DeleteListViewFile(strDestFile) = False Then
                            MousePointer = vbDefault
                            cmdBMP2GIF.Enabled = True
                            Exit Sub
                        Else
                            Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strFile)
                            Call RemoveFileFromListViews(strFile, strPath)
                        End If
                    End If
                End If
                ChDrive Left$(m_strBMP2GIFPath, 3)
                Call g_xobjFunc.GetFileAndPath(m_strBMP2GIFPath, strPath, strFile)
                ChDir strPath
                strCommandLine = g_xobjFunc.Long2Short(Environ$("COMSPEC")) & " /C "
                strCommandLine = strCommandLine & m_strBMP2GIFPath & " " & g_xobjFunc.Long2Short(m_strPreviewFilename) & " " & strDestFile
                dblProcessID = Shell(strCommandLine, vbHide)
                hProcess = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                Call WaitForSingleObject(hProcess, INFINITE)
                Call CloseHandle(hProcess)
                lngQuitTime& = GetTickCount() + 10000
                Do Until (GetTickCount() > lngQuitTime&) Or (g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True)
                Loop
                If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                    'kill old bmp file
                    Call DeleteListViewFile(m_strPreviewFilename)
                    'rename to full filename from short filename?
                    If UCase$(strDestFile) <> UCase$(Left$(m_strPreviewFilename, Len(m_strPreviewFilename) - 3) & "gif") Then
                        Name (strDestFile) As (Left$(m_strPreviewFilename, Len(m_strPreviewFilename) - 3) & "gif")
                    End If
                    strDestFile = Left$(m_strPreviewFilename, Len(m_strPreviewFilename) - 3) & "gif"
                    'wait a sec for image to finish writing to disk
                    lngQuitTime& = GetTickCount() + 750
                    Do Until (GetTickCount() > lngQuitTime&)
                    Loop
                    'new gif file found - which list to change? and do we refresh the other?
                    Call g_xobjFunc.GetFileAndPath(m_strPreviewFilename, strPath, strFile)
                    Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strNewFile)
                    Call RefreshListsAfterConversion(strFile, strNewFile)
                Else
                    MsgBox strDestFile & " not found! Conversion failed!", vbExclamation, gc_strAPPNAME
                    'check one more time to make sure file lists are properly updated
                    If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                        Call g_xobjFunc.GetFileAndPath(m_strPreviewFilename, strPath, strFile)
                        Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strNewFile)
                        Call RefreshListsAfterConversion(strFile, strNewFile)
                    End If
                End If
            Else
                MsgBox m_strBMP2GIFPath & " not found!", vbExclamation, gc_strAPPNAME
                cmdBMP2GIF.Enabled = False
            End If
        Else
            MsgBox m_strPreviewFilename & " not found!", vbExclamation, gc_strAPPNAME
            cmdBMP2GIF.Enabled = False
        End If
    End If

    MousePointer = vbDefault

End Sub
Private Sub cmdBMP2PNG_Click()

    Dim strCommandLine As String
    Dim strPath As String
    Dim strFile As String
    Dim dblProcessID As Double
    Dim hProcess As Long
    Dim lngQuitTime&
    Dim strDestFile As String
    Dim strNewFile As String

    Screen.MousePointer = vbHourglass
    cmdBMP2PNG.Enabled = False
    DoEvents

    If m_strPreviewFilename <> "" And m_strBMP2PNGPath <> "" Then
        If g_xobjFunc.FileExists(m_strPreviewFilename, False, gc_strAPPNAME) = True Then
            If g_xobjFunc.FileExists(m_strBMP2PNGPath, False, gc_strAPPNAME) = True Then
                'picture file and command line tool exist
                strDestFile = Left$(m_strPreviewFilename, Len(m_strPreviewFilename) - 3) & "png"
                If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                    If MsgBox(g_xobjFunc.GetINIMessage("Picture Selector", "ConvertReplaceFile", "Do you want to overwrite the following file?") & vbCrLf & vbCrLf & "'" & strDestFile & "'", vbYesNoCancel Or vbQuestion Or vbDefaultButton2, gc_strAPPNAME) <> vbYes Then
                        MousePointer = vbDefault
                        cmdBMP2PNG.Enabled = True
                        Exit Sub
                    Else
                        If DeleteListViewFile(strDestFile) = False Then
                            MousePointer = vbDefault
                            cmdBMP2PNG.Enabled = True
                            Exit Sub
                        Else
                            Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strFile)
                            strPath = g_xobjFunc.Long2Short(strPath)
                            Call RemoveFileFromListViews(strFile, strPath)
                        End If
                    End If
                End If
                ChDrive Left$(m_strBMP2PNGPath, 3)
                Call g_xobjFunc.GetFileAndPath(m_strBMP2PNGPath, strPath, strFile)
                ChDir strPath
                strCommandLine = g_xobjFunc.Long2Short(Environ$("COMSPEC")) & " /C "
                strCommandLine = strCommandLine & m_strBMP2PNGPath & " -Q -E """ & m_strPreviewFilename & """"
                dblProcessID = Shell(strCommandLine, vbHide)
                hProcess = OpenProcess(PROCESS_TERMINATE, False, dblProcessID)
                Call WaitForSingleObject(hProcess, INFINITE)
                Call CloseHandle(hProcess)
                lngQuitTime& = GetTickCount() + 10000
                Do Until (GetTickCount() > lngQuitTime&) Or (g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True)
                Loop
                If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                    'wait a sec for image to finish writing to disk
                    lngQuitTime& = GetTickCount() + 750
                    Do Until (GetTickCount() > lngQuitTime&)
                    Loop
                    'new png file found - which list to change? and do we refresh the other?
                    Call g_xobjFunc.GetFileAndPath(m_strPreviewFilename, strPath, strFile)
                    Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strNewFile)
                    Call RefreshListsAfterConversion(strFile, strNewFile)
                Else
                    MsgBox strDestFile & " not found! Conversion failed!", vbExclamation, gc_strAPPNAME
                    'check one more time to make sure file lists are properly updated
                    If g_xobjFunc.FileExists(strDestFile, False, gc_strAPPNAME) = True Then
                        Call g_xobjFunc.GetFileAndPath(m_strPreviewFilename, strPath, strFile)
                        Call g_xobjFunc.GetFileAndPath(strDestFile, strPath, strNewFile)
                        Call RefreshListsAfterConversion(strFile, strNewFile)
                    End If
                End If
            Else
                MsgBox m_strBMP2PNGPath & " not found!", vbExclamation, gc_strAPPNAME
                cmdBMP2PNG.Enabled = False
            End If
        Else
            MsgBox m_strPreviewFilename & " not found!", vbExclamation, gc_strAPPNAME
            cmdBMP2PNG.Enabled = False
        End If
    End If

    Screen.MousePointer = vbDefault

End Sub
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClear_Click()

    lblLink.Caption = g_strNoneSelected
    lblLink.FontBold = True
    cmdOK.Enabled = True

End Sub



Private Sub cmdHelp_Click()

    Call g_xobjFunc.ShellExecuteFile("http://www.bu22.com/wiki/picture_selector", "", False)

End Sub
Private Sub cmdOK_Click()

    'confirmation
    If chkAll.Value = vbChecked Then
        If MsgBox(g_xobjFunc.GetINIMessage("Batch Update", "GamesAreYouSure", "Are you sure you want to update ALL GAMES in the current view with the chosen data?"), vbYesNoCancel Or vbQuestion, gc_strAPPNAME) <> vbYes Then
            Exit Sub
        End If
    End If

    'return the data
    If lblLink.Caption = g_strNoneSelected Or lblLink.FontBold = True Then
        g_strScrnShotSelectorFilename = ""
    Else
        g_strScrnShotSelectorFilename = lblLink.Caption
    End If
    g_blnScrnShotSelectorOKPressed = True
    g_blnApplyScrnShotFileToAll = (chkAll.Value = vbChecked)

    Unload Me

End Sub




Private Sub cmdRefreshFind_Click()

    DirFind_Change

End Sub

Private Sub cmdRefreshPics_Click()

    DirPics_Change

End Sub


Private Sub DirFind_Change()

    m_strLastFindPath = DirFind.Path
    Call RefreshListView(lvFind, DirFind.Path)

End Sub

Private Sub DirFind_Click()

    DirFind.Path = DirFind.List(DirFind.ListIndex)

End Sub


Private Sub DirPics_Change()

    Dim lngIdx As Long

    'inside a screenshot path?
    With lstPaths
        For lngIdx = 0 To (.ListCount - 1)
            If UCase$(.List(lngIdx)) = UCase$(Left$(DirPics.Path, Len(.List(lngIdx)))) Then
                m_blnByCode = True
                .ListIndex = lngIdx
                m_blnByCode = False
                m_strLastPath = DirPics.Path
                Call RefreshListView(lvPics, DirPics.Path)
                Exit Sub
            End If
        Next lngIdx
    End With

    'not allowed - outside screenshot paths
    DirPics.Path = m_strLastPath

End Sub


Private Sub DirPics_Click()

    DirPics.Path = DirPics.List(DirPics.ListIndex)

End Sub


Private Sub DriveFind_Change()
    
    On Error GoTo No_Drive
    
    If m_blnByCode = False Then
        DirFind.Path = Left$(DriveFind.Drive, 2) & "\"
    End If

    Exit Sub
    
No_Drive:
    MsgBox Err.Description, vbExclamation, gc_strAPPNAME
    DriveFind.Drive = Left$(m_strLastFindPath, 3)
    
End Sub

Private Sub Form_Activate()

    If m_blnDoneActivate = False Then
        m_blnDoneActivate = True
        Call ShowTip(gc_lngTIPS_SCREENSHOTS2)
        If m_lngFocusIdx = 0 Then
            lvPics.SetFocus
        ElseIf m_lngFocusIdx = 1 Then
            lvFind.SetFocus
        Else
            DirFind.SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
    
    Dim lngPathIdx As Long
    Dim lngFileIdx As Long
    Dim lngColIdx As Long
    Dim blnFound As Boolean
    Dim strSSPathWithBackSlash As String
    Dim strPath As String
    Dim strFile As String
    Dim LItem As ListItem
    Dim lngCol As Long
    Dim lngIdx As Long
    Const DesignScaleHeight As Single = 5655    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 9015     'The required scalewidth for the form
    Dim sngRegWidth As Single                   'last set window width
    Dim sngRegHeight As Single                  'last set window height
    Dim sngRegTop As Single                     'last set window Y coord
    Dim sngRegLeft As Single                    'last set window X coord
    Dim blnRegMaximized As Boolean              'last set window maximized state
    Dim sngMinTwipsHeight As Single             'minimum height for form in twips
    Dim sngMinTwipsWidth As Single              'minimum width for form in twips
    Dim lngSortCol As Long
    Dim lngFindCol1Width As Long
    Dim lngFindCol2Width As Long
    Dim lngPicsCol1Width As Long
    Dim lngPicsCol2Width As Long

    m_blnDoneActivate = False
    m_strPreviewFilename = ""
    m_strDraggedFile = ""
    picFull.BackColor = g_lngPictureBackColor
    picPreview.BackColor = g_lngPictureBackColor
    If g_blnMamePicStyle = True Then
        m_blnARF = False
    Else
        m_blnARF = GetValidAutoRenameFilename(frmMain.lblName.Caption, m_strARF)
    End If

    'set the minimum size for the window
    sngMinTwipsWidth = DesignScaleWidth + (Width - ScaleWidth)
    sngMinTwipsHeight = DesignScaleHeight + (Height - ScaleHeight)
    lngPictureSelectorWindowMinX = sngMinTwipsWidth / Screen.TwipsPerPixelX
    lngPictureSelectorWindowMinY = sngMinTwipsHeight / Screen.TwipsPerPixelY

    'set the maximum size for the window
    lngPictureSelectorWindowMaxX = Screen.Width / Screen.TwipsPerPixelX
    lngPictureSelectorWindowMaxY = Screen.Height / Screen.TwipsPerPixelY

    'get current size/position/state variables
    With g_xobjRegistry
        'main window size/position
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorWidth", sngRegWidth, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorHeight", sngRegHeight, False, -100
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorTop", sngRegTop, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorLeft", sngRegLeft, False, -1
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorMaximized", blnRegMaximized, False, False
        'listview column widths
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorFindCol1Width", lngFindCol1Width, False, 3059
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorFindCol2Width", lngFindCol2Width, False, 675
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorPicsCol1Width", lngPicsCol1Width, False, 3059
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Window", "PictureSelectorPicsCol2Width", lngPicsCol2Width, False, 675
        lvFind.ColumnHeaders(1).Width = lngFindCol1Width
        lvFind.ColumnHeaders(2).Width = lngFindCol2Width
        lvPics.ColumnHeaders(1).Width = lngPicsCol1Width
        lvPics.ColumnHeaders(2).Width = lngPicsCol2Width
    End With

    'make sure not to big or small...
    If sngRegWidth < sngMinTwipsWidth Then
        sngRegWidth = sngMinTwipsWidth
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight < sngMinTwipsHeight Then
        sngRegHeight = sngMinTwipsHeight
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    '...or out of full screen view
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = (Screen.Height / 2) - (Height / 2)
        sngRegLeft = (Screen.Width / 2) - (Width / 2)
    End If
    'set the position and size
    If blnRegMaximized = False Then
        Top = sngRegTop
        Left = sngRegLeft
    Else
        WindowState = vbMaximized
    End If
    Width = sngRegWidth
    Height = sngRegHeight

    'subclass the window
    If g_blnInVBIDE = False Then
        Call PictureSelectorSubClass(Me.hwnd)
    End If

    'load the default text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText
    cmdHelp.Caption = g_strButtonHelpText
    cmdClear.Caption = g_strButtonClearText
    With g_xobjFunc
        Caption = .GetINIMessage("Picture Selector", "Caption", "Picture Selector")
        chkAll.Caption = .GetINIMessage("Picture Selector", "ApplyToAll", "A&pply to all games in current view")
        cmdARF.Caption = .GetINIMessage("Picture Selector", "AutoRenameButton", "&Auto Rename")
        fraFindPictures.Caption = .GetINIMessage("Picture Selector", "FindPictures", "Find Pictures")
        fraPicturePaths.Caption = .GetINIMessage("Picture Selector", "PicturePaths", "Picture Paths")
        fraPreview.Caption = .GetINIMessage("Picture Selector", "Preview", "Preview")
        lblLinkPrompt.Caption = .GetINIMessage("Picture Selector", "SelectedPicture", "Selected Picture") & ":"
        m_strIncompatibleScreenshot = .GetINIMessage("Picture Selector", "IncompatiblePicture", "Incompatible Picture")
        m_strEndsInX = .GetINIMessage("Picture Selector", "InvalidPictureEndsWithX", "Invalid Picture (Ends with '_x')")
        cmdRefreshFind.Caption = .GetINIMessage("Picture Selector", "RefreshFindButton", "&Refresh")
        cmdRefreshPics.Caption = .GetINIMessage("Picture Selector", "RefreshPicsButton", "R&efresh")
    End With

    'find the bmp2gif path
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "bmp2gif.exe", False, "") = True Then
        m_strBMP2GIFPath = g_xobjFunc.Long2Short(g_strAppPathWithBackslash & "bmp2gif.exe")
    Else
        m_strBMP2GIFPath = ""
    End If
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "bmp2png.exe", False, "") = True Then
        m_strBMP2PNGPath = g_xobjFunc.Long2Short(g_strAppPathWithBackslash & "bmp2png.exe")
    Else
        m_strBMP2PNGPath = ""
    End If

    'set the sorting columns for the listviews
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "FindPictureSortCol", lngSortCol, False, 0
    If lngSortCol = 0 Then
        lvFind.SortKey = 0
    Else
        lvFind.SortKey = 1
    End If
    lvFind.Sorted = True
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "PicturesPathSortCol", lngSortCol, False, 0
    If lngSortCol = 0 Then
        lvPics.SortKey = 0
    Else
        lvPics.SortKey = 1
    End If
    lvPics.Sorted = True

    'add picture paths to lstPaths
    With lstPaths
        For lngPathIdx = 0 To (g_lngTotalScrnShotPaths - 1)
            If ga_udtRegPathScrnShot(lngPathIdx).ePathType <> e_PathLocation_Invalid Then
                .AddItem ga_udtRegPathScrnShot(lngPathIdx).strRealPath
                .ItemData(.NewIndex) = lngPathIdx
            End If
        Next lngPathIdx
    End With

    'set the initial find directory
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPictureFindPath", strPath, True, ""
    If strPath <> "" Then
        If g_xobjFunc.DirExists(strPath) = False Then
            strPath = ""
        End If
    End If
    If strPath = "" Then
        With g_xobjGEMUS
            For lngIdx = 0 To (.TotalEmulators - 1)
                If g_xobjFunc.FileExists(.EmulatorPathExe(lngIdx), False, gc_strAPPNAME) = True Then
                    Call g_xobjFunc.GetFileAndPath(.EmulatorPathExe(lngIdx), strPath, strFile)
                    If g_xobjFunc.DirExists(strPath) = True Then
                        Exit For
                    Else
                        strPath = ""
                    End If
                End If
            Next lngIdx
            If strPath = "" Then
                strPath = App.Path
            End If
        End With
    End If
    m_blnByCode = True
    DriveFind.Drive = Left$(strPath, 3)
    m_blnByCode = False
    If UCase$(DirFind.Path) <> UCase$(strPath) Then
        DirFind.Path = strPath
    Else
        DirFind.Path = strPath
        Call RefreshListView(lvFind, strPath)
    End If

    'set the initial pictures directory
    m_blnByCode = True
    strPath = UCase$(g_strScrnShotSelectorOpeningPath)
    For lngIdx = 0 To (lstPaths.ListCount - 1)
        If UCase$(lstPaths.List(lngIdx)) = UCase$(Left$(g_strScrnShotSelectorOpeningPath, Len(lstPaths.List(lngIdx)))) Then
            lstPaths.ListIndex = lngIdx
        End If
    Next lngIdx
    m_blnByCode = False
    
    'select picture if exists
    If g_strScrnShotSelectorFilename <> "" Then
        'picture exists
        Call g_xobjFunc.GetFileAndPath(g_strScrnShotSelectorFilename, strPath, strFile)
        If UCase$(DirPics.Path) <> UCase$(strPath) Then
            DirPics.Path = strPath
        Else
            DirPics.Path = strPath
            Call RefreshListView(lvPics, strPath)
        End If
        If SelectFileInListView(lvPics, strFile) = False Then
            'select no file
            lblLink.Caption = g_strScrnShotSelectorDBFilename
            Set lvPics.SelectedItem = Nothing
            If lvPics.ListItems.Count > 0 Then
                lvPics.ListItems(1).Selected = False
            End If
        Else
            lvPics_ItemClick lvPics.SelectedItem
        End If
    Else
        'no picture file to select
        g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPicturePath", strPath, True, ""
        strSSPathWithBackSlash = g_strScrnShotSelectorOpeningPath
        If Right$(strSSPathWithBackSlash, 1) <> "\" Then strSSPathWithBackSlash = strSSPathWithBackSlash & "\"
        If Left$(UCase$(strPath), Len(strSSPathWithBackSlash)) <> UCase$(strSSPathWithBackSlash) Then
            strPath = g_strScrnShotSelectorOpeningPath
        End If
        DirPics.Path = strPath
        If UCase$(DirPics.Path) = UCase$(strPath) Then
            Call RefreshListView(lvPics, strPath)
        End If
        'select NO file
        If g_strScrnShotSelectorDBFilename = "" Then
            lblLink.FontBold = True
            lblLink.Caption = g_strNoneSelected
        Else
            lblLink.Caption = g_strScrnShotSelectorDBFilename
        End If
        Set lvPics.SelectedItem = Nothing
        If lvPics.ListItems.Count > 0 Then
            lvPics.ListItems(1).Selected = False
        End If
    End If

    'if nothing selected in the pics listview, select the first item in the find listview
    If lvPics.SelectedItem Is Nothing Then
        If lvFind.ListItems.Count > 0 Then
            Set lvFind.SelectedItem = lvFind.ListItems(1)
            lvFind.SelectedItem.Selected = True
            lvFind_ItemClick lvFind.SelectedItem
        End If
    End If

    'set the focus index for Form_Activate()
    If Not (lvPics.SelectedItem Is Nothing) Then
        m_lngFocusIdx = 0
    ElseIf Not (lvFind.SelectedItem Is Nothing) Then
        m_lngFocusIdx = 1
    Else
        m_lngFocusIdx = 2
    End If

    'all done
    g_blnScrnShotSelectorOKPressed = False
    cmdOK.Enabled = True
    Screen.MousePointer = vbDefault

End Sub









Private Sub Form_Resize()

    Dim lngFullXWidth As Long

    If WindowState <> vbMinimized Then
    
        'horizontal
        
        lngFullXWidth = ScaleWidth - 360
        fraFindPictures.Width = Int(lngFullXWidth / 2)
        fraPicturePaths.Left = 120 + fraFindPictures.Width + 120
        fraPicturePaths.Width = ScaleWidth - fraPicturePaths.Left - 120
        cmdCancel.Left = ScaleWidth - 120 - cmdCancel.Width
        cmdOK.Left = cmdCancel.Left - 120 - cmdOK.Width
        cmdHelp.Left = fraPicturePaths.Left
        fraPreview.Width = fraPicturePaths.Width
        
        DirFind.Width = fraFindPictures.Width - 240
        DriveFind.Width = DirFind.Width - cmdRefreshFind.Width - 120
        cmdRefreshFind.Left = DirFind.Left + DirFind.Width - cmdRefreshFind.Width
        lvFind.Width = DirFind.Width
        lblPreview.Width = fraPreview.Width - 240
        
        DirPics.Width = fraPicturePaths.Width - 240
        lstPaths.Width = DirPics.Width - cmdRefreshPics.Width - 120
        cmdRefreshPics.Left = DirPics.Left + DirPics.Width - cmdRefreshPics.Width
        lvPics.Width = DirPics.Width
        lblLink.Width = DirPics.Width
        cmdClear.Left = fraPicturePaths.Width - 120 - cmdClear.Width
    
        'vertical
        
        cmdCancel.Top = ScaleHeight - 120 - cmdCancel.Height
        cmdOK.Top = cmdCancel.Top
        cmdHelp.Top = cmdCancel.Top
        fraPreview.Top = ScaleHeight - 120 - fraPreview.Height
        fraFindPictures.Height = fraPreview.Top - 120 - fraFindPictures.Top
        fraPicturePaths.Height = cmdHelp.Top - 240
        
        lvFind.Height = fraPreview.Top - 360 - lvFind.Top
        
        chkAll.Top = fraPicturePaths.Height - fraPicturePaths.Top - chkAll.Height
        lblLink.Top = chkAll.Top - 120 - lblLink.Height
        lblLinkPrompt.Top = lblLink.Top - lblLinkPrompt.Height
        cmdClear.Top = lblLinkPrompt.Top
        
        lvPics.Height = cmdClear.Top - 240 - lvPics.Top
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'un subclass the window
    If g_blnInVBIDE = False Then
        If g_blnPictureSelectorSubclassed = True Then
            Call PictureSelectorUnSubClass(Me.hwnd)
        End If
    End If

    'save reg settings
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorWidth", Width, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorHeight", Height, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorLeft", Left, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorMaximized", False, False
        ElseIf WindowState = vbMaximized Then
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorMaximized", True, False
        End If
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "FindPictureSortCol", lvFind.SortKey, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Options", "PicturesPathSortCol", lvPics.SortKey, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPictureFindPath", DirFind.Path, True
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\BrowsePaths", "LastPicturePath", DirPics.Path, True
        'listview column widths
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorFindCol1Width", lvFind.ColumnHeaders(1).Width, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorFindCol2Width", lvFind.ColumnHeaders(2).Width, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorPicsCol1Width", lvPics.ColumnHeaders(1).Width, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\Window", "PictureSelectorPicsCol2Width", lvPics.ColumnHeaders(2).Width, False
    End With

End Sub


Private Sub lstPaths_Click()

    If m_blnByCode = False Then
        With lstPaths
            If .ListIndex > -1 Then
                DirPics.Path = .List(.ListIndex)
            End If
        End With
    End If

End Sub


Private Sub lvFind_AfterLabelEdit(Cancel As Integer, NewString As String)

    Dim strNewFilename As String
    Dim LItem As ListItem
    Dim strUpper As String

    'get the full new filename
    strNewFilename = DirFind.Path
    If Right$(strNewFilename, 1) <> "\" Then
        strNewFilename = strNewFilename & "\"
    End If
    strNewFilename = strNewFilename & NewString

    'try to rename it
    If RenameListViewFile(m_strOldFilename, strNewFilename) = True Then
        MousePointer = vbHourglass
        DoEvents
        'renamed - do we need to rename the file in the other list?
        If UCase$(DirFind.Path) = UCase$(DirPics.Path) Then
            strUpper = UCase$(lvFind.SelectedItem.Text)
            For Each LItem In lvPics.ListItems
                If UCase$(LItem.Text) = strUpper Then
                    LItem.Text = NewString
                    Exit For
                End If
            Next LItem
        End If
        lvFind.SelectedItem.Text = NewString
        Call ClickLVFind(lvFind.SelectedItem)
        MousePointer = vbDefault
    Else
        'error or no change
        Cancel = True
    End If
    cmdCancel.Cancel = True

End Sub

Private Sub lvFind_BeforeLabelEdit(Cancel As Integer)

    m_strOldFilename = DirFind.Path
    If Right$(m_strOldFilename, 1) <> "\" Then
        m_strOldFilename = m_strOldFilename & "\"
    End If
    m_strOldFilename = m_strOldFilename & lvFind.SelectedItem.Text
    cmdCancel.Cancel = False

End Sub

Private Sub lvFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvFind
        If ColumnHeader.Index = 1 Then
            If .SortKey = 1 Then
                Screen.MousePointer = vbHourglass
                DoEvents
                .Visible = False
                .SortKey = 0
                .Visible = True
                Screen.MousePointer = vbDefault
            End If
        Else
            If .SortKey = 0 Then
                Screen.MousePointer = vbHourglass
                DoEvents
                .Visible = False
                .SortKey = 1
                .Visible = True
                Screen.MousePointer = vbDefault
            End If
        End If
    End With

End Sub


Private Sub lvFind_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Call ClickLVFind(Item)

End Sub


Private Sub lvFind_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim strJustFile As String
    Dim strPath As String
    Dim strPathFile As String
    Dim LItem As ListItem
    Dim blnFound As Boolean

    If (KeyCode = vbKeyDelete) And (Shift = 0) Then
        If Not (lvFind.SelectedItem Is Nothing) Then
            strPath = DirFind.Path
            strJustFile = lvFind.SelectedItem.Text
            strPathFile = strPath
            If Right$(strPathFile, 1) <> "\" Then strPathFile = strPathFile & "\"
            strPathFile = strPathFile & strJustFile
            If g_xobjFunc.FileExists(strPathFile, False, gc_strAPPNAME) = True Then
                If MsgBox("'" & strPathFile & "'" & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Picture Selector", "ConfirmDelete", "Are you sure you want to delete this file?"), vbYesNoCancel Or vbQuestion Or vbDefaultButton2, gc_strAPPNAME) = vbYes Then
                    'proceed to delete
                    If DeleteListViewFile(strPathFile) = True Then
                        'remove from listview
                        lvFind.ListItems.Remove lvFind.SelectedItem.Index
                        Set lvFind.SelectedItem = Nothing
                        cmdBMP2PNG.Enabled = False
                        cmdBMP2GIF.Enabled = False
                        picPreview.Visible = False
                        lblPreview.Caption = ""
                        'remove from other listview?
                        If UCase$(DirFind.Path) = UCase$(DirPics.Path) Then
                            strJustFile = UCase$(strJustFile)
                            For Each LItem In lvPics.ListItems
                                If UCase$(LItem.Text) = strJustFile Then
                                    lvPics.ListItems.Remove LItem.Index
                                    Exit For
                                End If
                            Next LItem
                        End If
                    End If
                End If
            End If
        End If

    ElseIf (KeyCode = 116) And (Shift = 0) Then
    
        'REFRESH LIST (F5)

        blnFound = False
        If Not (lvFind.SelectedItem Is Nothing) Then
            strJustFile = UCase$(lvFind.SelectedItem.Text)
        Else
            strJustFile = ""
        End If
        Call RefreshListView(lvFind, DirFind.Path)
        If strJustFile <> "" Then
            For Each LItem In lvFind.ListItems
                If UCase$(LItem.Text) = strJustFile Then
                    blnFound = True
                    Set lvFind.SelectedItem = LItem
                    LItem.Selected = True
                    LItem.EnsureVisible
                    Call ClickLVFind(lvFind.SelectedItem)
                    lvFind.SetFocus
                    Exit For
                End If
            Next LItem
        End If
        If blnFound = False Then
            Set lvFind.SelectedItem = Nothing
            If lvFind.ListItems.Count > 0 Then
                lvFind.ListItems(1).Selected = False
            End If
            cmdClear.Value = True
            cmdBMP2PNG.Enabled = False
            cmdBMP2GIF.Enabled = False
            picPreview.Visible = False
            lblPreview.Caption = ""
            lvFind.SetFocus
        End If
    End If

End Sub

Private Sub lvFind_OLECompleteDrag(Effect As Long)

    Dim strPath As String
    Dim strPathFile As String
    Dim strJustFile As String

    If Effect = vbDropEffectMove Then
        'delete source file
        If Not (lvFind.SelectedItem Is Nothing) Then
            strPath = DirFind.Path
            strJustFile = lvFind.SelectedItem.Text
            strPathFile = strPath
            If Right$(strPathFile, 1) <> "\" Then strPathFile = strPathFile & "\"
            strPathFile = strPathFile & strJustFile
            If g_xobjFunc.FileExists(strPathFile, False, gc_strAPPNAME) = True Then
                'proceed to delete
                If DeleteListViewFile(strPathFile) = True Then
                    'remove from listview
                    lvFind.ListItems.Remove lvFind.SelectedItem.Index
                    Set lvFind.SelectedItem = Nothing
                End If
            End If
        End If
    End If

End Sub

Private Sub lvFind_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strFile As String
    Dim strPath As String
    Dim strJustFile As String
    Dim strExt As String
    Dim LI As ListItem
    Dim LItem As ListItem
    Dim strUpper As String
    Dim blnFound As Boolean

    On Error GoTo Err_Handler

    MousePointer = vbHourglass
    DoEvents

    If Data.GetFormat(vbCFFiles) = True Then
        strFile = Data.Files(1)
        If g_xobjFunc.FileExists(strFile, False, gc_strAPPNAME) = True Then
            Call g_xobjFunc.GetFileAndPath(strFile, strPath, strJustFile)
            If UCase$(strPath) <> UCase$(DirFind.Path) Then
                'is it a valid picture file?
                strExt = UCase$(Right$(strJustFile, 4))
                If strExt = ".GIF" Or strExt = ".BMP" Or strExt = ".JPG" Or strExt = ".PNG" Then
                    'copy the file
                    If CopyReplaceFile(strFile, DirFind.Path) = True Then
                        'add the file to the listview if not there already
                        strUpper = UCase$(strJustFile)
                        blnFound = False
                        With lvFind
                            For Each LI In .ListItems
                                If UCase$(LI.Text) = strUpper Then
                                    blnFound = True
                                    Set LItem = LI
                                    Exit For
                                End If
                            Next LI
                            If blnFound = False Then
                                Set LItem = .ListItems.Add(, , strJustFile)
                                LItem.SubItems(1) = Right$(strExt, 3)
                            End If
                            Set .SelectedItem = LItem
                            LItem.Selected = True
                            Set LItem = Nothing
                            Call ClickLVFind(.SelectedItem)
                            .Sorted = False
                            .Sorted = True
                            .SelectedItem.EnsureVisible
                            .SetFocus
                        End With
                        'notify the drag-source
                        If Shift = 0 Then
                            'move file
                            Effect = vbDropEffectMove
                        Else
                            'copy file
                            Effect = vbDropEffectCopy
                        End If
                    Else
                        'copy failed or aborted
                        Effect = vbDropEffectNone
                    End If
                Else
                    'invalid file type
                    MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "NotAValidPictureFile", "Not a valid picture file.") & "  (BMP, GIF, JPG, PNG)", vbInformation, gc_strAPPNAME
                    Effect = vbDropEffectNone
                End If
            Else
                'file already in this folder
                Effect = vbDropEffectNone
            End If
            
        Else
            'file doesn't exist
            Effect = vbDropEffectNone
        End If
    Else
        'not files
        Effect = vbDropEffectNone
    End If

    MousePointer = vbDefault

    Exit Sub
    
Err_Handler:
    MsgBox "Error Receiving dropped file.  " & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME
    Effect = vbDropEffectNone
    MousePointer = vbDefault

End Sub

Private Sub lvFind_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    If Data.GetFormat(vbCFFiles) = True Then
        If (Shift = 0) Then
            Effect = vbDropEffectMove
        ElseIf (Shift = 1) Or (Shift = 2) Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If

End Sub

Private Sub lvFind_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

    m_strDraggedFile = DirFind.Path
    If Right$(m_strDraggedFile, 1) <> "\" Then m_strDraggedFile = m_strDraggedFile & "\"
    m_strDraggedFile = m_strDraggedFile & lvFind.SelectedItem.Text
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    Data.Clear
    Data.Files.Add m_strDraggedFile
    Data.SetData , vbCFFiles

End Sub

Private Sub lvPics_AfterLabelEdit(Cancel As Integer, NewString As String)

    Dim strNewFilename As String
    Dim LItem As ListItem
    Dim strUpper As String

    'get the full new filename
    strNewFilename = DirPics.Path
    If Right$(strNewFilename, 1) <> "\" Then
        strNewFilename = strNewFilename & "\"
    End If
    strNewFilename = strNewFilename & NewString

    'try to rename it
    If RenameListViewFile(m_strOldFilename, strNewFilename) = True Then
        MousePointer = vbHourglass
        DoEvents
        'renamed - do we need to rename the file in the other list?
        If UCase$(DirPics.Path) = UCase$(DirFind.Path) Then
            strUpper = UCase$(lvPics.SelectedItem.Text)
            For Each LItem In lvFind.ListItems
                If UCase$(LItem.Text) = strUpper Then
                    LItem.Text = NewString
                    Exit For
                End If
            Next LItem
        End If
        lvPics.SelectedItem.Text = NewString
        Call ClickLVPics(lvPics.SelectedItem)
        MousePointer = vbDefault
    Else
        'error or no change
        Cancel = True
    End If
    cmdCancel.Cancel = True

End Sub
Private Sub lvPics_BeforeLabelEdit(Cancel As Integer)

    m_strOldFilename = DirPics.Path
    If Right$(m_strOldFilename, 1) <> "\" Then
        m_strOldFilename = m_strOldFilename & "\"
    End If
    m_strOldFilename = m_strOldFilename & lvPics.SelectedItem.Text
    cmdCancel.Cancel = False

End Sub


Private Sub lvPics_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Debug.Print ColumnHeader.Width
    With lvPics
        If ColumnHeader.Index = 1 Then
            If .SortKey = 1 Then
                Screen.MousePointer = vbHourglass
                DoEvents
                .Visible = False
                .SortKey = 0
                .Visible = True
                Screen.MousePointer = vbDefault
            End If
        Else
            If .SortKey = 0 Then
                Screen.MousePointer = vbHourglass
                DoEvents
                .Visible = False
                .SortKey = 1
                .Visible = True
                Screen.MousePointer = vbDefault
            End If
        End If
    End With

End Sub


Private Sub lvPics_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Call ClickLVPics(Item)

End Sub

Private Sub lvPics_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim strJustFile As String
    Dim strPath As String
    Dim strPathFile As String
    Dim LItem As ListItem
    Dim blnFound As Boolean

    If (KeyCode = vbKeyDelete) And (Shift = 0) Then
    
        'DELETE FILE (Delete Key)
    
        If Not (lvPics.SelectedItem Is Nothing) Then
            strPath = DirPics.Path
            strJustFile = lvPics.SelectedItem.Text
            strPathFile = strPath
            If Right$(strPathFile, 1) <> "\" Then strPathFile = strPathFile & "\"
            strPathFile = strPathFile & strJustFile
            If g_xobjFunc.FileExists(strPathFile, False, gc_strAPPNAME) = True Then
                If MsgBox("'" & strPathFile & "'" & vbCrLf & vbCrLf & g_xobjFunc.GetINIMessage("Picture Selector", "ConfirmDelete", "Are you sure you want to delete this file?"), vbYesNoCancel Or vbQuestion Or vbDefaultButton2, gc_strAPPNAME) = vbYes Then
                    'proceed to delete
                    If DeleteListViewFile(strPathFile) = True Then
                        'remove from listview
                        lvPics.ListItems.Remove lvPics.SelectedItem.Index
                        Set lvPics.SelectedItem = Nothing
                        cmdClear.Value = True
                        cmdBMP2PNG.Enabled = False
                        cmdBMP2GIF.Enabled = False
                        picPreview.Visible = False
                        lblPreview.Caption = ""
                        'remove from other listview?
                        If UCase$(DirPics.Path) = UCase$(DirFind.Path) Then
                            strJustFile = UCase$(strJustFile)
                            For Each LItem In lvFind.ListItems
                                If UCase$(LItem.Text) = strJustFile Then
                                    lvFind.ListItems.Remove LItem.Index
                                    Exit For
                                End If
                            Next LItem
                        End If
                    End If
                End If
            End If
        End If
    ElseIf (KeyCode = 116) And (Shift = 0) Then
    
        'REFRESH LIST (F5)

        blnFound = False
        If Not (lvPics.SelectedItem Is Nothing) Then
            strJustFile = UCase$(lvPics.SelectedItem.Text)
        Else
            strJustFile = ""
        End If
        Call RefreshListView(lvPics, DirPics.Path)
        If strJustFile <> "" Then
            For Each LItem In lvPics.ListItems
                If UCase$(LItem.Text) = strJustFile Then
                    blnFound = True
                    Set lvPics.SelectedItem = LItem
                    LItem.Selected = True
                    LItem.EnsureVisible
                    Call ClickLVPics(lvPics.SelectedItem)
                    lvPics.SetFocus
                    Exit For
                End If
            Next LItem
        End If
        If blnFound = False Then
            Set lvPics.SelectedItem = Nothing
            If lvPics.ListItems.Count > 0 Then
                lvPics.ListItems(1).Selected = False
            End If
            cmdClear.Value = True
            cmdBMP2PNG.Enabled = False
            cmdBMP2GIF.Enabled = False
            picPreview.Visible = False
            lblPreview.Caption = ""
            lvPics.SetFocus
        End If

    End If

End Sub


Private Sub lvPics_OLECompleteDrag(Effect As Long)

    Dim strPath As String
    Dim strPathFile As String
    Dim strJustFile As String

    If Effect = vbDropEffectMove Then
        'delete source file
        If Not (lvPics.SelectedItem Is Nothing) Then
            strPath = DirPics.Path
            strJustFile = lvPics.SelectedItem.Text
            strPathFile = strPath
            If Right$(strPathFile, 1) <> "\" Then strPathFile = strPathFile & "\"
            strPathFile = strPathFile & strJustFile
            If g_xobjFunc.FileExists(strPathFile, False, gc_strAPPNAME) = True Then
                'proceed to delete
                If DeleteListViewFile(strPathFile) = True Then
                    'remove from listview
                    lvPics.ListItems.Remove lvPics.SelectedItem.Index
                    Set lvPics.SelectedItem = Nothing
                    cmdClear.Value = True
                End If
            End If
        End If
    End If

End Sub

Private Sub lvPics_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strFile As String
    Dim strPath As String
    Dim strJustFile As String
    Dim strExt As String
    Dim LI As ListItem
    Dim LItem As ListItem
    Dim strUpper As String
    Dim blnFound As Boolean

    On Error GoTo Err_Handler

    MousePointer = vbHourglass
    DoEvents

    If Data.GetFormat(vbCFFiles) = True Then
        strFile = Data.Files(1)
        If g_xobjFunc.FileExists(strFile, False, gc_strAPPNAME) = True Then
            Call g_xobjFunc.GetFileAndPath(strFile, strPath, strJustFile)
            If UCase$(strPath) <> UCase$(DirPics.Path) Then
                'is it a valid picture file?
                strExt = UCase$(Right$(strJustFile, 4))
                If strExt = ".GIF" Or strExt = ".BMP" Or strExt = ".JPG" Or strExt = ".PNG" Then
                    'copy the file
                    If CopyReplaceFile(strFile, DirPics.Path) = True Then
                        'add the file to the listview if not there already
                        strUpper = UCase$(strJustFile)
                        blnFound = False
                        With lvPics
                            For Each LI In .ListItems
                                If UCase$(LI.Text) = strUpper Then
                                    blnFound = True
                                    Set LItem = LI
                                    Exit For
                                End If
                            Next LI
                            If blnFound = False Then
                                Set LItem = .ListItems.Add(, , strJustFile)
                                LItem.SubItems(1) = Right$(strExt, 3)
                            End If
                            Set .SelectedItem = LItem
                            LItem.Selected = True
                            Set LItem = Nothing
                            Call ClickLVPics(.SelectedItem)
                            .Sorted = False
                            .Sorted = True
                            .SelectedItem.EnsureVisible
                            .SetFocus
                        End With
                        'notify the drag-source
                        If Shift = 0 Then
                            'move file
                            Effect = vbDropEffectMove
                        Else
                            'copy file
                            Effect = vbDropEffectCopy
                        End If
                    Else
                        'copy failed or aborted
                        Effect = vbDropEffectNone
                    End If
                Else
                    'invalid file type
                    MsgBox g_xobjFunc.GetINIMessage("Picture Selector", "NotAValidPictureFile", "Not a valid picture file.") & "  (BMP, GIF, JPG, PNG)", vbInformation, gc_strAPPNAME
                    Effect = vbDropEffectNone
                End If
            Else
                'file already in this folder
                Effect = vbDropEffectNone
            End If
            
        Else
            'file doesn't exist
            Effect = vbDropEffectNone
        End If
    Else
        'not files
        Effect = vbDropEffectNone
    End If

    MousePointer = vbDefault

    Exit Sub
    
Err_Handler:
    MsgBox "Error Receiving dropped file.  " & Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", vbExclamation, gc_strAPPNAME
    Effect = vbDropEffectNone
    MousePointer = vbDefault

End Sub
Private Sub lvPics_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    If Data.GetFormat(vbCFFiles) = True Then
        If (Shift = 0) Then
            Effect = vbDropEffectMove
        ElseIf (Shift = 1) Or (Shift = 2) Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If

End Sub


Private Sub lvPics_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

    m_strDraggedFile = DirPics.Path
    If Right$(m_strDraggedFile, 1) <> "\" Then m_strDraggedFile = m_strDraggedFile & "\"
    m_strDraggedFile = m_strDraggedFile & lvPics.SelectedItem.Text
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    Data.Clear
    Data.Files.Add m_strDraggedFile
    Data.SetData , vbCFFiles

End Sub



