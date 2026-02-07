VERSION 5.00
Begin VB.Form frmEggs 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Happy Easter! :)"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEaster 
      Caption         =   "START"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.PictureBox picAction 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   5160
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picWork 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   720
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   7
      Left            =   1680
      Picture         =   "frmEggs.frx":0000
      Top             =   5160
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   6
      Left            =   720
      Picture         =   "frmEggs.frx":0636
      Top             =   5160
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   5
      Left            =   2640
      Picture         =   "frmEggs.frx":0BDE
      Top             =   4200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   4
      Left            =   1680
      Picture         =   "frmEggs.frx":10BF
      Top             =   4200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   3
      Left            =   720
      Picture         =   "frmEggs.frx":16A4
      Top             =   4200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   2640
      Picture         =   "frmEggs.frx":1CC2
      Top             =   3240
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   1
      Left            =   1680
      Picture         =   "frmEggs.frx":2172
      Top             =   3240
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSliderTile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Index           =   0
      Left            =   720
      Picture         =   "frmEggs.frx":26EF
      Top             =   3240
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmEggs"
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

Private m_lngEggIdx As Long
Private m_blnActivated As Boolean
Private m_blnInGame As Boolean
Private m_lngStartTime&

Private m_blnNoUnload As Boolean
Private Const cKEY_LEFT As Integer = 37
Private Const cKEY_UP As Integer = 38
Private Const cKEY_RIGHT As Integer = 39
Private Const cKEY_DOWN As Integer = 40


'TIMING
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'SLIDER
Private Type SlideCellType
    lngRow As Long
    lngCol As Long
End Type
Private ma_lngSlideTilePos(7) As SlideCellType
Private m_udtSlideHole As SlideCellType
Private ma_lngSlideGrid(2, 2) As Long
Private m_lngSlideMoves As Long
'A* solver


Private Function SLIDE_IsSolved() As Boolean

    SLIDE_IsSolved = _
        ma_lngSlideGrid(0, 0) = 0 And _
        ma_lngSlideGrid(0, 1) = 1 And _
        ma_lngSlideGrid(0, 2) = 2 And _
        ma_lngSlideGrid(1, 0) = 3 And _
        ma_lngSlideGrid(1, 1) = 4 And _
        ma_lngSlideGrid(1, 2) = 5 And _
        ma_lngSlideGrid(2, 0) = 6 And _
        ma_lngSlideGrid(2, 1) = 7 And _
        ma_lngSlideGrid(2, 2) = -1

End Function

Private Sub SLIDE_Reset()

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTile As Long

    'OOO
    'OOO
    'OO

    'Initialize the 3x3 tile grid and the tile array
    lngTile = 0
    For lngRow = 0 To 2
        For lngCol = 0 To 2
            If Not (lngRow = 2 And lngCol = 2) Then
                ma_lngSlideGrid(lngRow, lngCol) = lngTile
                ma_lngSlideTilePos(lngTile).lngRow = lngRow
                ma_lngSlideTilePos(lngTile).lngCol = lngCol
                lngTile = lngTile + 1
            End If
        Next lngCol
    Next lngRow
    ma_lngSlideGrid(2, 2) = -1
    m_udtSlideHole.lngRow = 2
    m_udtSlideHole.lngCol = 2

End Sub

Private Sub SLIDE_RandomizeGrid()

    Dim lngEndTime&
    Dim lngRndTileIdx As Long
    Dim lngEndPause&
    Static lngLastTileIdx As Long

    'Randomize the tiles
    lngEndTime& = timeGetTime() + 4000
    m_blnNoUnload = True
    Do While timeGetTime() < lngEndTime&
        'get a random tile to move
        Randomize
        lngRndTileIdx = SLIDE_GetRandomTileFromHole()
        If lngRndTileIdx <> lngLastTileIdx Then
            'Move it
            Call SLIDE_MoveTile(lngRndTileIdx)
            'update the display
            Call SLIDE_UpdateDisplay
            DoEvents
            lngEndPause& = timeGetTime() + 30
            Do While timeGetTime() < lngEndPause&
            Loop
            lngLastTileIdx = lngRndTileIdx
        End If
    Loop
    m_blnNoUnload = False

End Sub

Private Sub SLIDE_UpdateDisplay()

    'Print the tiles to the work area then flip to the action area

    Dim lngTileIdx As Long

    With picWork
        .Cls
        .Picture = LoadPicture()
        For lngTileIdx = 0 To 7
            .PaintPicture imgSliderTile(lngTileIdx).Picture, ma_lngSlideTilePos(lngTileIdx).lngCol * 60, ma_lngSlideTilePos(lngTileIdx).lngRow * 60, 60, 60, 0, 0, 60, 60, vbSrcCopy
            picAction.Picture = picWork.Image
        Next lngTileIdx
        picAction.Picture = picWork.Image
    End With

End Sub

Private Function SLIDE_GetRandomTileFromHole() As Long

    'Get all the tiles around the hole

    Dim lngTotalTiles As Long
    Dim a_lngTiles() As Long
    Dim lngRndIdx As Long

    lngTotalTiles = 0
    With m_udtSlideHole
        If .lngRow > 0 Then
            If ma_lngSlideGrid(.lngRow - 1, .lngCol) > -1 Then
                'add tile above the hole
                ReDim Preserve a_lngTiles(lngTotalTiles)
                a_lngTiles(lngTotalTiles) = ma_lngSlideGrid(.lngRow - 1, .lngCol)
                lngTotalTiles = lngTotalTiles + 1
            End If
        End If
        If .lngRow < 2 Then
            If ma_lngSlideGrid(.lngRow + 1, .lngCol) > -1 Then
                'add tile below the hole
                ReDim Preserve a_lngTiles(lngTotalTiles)
                a_lngTiles(lngTotalTiles) = ma_lngSlideGrid(.lngRow + 1, .lngCol)
                lngTotalTiles = lngTotalTiles + 1
            End If
        End If
        If .lngCol > 0 Then
            If ma_lngSlideGrid(.lngRow, .lngCol - 1) > -1 Then
                'add tile to the left of the hole
                ReDim Preserve a_lngTiles(lngTotalTiles)
                a_lngTiles(lngTotalTiles) = ma_lngSlideGrid(.lngRow, .lngCol - 1)
                lngTotalTiles = lngTotalTiles + 1
            End If
        End If
        If .lngCol < 2 Then
            If ma_lngSlideGrid(.lngRow, .lngCol + 1) > -1 Then
                'add tile to the right of the hole
                ReDim Preserve a_lngTiles(lngTotalTiles)
                a_lngTiles(lngTotalTiles) = ma_lngSlideGrid(.lngRow, .lngCol + 1)
                lngTotalTiles = lngTotalTiles + 1
            End If
        End If
    End With
    lngRndIdx = Int((lngTotalTiles - 1 + 1) * Rnd + 1)
    SLIDE_GetRandomTileFromHole = a_lngTiles(lngRndIdx - 1)

End Function

Private Sub SLIDE_MoveTile(ByVal lngTileIdx_in As Long)

    'move the passed-in tile into the hole!

    Dim lngTileRow As Long
    Dim lngTileCol As Long

    With ma_lngSlideTilePos(lngTileIdx_in)
        lngTileRow = .lngRow
        lngTileCol = .lngCol
        .lngRow = m_udtSlideHole.lngRow
        .lngCol = m_udtSlideHole.lngCol
        m_udtSlideHole.lngRow = lngTileRow
        m_udtSlideHole.lngCol = lngTileCol
        ma_lngSlideGrid(m_udtSlideHole.lngRow, m_udtSlideHole.lngCol) = -1
        ma_lngSlideGrid(.lngRow, .lngCol) = lngTileIdx_in
    End With

End Sub

Private Sub DoSlider()

    'init the display
    Call InitDisplay(180, 180, 0, 0, &H0&, &H0&)
    
    'Init the slider
    Call SLIDE_Reset
    Call SLIDE_UpdateDisplay

End Sub

Private Sub InitDisplay(ByVal lngHeight_in As Long, ByVal lngWidth_in As Long, ByVal lngLeft_in As Long, ByVal lngTop_in As Long, ByVal lngBackColor_in As Long, ByVal lngBorderColor_in As Long)

    'init the action and work windows
    With picAction
        .Height = lngHeight_in
        .Width = lngWidth_in
        .Top = lngTop_in
        .Left = lngLeft_in
        .BackColor = lngBackColor_in
        .Visible = True
    End With
    With picWork
        .Height = lngHeight_in
        .Width = lngWidth_in
        .BackColor = lngBackColor_in
        .Cls
    End With
    'set the border around the action window


End Sub


Private Sub cmdEaster_Click()

    If cmdEaster.Caption = "START" Then
        cmdEaster.Enabled = False
        DoEvents
        If m_lngEggIdx = 1 Then
            Call SLIDE_RandomizeGrid
            m_lngStartTime& = timeGetTime()
            m_lngSlideMoves = 0
            m_blnInGame = True
            picAction.Enabled = True
        End If
        cmdEaster.Caption = "EXIT"
        cmdEaster.Enabled = True
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Activate()

    If Not m_blnActivated Then
        m_blnActivated = True
        If m_lngEggIdx = 1 Then
            'start slider
            Call DoSlider
        ElseIf m_lngEggIdx = 2 Then
            'start solitaire
        Else
            'invalid egg - close window
            Unload Me
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

    If m_blnInGame = True Then
        If KeyCode = cKEY_UP Then
            If m_udtSlideHole.lngRow < 2 Then
                Call SLIDE_PicActionClick(m_udtSlideHole.lngCol * 64, (m_udtSlideHole.lngRow + 1) * 60)
            End If
        ElseIf KeyCode = cKEY_DOWN Then
            If m_udtSlideHole.lngRow > 0 Then
                Call SLIDE_PicActionClick(m_udtSlideHole.lngCol * 64, (m_udtSlideHole.lngRow - 1) * 60)
            End If
        ElseIf KeyCode = cKEY_LEFT Then
            If m_udtSlideHole.lngCol < 2 Then
                Call SLIDE_PicActionClick((m_udtSlideHole.lngCol + 1) * 64, m_udtSlideHole.lngRow * 60)
            End If
        ElseIf KeyCode = cKEY_RIGHT Then
            If m_udtSlideHole.lngCol > 0 Then
                Call SLIDE_PicActionClick((m_udtSlideHole.lngCol - 1) * 64, m_udtSlideHole.lngRow * 60)
            End If
        End If
    End If

    

End Sub


Private Sub Form_Load()

    Width = (Width - (ScaleWidth * Screen.TwipsPerPixelX)) + (Screen.TwipsPerPixelX * 180)
    Height = (Height - (ScaleHeight * Screen.TwipsPerPixelY)) + (Screen.TwipsPerPixelY * 180) + (Screen.TwipsPerPixelY * cmdEaster.Height)
    cmdEaster.Top = 180
    cmdEaster.Width = ScaleWidth
    cmdEaster.Left = 0

    m_lngEggIdx = 1
    m_blnActivated = False
    m_blnInGame = False
    m_blnNoUnload = False

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If m_blnNoUnload = True Then
        Cancel = True
    End If

End Sub

Private Sub picAction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_lngEggIdx = 1 Then
        Call SLIDE_PicActionClick(X, Y)
    End If

End Sub
Private Sub SLIDE_PicActionClick(ByVal X As Single, ByVal Y As Single)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnDoIt As Boolean

    If m_blnInGame = True Then
        lngRow = Int(Y / 60)
        lngCol = Int(X / 60)
        If (lngRow > -1) And (lngRow < 3) And (lngCol > -1) And (lngCol < 3) Then
            If ma_lngSlideGrid(lngRow, lngCol) > -1 Then
                blnDoIt = False
                If lngRow > 0 Then
                    If ma_lngSlideGrid(lngRow - 1, lngCol) = -1 Then
                        blnDoIt = True
                    End If
                End If
                If blnDoIt = False Then
                    If lngRow < 2 Then
                        If ma_lngSlideGrid(lngRow + 1, lngCol) = -1 Then
                            blnDoIt = True
                        End If
                    End If
                End If
                If blnDoIt = False Then
                    If lngCol > 0 Then
                        If ma_lngSlideGrid(lngRow, lngCol - 1) = -1 Then
                            blnDoIt = True
                        End If
                    End If
                End If
                If blnDoIt = False Then
                    If lngCol < 2 Then
                        If ma_lngSlideGrid(lngRow, lngCol + 1) = -1 Then
                            blnDoIt = True
                        End If
                    End If
                End If
                If blnDoIt = True Then
                    Call SLIDE_MoveTile(ma_lngSlideGrid(lngRow, lngCol))
                    Call SLIDE_UpdateDisplay
                    If SLIDE_IsSolved() = True Then
                        MsgBox "Boom Chica Wow Waah!" & vbCrLf & vbCrLf & "You won the game in " & Trim$(Int((timeGetTime() - m_lngStartTime&) / 1000)) & " seconds!" & vbCrLf & vbCrLf & "You moved " & m_lngSlideMoves & " tiles!" & vbCrLf & vbCrLf & "Is that all you got?", vbInformation, gc_strAPPNAME
                        cmdEaster.Enabled = False
                        DoEvents
                        Call SLIDE_Reset
                        Call SLIDE_UpdateDisplay
                        Call SLIDE_RandomizeGrid
                        m_lngSlideMoves = 0
                        m_lngStartTime& = timeGetTime()
                        cmdEaster.Enabled = True
                    Else
                        m_lngSlideMoves = m_lngSlideMoves + 1
                    End If
                End If
            End If
        End If
    End If

End Sub
