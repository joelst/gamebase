VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   960
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   27
      Top             =   7560
      Width           =   5175
   End
   Begin VB.PictureBox picFont 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   960
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   810
      TabIndex        =   26
      Top             =   11280
      Visible         =   0   'False
      Width           =   12150
   End
   Begin VB.PictureBox picWave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   7560
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   5040
      Width           =   5175
   End
   Begin TabDlg.SSTab tabAbout 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "About"
      TabPicture(0)   =   "frmAbout.frx":241A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblURL"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPrompt(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPrompt(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVersion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPrompt(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProgName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCopyright"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblEmail"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgLogo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPrompt(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraStats"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Version History"
      TabPicture(1)   =   "frmAbout.frx":2436
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtHistory"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Database"
      TabPicture(2)   =   "frmAbout.frx":2452
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCredits"
      Tab(2).Control(1)=   "lblGBName"
      Tab(2).Control(2)=   "imgGameBase"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Credits"
      TabPicture(3)   =   "frmAbout.frx":246E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picView"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "License"
      TabPicture(4)   =   "frmAbout.frx":248A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtLicense"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtLicense 
         BackColor       =   &H8000000F&
         Height          =   3075
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   480
         Width           =   5295
      End
      Begin VB.PictureBox picView 
         BackColor       =   &H00000000&
         Height          =   3100
         Left            =   -74880
         ScaleHeight     =   203
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   363
         TabIndex        =   24
         Top             =   480
         Width           =   5500
      End
      Begin VB.TextBox txtCredits 
         BackColor       =   &H8000000F&
         Height          =   2775
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   65526
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtHistory 
         BackColor       =   &H8000000F&
         Height          =   3075
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   480
         Width           =   5295
      End
      Begin VB.Frame fraStats 
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   5235
         Begin VB.Label lblGameBase 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4935
            TabIndex        =   20
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   120
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loaded:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   705
         End
         Begin VB.Label lblFirstRunOn 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01 Jan 2002"
            Height          =   195
            Left            =   4170
            TabIndex        =   16
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   885
         End
         Begin VB.Label lblMPGame 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            Height          =   195
            Left            =   1545
            TabIndex        =   15
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   3510
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Most played game:"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblTotalSids 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4965
            TabIndex        =   13
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblTotalGames 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4965
            TabIndex        =   12
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total sids played:"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   11
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total games played:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   10
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1425
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First run on:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   9
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   825
         End
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks to everyone for design ideas and testing."
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         UseMnemonic     =   0   'False
         Width           =   3450
      End
      Begin VB.Label lblGBName 
         BackStyle       =   0  'Transparent
         Caption         =   "GameBase Name"
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
         Left            =   -72720
         TabIndex        =   21
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Image imgGameBase 
         Appearance      =   0  'Flat
         Height          =   3000
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1800
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   240
         Picture         =   "frmAbout.frx":24A6
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "James Burrows"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2040
         MouseIcon       =   "frmAbout.frx":2D70
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         Caption         =   "Copyright © 1997 - 2007"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1740
      End
      Begin VB.Label lblProgName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GameBase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblPrompt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Universal Retro-Gaming Frontend"
         Height          =   195
         Index           =   3
         Left            =   2775
         TabIndex        =   5
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   2685
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2.0 beta - Feb 20 2001"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         UseMnemonic     =   0   'False
         Width           =   2190
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "for updates."
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visit"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.bu22.com/"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   570
         MouseIcon       =   "frmAbout.frx":307A
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmAbout"
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
'QUICK CHANGE CONSTANTS
Private Const mc_lngTIMERWAIT As Long = 10 '12         'main loop timer interval value in milliseconds

'flag for stopping the loop and unloading the form
Private m_blnUnload As Boolean

'flag for saying we have already started
Private m_blnStarted As Boolean

'flag to say we can do the waves
Private m_blnDoWaves As Boolean

'---------
'Starfield
'---------

'star structure
Private Type StarType
    sngUniverseX As Single
    sngUniverseY As Single
    sngUniverseZ As Single
    sngRotationX As Single
    sngRotationY As Single
    sngRotationZ As Single
    sngScreenX As Single
    sngScreenY As Single
    sngVelocityZ As Single
    blnSparkle As Boolean
End Type
'number of stars
Private m_lngMaxStars As Long
'The array of stars
Private ma_udtStars() As StarType

'Camera structure
Private Type CameraType
    sngA As Single
    sngB As Single
    sngC As Single
    sngAVel As Single
    sngBVel As Single
    sngCVel As Single
    sngAAcc As Single
    sngBAcc As Single
    sngCAcc As Single
End Type
'The camera
Private m_udtCamera As CameraType

'Color Table
Private ma_lngStarColours(639)
Private ma_lngStarSparkle(639)

'Field Constants
Private Const mc_sngFIELD_MINX As Single = -480#
Private Const mc_sngFIELD_MAXX As Single = 480#
Private Const mc_sngFIELD_WID As Single = (mc_sngFIELD_MAXX - mc_sngFIELD_MINX)
Private Const mc_sngFIELD_MINY As Single = -480#
Private Const mc_sngFIELD_MAXY As Single = 480#
Private Const mc_sngFIELD_HEI As Single = (mc_sngFIELD_MAXY - mc_sngFIELD_MINY)
Private Const mc_sngFIELD_MINZ As Single = -128#
Private Const mc_sngFIELD_MAXZ As Single = 640#
Private Const mc_sngFIELD_LEN As Single = (mc_sngFIELD_MAXZ - mc_sngFIELD_MINZ)

'maximum random value
Private Const mc_RANDMAX As Long = 32766

'engine variables
Private m_sngStarSinA As Single
Private m_sngStarCosA As Single
Private m_sngStarSinB As Single
Private m_sngStarCosB As Single
Private m_sngStarSinC As Single
Private m_sngStarCosC As Single
Private m_sngStarMaxVelocity As Single
Private m_sngStarVelocity As Single
Private m_sngStarZoom As Single
Private m_sngStarNSX As Single
Private m_sngStarNSY As Single
Private m_sngFieldExpanse As Single
Private m_sngScrZoom As Single
Private m_lngScreenCentreX As Long
Private m_lngScreenCentreY As Long
Private m_lngScreenWidth As Long
Private m_lngScreenHeight As Long

'---------------------
'Fading-in Font Engine
'---------------------

'Font Curring Table
Private Const mc_lngFONTROWS As Long = 5
Private Const mc_lngMAXFONTROWS As Long = mc_lngFONTROWS - 1
Private Const mc_lngFONTCOLS As Long = 126
Private Const mc_lngMAXFONTCOLS As Long = mc_lngFONTCOLS - 1
Private Const mc_lngXDIVIDERCOLOUR As Long = vbRed
Private Const mc_lngSTARTINGLETTER = 32
Private Type FontCurringType
    lngStartX As Long
    lngWidth As Long
    lngEndX As Long
    lngStartY As Long
    lngEndY As Long
End Type
Private ma_udtFont(mc_lngMAXFONTCOLS, mc_lngMAXFONTROWS) As FontCurringType
Private m_lngFontPixelHeight As Long

'Text Wave Display
Private Const mc_lngLINESPACING As Long = 4
Private Const mc_lngMAXWAVEWAITREPS As Long = 250
Private Const mc_lngMAXWAVEPAUSEREPS As Long = 100
Private Const mc_sngWAVEWAITSTEP As Single = 0.3 'letters display faster/slower
Private Const mc_sngWAVEBRIGHTSTEP As Single = 0.3 'whole wave fades in quicker/slower
Private Const mc_sngWAVEFADEOUTSPEED As Single = 0.3 'speed of wave fadeout
Private Type WaveCharacterType
    lngDestX As Long
    lngDestY As Long
    lngWidth As Long
    lngAsciiValue As Long
    sngBrightnessRow As Single
    blnFadingIn As Boolean
End Type
Private Type WaveDisplayType
    udtChars() As WaveCharacterType
    lngMaxChars As Long
    lngBlitHeight As Long
    lngBlitWidth As Long
    lngBlitX As Long
    lngBlitY As Long
    lngLastFaderIdx As Long
    lngWaitCount As Long
    blnFadingIn As Boolean
End Type
Private ma_udtWaves() As WaveDisplayType
Private m_lngMaxWaves As Long
Private m_lngCurrentWave As Long
Private m_lngWavePause As Long

'for blitting the font
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hdcSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean
Private Sub DoCredits()

    '-------------------
    'Main Starfield loop
    '-------------------
    
    Dim lngNewTime&         'Timer variable
    Dim lngLastTime&        'Timer variable
    
    'initialise variables
    lngLastTime& = 0

    'start the loop until the form is unloaded
    Do Until m_blnUnload = True
        lngNewTime& = GetTickCount()
        If lngNewTime& > lngLastTime& Then
            lngLastTime& = lngNewTime& + mc_lngTIMERWAIT
            
            '''''''''''''''''''''''''''''''
            'time to process the next frame
            '''''''''''''''''''''''''''''''
            
            'clear the buffer
            picBack.Cls
            
            'update the starfield
            Call UpdateStarField
            
            'update the waves
            If m_blnDoWaves = True Then
                If m_lngWavePause = 0 Then
                    Call UpdateWaves
                Else
                    If m_lngWavePause = mc_lngMAXWAVEPAUSEREPS Then
                        m_lngWavePause = 0
                    Else
                        m_lngWavePause = m_lngWavePause + 1
                    End If
                End If
            End If
            
            'flip to viewable surface
            picView.Picture = picBack.Image
            
        End If
        'allow the form to unload if needed and allow user to change tabs
        DoEvents
    Loop

End Sub
Private Sub ResetWave(ByVal lngWaveIdx)

    '--------------------------------------------------------
    'Resets a wave's brightness/fade and waitcount properties
    '--------------------------------------------------------
    
    Dim lngCharIdx As Long              'loop index
    Dim sngBrightness As Single         'brightness value
    Dim intRandomStyle As Integer       'fade style
    Dim lngLastChar As Long             'for finding last char to fade out
    Dim a_sngRandomTable() As Single    'for the random style
    Dim lngMaxRnd As Long               'for the random style
    Dim lngPickedRndIdx As Long         'for the random style
    Dim lngTempRndVal As Long           'for the random style
    Dim lngMiddle As Long               'for ends to mddle/middle to ends style
    
    'reset each character in the wave
    With ma_udtWaves(lngWaveIdx)
    
        'get which style to do
        intRandomStyle = Int(4 * Rnd) '0 to 3
        
        Select Case intRandomStyle
            
            Case 0 '1st char to last char
            
                sngBrightness = -1
                For lngCharIdx = 0 To .lngMaxChars
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        .sngBrightnessRow = sngBrightness
                    End With
                    sngBrightness = sngBrightness - mc_sngWAVEWAITSTEP
                Next lngCharIdx
                .lngLastFaderIdx = .lngMaxChars
                .lngWaitCount = 0
                .blnFadingIn = True
            
            Case 1 'ends to middle
            
                lngMiddle = .lngMaxChars / 2
                
                sngBrightness = -1
                For lngCharIdx = 0 To (lngMiddle - 1)
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        .sngBrightnessRow = sngBrightness
                    End With
                    sngBrightness = sngBrightness - (mc_sngWAVEWAITSTEP * 2)
                Next lngCharIdx
                sngBrightness = (-1) + mc_sngWAVEWAITSTEP
                For lngCharIdx = .lngMaxChars To lngMiddle Step -1
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        .sngBrightnessRow = sngBrightness
                    End With
                    sngBrightness = sngBrightness - (mc_sngWAVEWAITSTEP * 2)
                Next lngCharIdx
                
                .lngLastFaderIdx = lngMiddle
                .lngWaitCount = 0
                .blnFadingIn = True
                
            Case 2 'middle to ends
            
                lngMiddle = .lngMaxChars / 2
                
                sngBrightness = -1
                For lngCharIdx = (lngMiddle - 1) To 0 Step -1
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        .sngBrightnessRow = sngBrightness
                    End With
                    sngBrightness = sngBrightness - (mc_sngWAVEWAITSTEP * 2)
                Next lngCharIdx
                sngBrightness = (-1) + mc_sngWAVEWAITSTEP
                For lngCharIdx = lngMiddle To .lngMaxChars
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        .sngBrightnessRow = sngBrightness
                    End With
                    sngBrightness = sngBrightness - (mc_sngWAVEWAITSTEP * 2)
                Next lngCharIdx
                
                .lngLastFaderIdx = .lngMaxChars
                .lngWaitCount = 0
                .blnFadingIn = True
            
            Case 3 'random

                'fill the random table
                ReDim a_sngRandomTable(.lngMaxChars)
                For lngCharIdx = 0 To .lngMaxChars
                    a_sngRandomTable(lngCharIdx) = -((mc_sngWAVEWAITSTEP * lngCharIdx) + 1)
                Next lngCharIdx
                lngLastChar = 0
                lngMaxRnd = UBound(a_sngRandomTable)
                'loop through each char
                For lngCharIdx = 0 To .lngMaxChars
                    With .udtChars(lngCharIdx)
                        .blnFadingIn = True
                        'begin random jiggerypokery
                        lngPickedRndIdx = Int((lngMaxRnd + 1) * Rnd)
                        .sngBrightnessRow = a_sngRandomTable(lngPickedRndIdx)
                        lngTempRndVal = .sngBrightnessRow
                        a_sngRandomTable(lngPickedRndIdx) = a_sngRandomTable(lngMaxRnd)
                        a_sngRandomTable(lngMaxRnd) = lngTempRndVal
                        lngMaxRnd = lngMaxRnd - 1
                        If .sngBrightnessRow < ma_udtWaves(lngWaveIdx).udtChars(lngLastChar).sngBrightnessRow Then
                            lngLastChar = lngCharIdx
                        End If
                        'end random jiggerypokery
                    End With
                    sngBrightness = sngBrightness - mc_sngWAVEWAITSTEP
                Next lngCharIdx
                .lngLastFaderIdx = lngLastChar
                .lngWaitCount = 0
                .blnFadingIn = True
                
        End Select
    
    End With

End Sub

Private Sub ResetWaves()

    '---------------------------------------------------------------
    'Resets the current wave to 0 and resets all brightness counters
    'also sets the state to pause
    '---------------------------------------------------------------

    Dim lngWaveIdx As Long      'loop index
    
    'reset module variables
    m_lngCurrentWave = 0
    m_lngWavePause = 1

    'reset each wave
    For lngWaveIdx = 0 To m_lngMaxWaves
        Call ResetWave(lngWaveIdx)
    Next lngWaveIdx

End Sub

Private Sub UpdateStarField()

    '-------------------
    'Main Starfield loop
    '-------------------
    
    Dim lngStarIdx As Long  'loop index
    Dim sngTempRotationX    'Temporary storage for rotation calculation
    Dim sngZ As Single
    Dim lngColour As Long

    'set the Screen Zoom and star velocity
    m_sngScrZoom = CSng(m_lngScreenCentreX) * m_sngStarZoom
    m_sngStarVelocity = m_sngStarVelocity + ((m_sngStarMaxVelocity - m_sngStarVelocity) * 0.01)

    'rotation calculations
    With m_udtCamera
        .sngAAcc = (.sngAAcc + RangeRand(-0.00001, 0.00002)) * 0.99
        .sngBAcc = (.sngBAcc + RangeRand(-0.00001, 0.00002)) * 0.99
        .sngCAcc = (.sngCAcc + RangeRand(-0.00001, 0.00002)) * 0.992
        .sngAVel = (.sngAVel + .sngAAcc) * 0.99
        .sngBVel = (.sngBVel + .sngBAcc) * 0.99
        .sngCVel = (.sngCVel + .sngCAcc) * 0.992
        .sngA = (.sngA + .sngAVel) * 0.99
        .sngB = (.sngB + .sngBVel) * 0.99
        .sngC = (.sngC + .sngCVel) * 0.992
        m_sngStarSinA = CSng(Sin(-.sngA))
        m_sngStarCosA = CSng(Cos(-.sngA))
        m_sngStarSinB = CSng(Sin(-.sngB))
        m_sngStarCosB = CSng(Cos(-.sngB))
        m_sngStarSinC = CSng(Sin(-.sngC))
        m_sngStarCosC = CSng(Cos(-.sngC))
    End With
    
    'loop through each star
    For lngStarIdx = 0 To m_lngMaxStars
        With ma_udtStars(lngStarIdx)
            'move star closer
            .sngUniverseZ = .sngUniverseZ - (.sngVelocityZ * m_sngStarVelocity)
            If .sngUniverseZ < mc_sngFIELD_MINZ Then
                Call ResetStar(lngStarIdx)
            End If
            'rotate Y
            .sngRotationX = .sngUniverseX * m_sngStarCosA + .sngUniverseZ * m_sngStarSinA
            .sngRotationZ = .sngUniverseZ * m_sngStarCosA - .sngUniverseX * m_sngStarSinA
            'rotate X
            .sngRotationY = .sngUniverseY * m_sngStarCosB + .sngRotationZ * m_sngStarSinB
            .sngRotationZ = .sngRotationZ * m_sngStarCosB - .sngUniverseY * m_sngStarSinB
            'rotate Z
            sngTempRotationX = .sngRotationX * m_sngStarCosC + .sngRotationY * m_sngStarSinC
            .sngRotationY = .sngRotationY * m_sngStarCosC - .sngRotationX * m_sngStarSinC
            .sngRotationX = sngTempRotationX
            'calculate projected screen coordinates
            m_sngStarNSX = CSng(m_lngScreenCentreX) + .sngRotationX / .sngRotationZ * m_sngScrZoom
            m_sngStarNSY = CSng(m_lngScreenCentreY) + .sngRotationY / .sngRotationZ * m_sngScrZoom
            'clip to picture box if necessary
            If (.sngRotationZ > 1#) And (.sngRotationZ < 640#) Then
                If (.sngScreenX > 2) And (.sngScreenX < m_lngScreenWidth) And (.sngScreenY > 2) And (.sngScreenY < m_lngScreenHeight) And (m_sngStarNSX > 2) And (m_sngStarNSX < m_lngScreenWidth) And (m_sngStarNSY > 2) And (m_sngStarNSY < m_lngScreenHeight) Then
                    'plot star
                    If .blnSparkle Then
                        lngColour = ma_lngStarSparkle(Int(.sngRotationZ))
                    Else
                        lngColour = ma_lngStarColours(Int(.sngRotationZ))
                    End If
                    sngZ = .sngRotationZ
                    If sngZ > 480 Then
                        picBack.PSet (.sngScreenX, .sngScreenY), lngColour
                    ElseIf sngZ > 320 Then
                        picBack.Line (.sngScreenX, .sngScreenY)-(.sngScreenX + 1, .sngScreenY + 1), lngColour, BF
                    ElseIf sngZ > 160 Then
                        picBack.Line (.sngScreenX, .sngScreenY)-(.sngScreenX + 2, .sngScreenY + 2), lngColour, BF
                        picBack.PSet (.sngScreenX, .sngScreenY), &H555555
                        picBack.PSet (.sngScreenX + 2, .sngScreenY), &H555555
                        picBack.PSet (.sngScreenX, .sngScreenY + 2), &H555555
                        picBack.PSet (.sngScreenX + 2, .sngScreenY + 2), &H555555
                    Else
                        picBack.Line (.sngScreenX - 1, .sngScreenY - 1)-(.sngScreenX + 2, .sngScreenY + 2), lngColour, BF
                        picBack.PSet (.sngScreenX - 1, .sngScreenY - 1), &H555555
                        picBack.PSet (.sngScreenX + 2, .sngScreenY - 1), &H555555
                        picBack.PSet (.sngScreenX - 1, .sngScreenY + 2), &H555555
                        picBack.PSet (.sngScreenX + 2, .sngScreenY + 2), &H555555
                    End If
                End If
            End If
            'store the new star position
            .sngScreenX = m_sngStarNSX
            .sngScreenY = m_sngStarNSY
        End With
    Next lngStarIdx

End Sub


Private Function GetNumOfCRandLF(strWave_in As String) As Long

    '-----------------------------------------------------------
    'calculates the number of carriage
    'returns and line feeds in the passed in string
    '
    'Inputs:        strWave_in      String to check
    '
    'Returns:       Number of CRLF chars sequences in the string
    '-----------------------------------------------------------
    
    Dim lngTempCount As Long        'temporary counter
    Dim lngStart As Long            'starting position for instr search
    Dim lngFoundPos As Long
    
    'initialise counter
    lngTempCount = 0
    
    'Count # of Carriage Returns Line Feed pairs
    lngStart = 1
    Do
        lngFoundPos = InStr(lngStart, strWave_in, vbCrLf)
        If lngFoundPos > 0 Then
            lngStart = lngFoundPos + 2
            lngTempCount = lngTempCount + 1
        Else
            Exit Do
        End If
    Loop
    
    'transfer to return value
    GetNumOfCRandLF = lngTempCount

End Function

Private Sub InitStarField()

    '--------------------------------------------------
    'Initialises the variables needed for the starfield
    '--------------------------------------------------
    
    Dim lngStarIdx As Long                      'loop index
    Dim lngColourIdx As Long                    'loop index
    Dim lngIntensity As Long                    'color intensity
    Dim lngNumStars As Long                     'total number of stars
    Dim lngRndAmount As Long
    Dim lngRndDirection As Long
    Const c_lngSTARTINTENSITY As Long = 255     'white
    
    'dimension the array
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\Options", "AboutScreenStars", lngNumStars, False, 512
    If lngNumStars < 64 Then
        lngNumStars = 64
    ElseIf lngNumStars > 4096 Then
        lngNumStars = 4096
    End If
    m_lngMaxStars = lngNumStars - 1
    ReDim ma_udtStars(m_lngMaxStars)

    'Initialise the stars
    For lngStarIdx = 0 To m_lngMaxStars
        Call ResetStar(lngStarIdx)
        ma_udtStars(lngStarIdx).sngUniverseZ = RangeRand(0#, 640#)
    Next lngStarIdx
    
    'Initialise the camera
    With m_udtCamera
        .sngA = 0
        .sngAAcc = 0
        .sngAVel = 0
        .sngB = 0
        .sngBAcc = 0
        .sngBVel = 0
        .sngC = 0
        .sngCAcc = 0
        .sngCVel = 0
    End With
    
    'Initialise the colour tables
    lngIntensity = c_lngSTARTINTENSITY
    ma_lngStarColours(0) = RGB(lngIntensity, lngIntensity, lngIntensity)
    For lngColourIdx = 1 To 639
        ma_lngStarColours(lngColourIdx) = RGB(lngIntensity, lngIntensity, lngIntensity)
        If lngColourIdx Mod 4 = 0 Then
            lngIntensity = lngIntensity - 1
        End If
    Next lngColourIdx
    lngIntensity = c_lngSTARTINTENSITY
    ma_lngStarSparkle(0) = RGB(lngIntensity, lngIntensity, lngIntensity)
    For lngColourIdx = 1 To 639
        If Int((3 - 0 + 1) * Rnd + 0) = 0 Then
            'give it it's normal colour
            ma_lngStarSparkle(lngColourIdx) = RGB(lngIntensity, lngIntensity, lngIntensity)
        Else
            'randomise the colour a bit
            lngRndAmount = Int((48 - 0 + 1) * Rnd + 0)
            lngRndDirection = Int((1 - 0 + 1) * Rnd + 0)
            If lngRndDirection = 0 Then
                'down
                If lngIntensity - lngRndAmount < 0 Then
                    lngRndAmount = 0
                End If
                lngRndAmount = -lngRndAmount
            Else
                'up
                If lngIntensity + lngRndAmount > 255 Then
                    lngRndAmount = 255
                End If
            End If
            ma_lngStarSparkle(lngColourIdx) = RGB(lngIntensity + lngRndAmount, lngIntensity + lngRndAmount, lngIntensity + lngRndAmount)
        End If
        If lngColourIdx Mod 4 = 0 Then
            lngIntensity = lngIntensity - 1
        End If
    Next lngColourIdx


    'Initialise other variables
    m_sngStarMaxVelocity = 20#
    m_sngStarVelocity = 0#
    m_sngStarZoom = (3# / 2#)
    m_sngFieldExpanse = 1.5

End Sub

Private Function RangeRand(ByVal sngMin As Single, sngRange As Single) As Single

    '------------------------------------------------------------
    'Calculates a single value between sngMin and sngMin+sngRange
    '
    'Inputs:        sngMin          The minimum returnable value
    '               sngRange        The Range of valid values
    '
    'Returns:       The Calculated Random value
    '------------------------------------------------------------
    
    Dim intRand As Integer  'result of Rnd() operation
    
    ' #define RangeRand(min,range) (min + (float)rand()/(float)RAND_MAX*range)
    
    'get the random value
    intRand = Int((mc_RANDMAX + 1) * Rnd)

    'calculate the value in the range
    RangeRand = sngMin + CSng(intRand) / CSng(mc_RANDMAX) * sngRange

End Function

Private Sub ResetStar(ByVal lngStarIdx As Long)

    '---------------------------------------------------------
    'Resets a star to the back of the universe
    'and plonks it at a random X and Y position
    '
    'Inputs:        lngStarIdx      Index of the star to reset
    '---------------------------------------------------------
    
    Dim intRand As Integer  'result of Rnd() operation
        
    'get the random value
    intRand = Int((mc_RANDMAX + 1) * Rnd)
    
    'reset the star
    With ma_udtStars(lngStarIdx)
        .sngVelocityZ = 1# / (0.1 + 9.9 * CSng(intRand) / CSng(mc_RANDMAX)) ^ (m_sngFieldExpanse)
        .sngUniverseX = RangeRand(mc_sngFIELD_MINX, mc_sngFIELD_WID)
        .sngUniverseY = RangeRand(mc_sngFIELD_MINY, mc_sngFIELD_HEI)
        .sngUniverseZ = mc_sngFIELD_MAXZ
        .blnSparkle = (Int((2 - 0 + 1) * Rnd + 0) = 0)
    End With

End Sub

Private Sub UpdateWaves()

    '-------------------------------------------------------
    'Updates the Waves to show the next frame of fading text
    '-------------------------------------------------------
    
    Dim lngCharIdx As Long              'loop index
    Dim lngFadeIdx As Long              'loop index
    
    'clear the work area
    picWave.Picture = LoadPicture()

    'update the wave
    With ma_udtWaves(m_lngCurrentWave)
        If .blnFadingIn = True Then
            If .lngWaitCount = 0 Then
                'FADING IN
                For lngCharIdx = 0 To .lngMaxChars
                    With .udtChars(lngCharIdx)
                        If .sngBrightnessRow >= 0 And .sngBrightnessRow <= mc_lngMAXFONTROWS Then
                            'blit
                            Call BitBlt(picWave.hDC, .lngDestX, .lngDestY, .lngWidth, m_lngFontPixelHeight, picFont.hDC, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartX, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartY, SRCCOPY)
                        End If
                        .sngBrightnessRow = .sngBrightnessRow + mc_sngWAVEBRIGHTSTEP
                        If .sngBrightnessRow >= mc_lngMAXFONTROWS Then
                            .sngBrightnessRow = mc_lngMAXFONTROWS
                            'just reached brightest
                            If lngCharIdx = ma_udtWaves(m_lngCurrentWave).lngLastFaderIdx Then
                                ma_udtWaves(m_lngCurrentWave).lngWaitCount = 1
                            End If
                        End If
                    End With
                Next lngCharIdx
            Else
                'WAITING TO FADE OUT
                .lngWaitCount = .lngWaitCount + 1
                For lngCharIdx = 0 To .lngMaxChars
                    With .udtChars(lngCharIdx)
                        Call BitBlt(picWave.hDC, .lngDestX, .lngDestY, .lngWidth, m_lngFontPixelHeight, picFont.hDC, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartX, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartY, SRCCOPY)
                    End With
                Next lngCharIdx
                If .lngWaitCount >= mc_lngMAXWAVEWAITREPS Then
                    'START FADING OUT
                    .blnFadingIn = False
                    For lngCharIdx = 0 To .lngMaxChars
                        .udtChars(lngCharIdx).blnFadingIn = False
                    Next lngCharIdx
                End If
            End If
        Else
            'FADING OUT
            For lngFadeIdx = 0 To .lngMaxChars
                With .udtChars(lngFadeIdx)
                    If .sngBrightnessRow >= 0 Then
                        'blit
                        Call BitBlt(picWave.hDC, .lngDestX, .lngDestY, .lngWidth, m_lngFontPixelHeight, picFont.hDC, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartX, ma_udtFont(.lngAsciiValue, .sngBrightnessRow).lngStartY, SRCCOPY)
                    End If
                    .sngBrightnessRow = .sngBrightnessRow - mc_sngWAVEFADEOUTSPEED
                End With
            Next lngFadeIdx
        End If
        
        'now transparently blit the wave buffer onto the back buffer
        picWave.Picture = picWave.Image
        Call TransparentBlt(picBack.hDC, .lngBlitX, .lngBlitY, .lngBlitWidth, .lngBlitHeight, picWave.hDC, .lngBlitX, .lngBlitY, .lngBlitWidth, .lngBlitHeight, 0)
        
        'end of current wave?
        If (.udtChars(.lngLastFaderIdx).blnFadingIn = False) And (.udtChars(.lngMaxChars).sngBrightnessRow <= 0) Then
            m_lngWavePause = 1
            Call ResetWave(m_lngCurrentWave)
            m_lngCurrentWave = m_lngCurrentWave + 1
            If m_lngCurrentWave > m_lngMaxWaves Then
                m_lngCurrentWave = 0
            End If
        End If
        
    End With

End Sub

Private Sub GetINIMessages()

    '----------------------------------------------------
    'Gets the language specific text for the About Screen
    'from the GBText.ini file and sets the controls up
    'to utilise the got messages
    '----------------------------------------------------

    Caption = g_xobjFunc.GetINIMessage("About Screen", "About", "About") & " " & DecMess("dr`Adl`F") & " v" & g_strAPPVERSION & IIf((g_blnBETA), " Beta", "") & "..."
    With tabAbout
        .TabCaption(0) = g_xobjFunc.GetINIMessage("About Screen", "Tab_About_Title", "&About")
        .TabCaption(1) = g_xobjFunc.GetINIMessage("About Screen", "Tab_VersionHistory_Title", "&Version history")
    End With

End Sub

Private Sub Form_Activate()

    If (Not m_blnStarted) And (Not m_blnUnload) Then
        m_blnStarted = True
        Call DoCredits
    End If

End Sub

Private Sub Form_Load()

    Dim strDate As String
    Dim lngTotalPlayed As Long
    Dim strMostPlayed As String
    Const DesignScaleHeight As Single = 3975    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6015     'The required scalewidth for the form
    Dim intFileNum As Integer
    Dim strHistory As String
    Dim strLicense As String
    
    'for all the fancy stuff
    Randomize
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'flag to say we haven't quit yet
    m_blnUnload = False
    'flag to say we haven't started yet
    m_blnStarted = False
    'we're not on the Credits tab yet!
    m_blnDoWaves = False

    'screen dimensions
    With picView
        m_lngScreenCentreX = .ScaleWidth / 2
        m_lngScreenCentreY = .ScaleHeight / 2
        m_lngScreenWidth = .ScaleWidth
        m_lngScreenHeight = .ScaleHeight
    End With

    'load the version history file into the textbox
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "History.txt", False, "") = True Then
        intFileNum = FreeFile()
        Open g_strAppPathWithBackslash & "History.txt" For Input As #intFileNum
        strHistory = Input$(LOF(intFileNum), intFileNum)
        Close #intFileNum
        If Len(strHistory) > gc_lngMAXMEMOTEXTLENGTH Then
            strHistory = Left$(strHistory, gc_lngMAXMEMOTEXTLENGTH)
        End If
        txtHistory.Text = strHistory
    Else
        txtHistory.Text = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!") & " (" & g_strAppPathWithBackslash & "History.txt)"
    End If

    'load the version license file into the textbox
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "License.txt", False, "") = True Then
        intFileNum = FreeFile()
        Open g_strAppPathWithBackslash & "License.txt" For Input As #intFileNum
        strLicense = Input$(LOF(intFileNum), intFileNum)
        Close #intFileNum
        If Len(strLicense) > gc_lngMAXMEMOTEXTLENGTH Then
            strLicense = Left$(strLicense, gc_lngMAXMEMOTEXTLENGTH)
        End If
        txtLicense.Text = strLicense
    Else
        txtLicense.Text = g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!") & " (" & g_strAppPathWithBackslash & "License.txt)"
    End If

    'get the language specific text from the .ini files
    Call GetINIMessages

    'set labels (this encode/decode stuff is pointless now the frontend is open source)
    lblProgName = DecMess("dr`Adl`F")
    lblPrompt(3) = DecMess("cmdsmnqEfmhl`F,nqsdQk`rqduhmTdgS")
    lblCopyright = DecMess("/0/1,8880¨sgfhqxonB")
    lblEmail = DecMess("rvnqqtArdl`I")
    lblURL = DecMess(".lnb-11ta-vvv..9ossg")
    lblGameBase = g_strDatabaseName
    tabAbout.TabCaption(2) = "&Database"
    tabAbout.TabCaption(4) = "&License"
    lblVersion = "v" & g_strAPPVERSION & IIf((g_blnBETA), " Beta", "") & " (db v" & Format$(g_sngDBVersion, "0.0") & ") - " & gc_strRELEASEDATE
    lblPrompt(12) = g_xobjFunc.GetINIMessage("About Screen", "FirstRunOn", "Since") & ":"
    lblPrompt(13) = g_xobjFunc.GetINIMessage("About Screen", "TotalGamesPlayed", "Total games played") & ":"
    lblPrompt(14) = g_xobjFunc.GetINIMessage("About Screen", "TotalMusicPlayed", "Total Music played") & ":"
    lblPrompt(15) = g_xobjFunc.GetINIMessage("About Screen", "MostPlayedGame", "Most played game") & ":"
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "InstallDate", strDate, True, Format$(Now, "dd-mm-yyyy")
    lblFirstRunOn = GetFormattedDate(strDate)
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalGamesPlayed", lngTotalPlayed, False, 0
    lblTotalGames = Trim$(lngTotalPlayed)
    g_xobjRegistry.GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Stats", "TotalMusicPlayed", lngTotalPlayed, False, 0
    lblTotalSids = Trim$(lngTotalPlayed)
    Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT TOP 1 TimesPlayed, Name FROM Games ORDER BY TimesPlayed DESC", dbOpenSnapshot)
    With g_rsGames
        If .EOF Then
            strMostPlayed = g_strNone
        Else
            .MoveFirst
            If !TimesPlayed = 0 Then
                strMostPlayed = g_strNone
            Else
                strMostPlayed = EllipseTextEnd(!Name, lblMPGame.Width)
            End If
        End If
        .Close
    End With

    lblMPGame = strMostPlayed
    'gamebase tab
    lblGBName = g_strDatabaseName ' & " (" & g_xobjFunc.GetINIMessage("About Screen", "Update", "#upd") & " " & g_lngOfficialDBUpdate & ")"
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & g_strGameBaseName & "\Credits.txt", False, "") = True Then
        intFileNum = FreeFile()
        Open g_strAppPathWithBackslash & g_strGameBaseName & "\Credits.txt" For Input As #intFileNum
        strHistory = Input$(LOF(intFileNum), intFileNum)
        Close #intFileNum
        If Len(strHistory) > gc_lngMAXMEMOTEXTLENGTH Then
            strHistory = Left$(strHistory, gc_lngMAXMEMOTEXTLENGTH)
        End If
        txtCredits.Text = strHistory
    Else
        txtCredits.Text = ""
    End If
    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & g_strGameBaseName & "\Gfx\AboutScreen_120x200.gif", False, "") = True Then
        imgGameBase.Picture = LoadPicture(g_strAppPathWithBackslash & g_strGameBaseName & "\Gfx\AboutScreen_120x200.gif")
    Else
        imgGameBase.Visible = False
        lblGBName.Left = imgGameBase.Left
        txtCredits.Left = imgGameBase.Left
        txtCredits.Width = 5295
        lblGBName.Width = 5295
    End If

    'Initialise the starfield
    Call InitStarField
    
    'get the font pixel height
    m_lngFontPixelHeight = picFont.ScaleHeight / mc_lngFONTROWS
    'read in the credits font curring table
    Call ReadInCreditsFontTable
    'read in the credit text waves
    Call ReadInCreditTextWaves
    'reset them!
    Call ResetWaves

End Sub
Private Sub ReadInCreditsFontTable()

    '---------------------------------------------
    'Reads in the Credits Font Curring table.
    'i.e. calculate and store the offsets and
    'widths for each character in the bitmap font.
    '
    'Remarks:       The font bitmap is stored
    '               in the picFont picture-box and
    '               is separated into rows and
    '               columns.  The rows are going
    '               from top to bottom (bright to
    '               dark).  The columns are
    '               separated with a single pixel
    '               column of a specific colour.
    '---------------------------------------------
    
    Dim lngCharIdx As Long      'loop index
    Dim lngRowIdx As Long       'loop index
    Dim lngPixelXIdx As Long    'loop index
    Dim lngPixelXMax As Long    'total x-pixels in font bitmap
    Dim lngStartX As Long       'xpos-start of each character
    Dim lngWidth As Long        'width counter for each character

    'set the StartY/EndY properties for each character
    For lngCharIdx = mc_lngSTARTINGLETTER To mc_lngMAXFONTCOLS
        For lngRowIdx = 0 To mc_lngMAXFONTROWS
            With ma_udtFont(lngCharIdx, lngRowIdx)
                .lngStartY = lngRowIdx * m_lngFontPixelHeight
                .lngEndY = (.lngStartY + m_lngFontPixelHeight) - 1
            End With
        Next lngRowIdx
    Next lngCharIdx

    'calculate the StartX/Width/EndX properties for each character
    'and store in the first (top) row of the font array
    lngPixelXMax = picFont.ScaleWidth - 1
    lngStartX = 0
    lngWidth = 0
    lngCharIdx = mc_lngSTARTINGLETTER - 1 'character BEFORE mc_lngSTARTINGLETTER
    'loop through each pixel and when we find a divider pixel
    'save the info for the PREVIOUSLY TRAVERSED character
    For lngPixelXIdx = 0 To lngPixelXMax
        If picFont.Point(lngPixelXIdx, 0) = mc_lngXDIVIDERCOLOUR Then
            'found divider: save Width and calculate StartX and EndX
            'NOTE: 1st time it comes in here it's for the character BEFORE mc_lngSTARTINGLETTER
            With ma_udtFont(lngCharIdx, 0)
                .lngWidth = lngWidth
                .lngStartX = lngStartX
                .lngEndX = (lngStartX + lngWidth) - 1
            End With
            'reset variables ready for next letter
            lngStartX = lngStartX + 1 'divider gap
            lngStartX = lngStartX + lngWidth
            lngWidth = 0
            'move to next character
            lngCharIdx = lngCharIdx + 1
        Else
            'divider not found: amend Width
            lngWidth = lngWidth + 1
        End If
    Next lngPixelXIdx
    'store the last character's info as we didn't get a divider as the last pixel
    With ma_udtFont(mc_lngMAXFONTCOLS, 0)
        .lngWidth = lngWidth
        .lngStartX = lngStartX
        .lngEndX = (lngStartX + lngWidth) - 1
    End With
    
    'copy the calculated StartX/Width/EndX properties
    'to the other rows in the font array
    For lngCharIdx = mc_lngSTARTINGLETTER To mc_lngMAXFONTCOLS
        For lngRowIdx = 1 To mc_lngMAXFONTROWS
            With ma_udtFont(lngCharIdx, lngRowIdx)
                .lngWidth = ma_udtFont(lngCharIdx, 0).lngWidth
                .lngStartX = ma_udtFont(lngCharIdx, 0).lngStartX
                .lngEndX = ma_udtFont(lngCharIdx, 0).lngEndX
            End With
        Next lngRowIdx
    Next lngCharIdx

End Sub
Private Sub ReadInCreditTextWaves()

    '-----------------------------------------------
    'Reads in the Credits Text Waves and sets up
    'the arrays for the fade in and out effects.
    'This routine calculates the starting offsets
    'and fade values for every character in every
    'line of every wave so that the runtime engine
    'has less to do.
    '
    'Assumptions:       Each Wave string cannot end
    '                   with a single Carriage
    '                   Return or Line Feed or a
    '                   combination of both (vbCRLF)
    '                   It must end with a printable
    '                   character.  Also, CR and LF
    '                   pairs must be together and
    '                   not split.
    '-----------------------------------------------
    
    Const c_lngTOTALWAVES As Long = 53                  'total text waves
    Const c_lngMAXWAVES As Long = c_lngTOTALWAVES - 1   'max text waves (array UBound)
    Const NL As String = vbCrLf                         'new-line
    Const BL As String = vbCrLf & vbCrLf                '2 x new-line (inserts blank line)
    Dim lngAscCR As Long                                'asc value of carriage return
    Dim lngAscLF As Long                                'asc value of line feed
    Dim lngAscChar As Long                              'asc value of char
    Dim a_strWaves(c_lngMAXWAVES) As String             'array of waves text
    Dim lngWaveIdx As Long                              'loop index
    Dim lngStringCharIdx As Long                        'loop index
    Dim lngLineIdx As Long                              'loop index
    Dim lngArrayCharIdx As Long                         'index in char array
    Dim lngWaveLen As Long                              'length of text wave string
    Dim lngNumCRLF As Long                              'number of CRLFs in string
    Dim lngLinesInWave As Long                          'Numver of lines in wave text
    Dim lngCharDestTop As Long                          'Destination position on screen for char
    Dim a_lngPixelWidth() As Long                       'pixel width of lines in a wave
    Dim a_lngCharsInLine() As Long                      'number of chars in each line of a wave
    Dim lngCurrentLine As Long                          'current line in wave
    Dim lngLineStartX As Long                           'x-starting position for each line in a wave
    Dim lngCharIdx As Long                              'index for setting x-pos of each char
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Setup the text for each wave to prepare for reading into the runtime array
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'setup the text for each wave
    a_strWaves(0) = "Jimbo" & NL & "presents"
    a_strWaves(1) = "A" & BL & "BU22.COM" & BL & "Production"
    a_strWaves(2) = "== GameBase ==" & BL & "The Universal Emulator Frontend" & NL & "and Database Utility"
    a_strWaves(3) = "Concept, design" & NL & "and programming" & NL & "by Jimbo"
    a_strWaves(4) = "Additional programming" & NL & "by Brataccas"
    a_strWaves(5) = "Artwork by" & NL & "Jacek Bogucki"
    a_strWaves(6) = "Splash Screen" & NL & "by user42 and .mad."
    a_strWaves(7) = "Greets and thanks go out to..." & NL & "(in no particular order)"
    a_strWaves(8) = ".mad."
    a_strWaves(9) = "K.C."
    a_strWaves(10) = "AAG"
    a_strWaves(11) = "Sut"
    a_strWaves(12) = "user42"
    a_strWaves(13) = "Sjakie43"
    a_strWaves(14) = "Vicman"
    a_strWaves(15) = "hardmanm"
    a_strWaves(16) = "SkydivinGirl"
    a_strWaves(17) = "ldaneels"
    a_strWaves(18) = "Underfly"
    a_strWaves(19) = "Retroplay"
    a_strWaves(20) = "Arthur"
    a_strWaves(21) = "Mr Borg and The GB64 Team"
    a_strWaves(22) = "Belgarath and eLowar"
    a_strWaves(23) = "JohnCKirk"
    a_strWaves(24) = "fgerbig"
    a_strWaves(25) = "Per Bolmstedt (SOMCSB!)"
    a_strWaves(26) = "Kim Lemon"
    a_strWaves(27) = "Igor Pavlov (7-Zip author)"
    a_strWaves(28) = "Slay Radio (you guys rock)"
    a_strWaves(29) = "Seriously, go there now:" & NL & "www.slayradio.org"
    a_strWaves(30) = "Tone"
    a_strWaves(31) = "Jinjoid"
    a_strWaves(32) = "All the C64 Remixers"
    a_strWaves(33) = "My fellow BYOAC'ers"
    a_strWaves(34) = "That person who sent me" & NL & "the Amiga Forever CD" & NL & "(whoever you are)"
    a_strWaves(35) = "Kate and Lukey :)"
    a_strWaves(36) = "Anyone who's ever made" & NL & "a GameBase database"
    a_strWaves(37) = "and everyone else who has" & NL & "supported GameBase!"
    a_strWaves(38) = "Some interesting trivia" & NL & "(or not)..."
    a_strWaves(39) = "There is one easter egg" & NL & "in this version" & NL & "of GameBase..."
    a_strWaves(40) = "...which should be easy to find" & NL & "since it's now open source!"
    a_strWaves(41) = "GameBase might have" & NL & "been called Play64, C64GS," & NL & "Archive64, Cat64, C64DB," & NL & "CBMBase, Click64 or Cafe64!"
    a_strWaves(42) = "Martin Pugh (Pugsy) was the" & NL & "first ever beta tester" & NL & "of GameBase (then gamebase64)"
    a_strWaves(43) = "No Spectrum users were" & NL & "harmed during the making" & NL & "of this program"
    a_strWaves(44) = "This program was brought to you" & NL & "by the letter C" & NL & "and the number 64"
    a_strWaves(45) = "(c) 1999-2010 Jimbo"
    a_strWaves(46) = "This program is FREEWARE." & NL & "If you paid for it," & NL & "you may have been ripped off!"
    a_strWaves(47) = "BTW, if anyone wants to" & NL & "buy me anything, I'm after" & NL & "a HardSID Uno for my cab ;)"
    a_strWaves(48) = "LOL"
    a_strWaves(49) = "Sorry to everyone if their feature" & NL & "didn't make it in this" & NL & "version of GameBase"
    a_strWaves(50) = "There's always the next version!" & NL & "Mwaahahahaha!!"
    a_strWaves(51) = "Time 2 wrap..."
    a_strWaves(52) = "              "

    'store the maximum
    m_lngMaxWaves = c_lngMAXWAVES
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'read in and calculate the ascii value, the brightness-row, the destination
    'position on the starfield and the wait count for each char in each text wave
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'redimension the ACTUAL waves array to the specified number of waves
    ReDim ma_udtWaves(c_lngMAXWAVES)
    
    'store ascii values for CR and LF
    lngAscCR = Asc(vbCr)
    lngAscLF = Asc(vbLf)
    
    'loop though each wave
    For lngWaveIdx = 0 To c_lngMAXWAVES
        With ma_udtWaves(lngWaveIdx)
            
            'reset counters for each wave
            lngArrayCharIdx = 0
            lngWaveLen = Len(a_strWaves(lngWaveIdx))
            
            'calculate num rows, blit-startY and blit-height
            lngNumCRLF = GetNumOfCRandLF(a_strWaves(lngWaveIdx))
            lngLinesInWave = lngNumCRLF + 1
            .lngBlitHeight = (lngLinesInWave * m_lngFontPixelHeight) + ((lngLinesInWave - 1) * mc_lngLINESPACING)
            .lngBlitY = m_lngScreenCentreY - (.lngBlitHeight / 2)
            lngCharDestTop = .lngBlitY
            lngCurrentLine = 0
            ReDim a_lngPixelWidth(lngLinesInWave - 1)
            a_lngPixelWidth(lngCurrentLine) = 0
            ReDim a_lngCharsInLine(lngLinesInWave - 1)
            a_lngCharsInLine(lngCurrentLine) = 0

            'redim the wave's character array
            .lngMaxChars = lngWaveLen - (lngNumCRLF * 2) - 1
            ReDim .udtChars(.lngMaxChars)
            
            'loop through the wave's characters
            'calculating and storing the info
            For lngStringCharIdx = 1 To lngWaveLen
                
                'test the character
                lngAscChar = Asc(Mid$(a_strWaves(lngWaveIdx), lngStringCharIdx, 1))
                If lngAscChar = lngAscCR Then
                    ''''''''''''
                    'End of line
                    ''''''''''''
                    'move Y destination coordinate to next line
                    lngCharDestTop = lngCharDestTop + m_lngFontPixelHeight + mc_lngLINESPACING
                    'move to next line
                    lngCurrentLine = lngCurrentLine + 1
                    a_lngPixelWidth(lngCurrentLine) = 0
                    a_lngCharsInLine(lngCurrentLine) = 0
                    
                ElseIf lngAscChar = lngAscLF Then
                    'do nothing - just move to next char
                    
                Else
                    ''''''''''''''''''''''''
                    'Start or middle of line
                    ''''''''''''''''''''''''
                    'save props
                    With .udtChars(lngArrayCharIdx)
                        .lngAsciiValue = lngAscChar
                        .lngDestY = lngCharDestTop
                        .lngWidth = ma_udtFont(lngAscChar, 0).lngWidth
                    End With
                    'add to width of line
                    a_lngPixelWidth(lngCurrentLine) = a_lngPixelWidth(lngCurrentLine) + ma_udtFont(lngAscChar, 0).lngWidth
                    'add to num of chars in current line
                    a_lngCharsInLine(lngCurrentLine) = a_lngCharsInLine(lngCurrentLine) + 1
                    'move to next char space
                    lngArrayCharIdx = lngArrayCharIdx + 1
                End If
                      
            Next lngStringCharIdx
            
            'now we have the widths of all lines in the wave.
            'find out which is the widest line so that we can calculate
            'the blit-x and blit-width values for the wave...
            .lngBlitWidth = 0
            For lngLineIdx = 0 To (lngLinesInWave - 1)
                If a_lngPixelWidth(lngLineIdx) > .lngBlitWidth Then
                    .lngBlitWidth = a_lngPixelWidth(lngLineIdx)
                End If
            Next lngLineIdx
            .lngBlitX = m_lngScreenCentreX - (.lngBlitWidth / 2)
            
            'now we just need to set the x-position for each char in each line...
            lngCharIdx = 0
            For lngLineIdx = 0 To (lngLinesInWave - 1)
                'get starting x-position for line
                If a_lngCharsInLine(lngLineIdx) > 0 Then
                    lngLineStartX = m_lngScreenCentreX - (a_lngPixelWidth(lngLineIdx) / 2)
                    'set each char's x-position
                    For lngArrayCharIdx = 0 To (a_lngCharsInLine(lngLineIdx) - 1)
                        .udtChars(lngCharIdx).lngDestX = lngLineStartX
                        lngLineStartX = lngLineStartX + ma_udtFont(.udtChars(lngCharIdx).lngAsciiValue, 0).lngWidth
                        lngCharIdx = lngCharIdx + 1
                    Next lngArrayCharIdx
                End If
            Next lngLineIdx
            
        End With
    Next lngWaveIdx

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'exit the credits loop
    m_blnUnload = True
    DoEvents

End Sub

Private Sub lblEmail_Click()

        Call g_xobjFunc.ShellExecuteFile("mailto:gamebase.frontend@leesti.com", "", False)

End Sub

Private Sub lblFW_Click()

End Sub

Private Sub lblURL_Click()
    
    Call g_xobjFunc.ShellExecuteFile("http://www.bu22.com/", "", False)
    
End Sub

Private Sub tabAbout_Click(PreviousTab As Integer)

    With tabAbout
        If .Tab = 3 Then
            If PreviousTab <> 3 Then
                m_blnDoWaves = True
            End If
        Else
            If PreviousTab = 3 Then
                m_blnDoWaves = False
                Call ResetWaves
            End If
        End If
    End With

End Sub

