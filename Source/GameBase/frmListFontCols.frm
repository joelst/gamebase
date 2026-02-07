VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListFontCols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Font/Colours"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmListFontCols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtHidden 
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Change Font..."
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox picExample 
      BackColor       =   &H80000005&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Example Text"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label lblBack 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmListFontCols.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Background Colour"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblFore 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmListFontCols.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Text Colour"
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "frmListFontCols"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()

    On Error GoTo Err_Handler
    
    'show the font selector
    cDialog.Flags = cdlCFANSIOnly Or cdlCFForceFontExist Or cdlCFHelpButton Or cdlCFScreenFonts
    cDialog.FontBold = lblFont.Font.Bold
    cDialog.FontItalic = lblFont.Font.Italic
    cDialog.FontName = lblFont.Font.Name
    cDialog.FontSize = lblFont.Font.Size
    cDialog.ShowFont
    
    'change the example text font
    lblFont.Font.Bold = cDialog.FontBold
    lblFont.Font.Italic = cDialog.FontItalic
    lblFont.Font.Name = cDialog.FontName
    lblFont.Font.Size = cDialog.FontSize
    With lblFont
        .Top = (picExample.ScaleHeight / 2) - (.Height / 2)
        .Left = (picExample.ScaleWidth / 2) - (.Width / 2)
    End With
    
Err_Handler:
    'user pressed cancel

End Sub

Private Sub cmdOK_Click()
    
    Dim lngListView As Long     'loop index

    'set the font and colours in frmMain's listviews
    With frmMain
        For lngListView = 0 To .lvMain.Count - 1
            With .lvMain(lngListView)
                .BackColor = lblBack.BackColor
                .ForeColor = lblFore.BackColor
                .Font.Size = lblFont.Font.Size
                .Font.Bold = lblFont.Font.Bold
                .Font.Italic = lblFont.Font.Italic
                .Font.Name = lblFont.Font.Name
            End With
        Next lngListView
        With .lvLinks
            .BackColor = lblBack.BackColor
            .ForeColor = lblFore.BackColor
            .Font.Size = lblFont.Font.Size
            .Font.Bold = lblFont.Font.Bold
            .Font.Italic = lblFont.Font.Italic
            .Font.Name = lblFont.Font.Name
        End With
    End With
    
    'save to registry and close window
    With g_xobjRegistry
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "BackColour", lblBack.BackColor, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "ForeColour", lblFore.BackColor, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontBold", lblFont.Font.Bold, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontItalic", lblFont.Font.Italic, False
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontName", lblFont.Font.Name, True
        .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\GameList", "FontSize", lblFont.Font.Size, False
    End With
    
    Unload Me
    
End Sub


Private Sub cmdReset_Click()

    'set the font and colours
    With txtHidden
        lblFont.Font.Bold = .Font.Bold
        lblFont.Font.Italic = .Font.Italic
        lblFont.Font.Name = .Font.Name
        lblFont.Font.Size = .Font.Size
        lblFont.ForeColor = .ForeColor
        lblFore.BackColor = .ForeColor
        lblBack.BackColor = .BackColor
        picExample.BackColor = .BackColor
    End With
    
    'centre the example text
    With lblFont
        .Top = (picExample.ScaleHeight / 2) - (.Height / 2)
        .Left = (picExample.ScaleWidth / 2) - (.Width / 2)
    End With

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2535    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 4095     'The required scalewidth for the form
    
    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'read in the ini file text
    With g_xobjFunc
        Caption = .GetINIMessage("FontCols Window", "Caption", "List Font/Colours")
        lblFont = .GetINIMessage("FontCols Window", "ExampleText", "Example Text")
        cmdFont.Caption = .GetINIMessage("FontCols Window", "ChangeFontButton", "Change Font") & "..."
        cmdReset.Caption = .GetINIMessage("FontCols Window", "ResetButton", "Reset")
        lblFore.ToolTipText = .GetINIMessage("FontCols Window", "MouseOver_Foreground", "Change Text Colour")
        lblBack.ToolTipText = .GetINIMessage("FontCols Window", "MouseOver_Background", "Change Background Colour")
    End With

    'set the button text
    cmdOK.Caption = g_strButtonOKText
    cmdCancel.Caption = g_strButtonCancelText

    'set the font and colours
    With frmMain.lvMain(0)
        lblFont.Font.Bold = .Font.Bold
        lblFont.Font.Italic = .Font.Italic
        lblFont.Font.Name = .Font.Name
        lblFont.Font.Size = .Font.Size
        lblFont.ForeColor = .ForeColor
        lblFore.BackColor = .ForeColor
        lblBack.BackColor = .BackColor
        picExample.BackColor = .BackColor
    End With
    
    'centre the example text
    With lblFont
        .Top = (picExample.ScaleHeight / 2) - (.Height / 2)
        .Left = (picExample.ScaleWidth / 2) - (.Width / 2)
    End With

End Sub


Private Sub lblBack_Click()

    On Error GoTo Err_Handler
    
    'show the colour selector
    cDialog.Flags = cdlCCRGBInit
    cDialog.Color = picExample.BackColor
    cDialog.ShowColor
    
    'change the background colour
    picExample.BackColor = cDialog.Color
    lblBack.BackColor = cDialog.Color
    
Err_Handler:
    'user pressed cancel

End Sub

Private Sub lblFore_Click()

    On Error GoTo Err_Handler
    
    'show the colour selector
    cDialog.Flags = cdlCCRGBInit
    cDialog.Color = lblFont.ForeColor
    cDialog.ShowColor
    
    'change the text colour
    lblFont.ForeColor = cDialog.Color
    lblFore.BackColor = cDialog.Color
    
Err_Handler:
    'user pressed cancel

End Sub


