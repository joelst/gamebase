VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tips"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picWindow 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.Image imgTip 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   120
         Picture         =   "frmTip.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Title goes here"
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
         Left            =   720
         TabIndex        =   4
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   1620
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Text goes here"
         Height          =   2055
         Left            =   720
         TabIndex        =   5
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   3975
      End
      Begin VB.Line lineTip 
         X1              =   720
         X2              =   4680
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CheckBox chkThisTip 
      Caption         =   "&Don't show this Tip again"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CheckBox chkAllTips 
      Caption         =   "&Never show Tips"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3300
      Width           =   3735
   End
End
Attribute VB_Name = "frmTip"
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

Private m_blnClickByCode As Boolean



Private Sub chkAllTips_Click()

    If m_blnClickByCode = False Then
        'this means that the other checkbox should be checked too
        m_blnClickByCode = True
        chkThisTip.Value = vbChecked
        m_blnClickByCode = False
    End If
    
End Sub


Private Sub chkThisTip_Click()

    If m_blnClickByCode = False Then
        'need to uncheck the other checkbox if unchecking this one
        If chkThisTip.Value = vbUnchecked Then
            m_blnClickByCode = True
            chkAllTips.Value = vbUnchecked
            m_blnClickByCode = False
        End If
    End If
    
End Sub


Private Sub cmdClose_Click()

    '----------------------------------------------------------
    'if we have to save the "view tip(s) again" info then do so
    '----------------------------------------------------------

    If chkAllTips.Value = vbChecked Then
       'amend ALL tips not to show again
       g_udtSelectedTip.blnShowIt = False
       Call ResetTips(False)
        
    ElseIf chkThisTip.Value = vbChecked Then
        'amend THIS tip not to show again
        g_udtSelectedTip.blnShowIt = False
        Call g_xobjRegistry.SaveSetting("HKEY_CURRENT_USER", "Software\GB64\GameBase\Tips", g_udtSelectedTip.strINIAndRegKey, g_udtSelectedTip.blnShowIt, False)
        
    End If

    Unload Me

End Sub


Private Sub Form_Load()

    Const DesignScaleHeight As Single = 3645    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 5175     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)
    
    'get language specific text
    With g_xobjFunc
        Caption = gc_strAPPNAME & " " & .GetINIMessage("Tips", "Tips", "Tips")
        cmdClose.Caption = g_strButtonCloseText
        chkThisTip.Caption = .GetINIMessage("Tips", "ThisTip", "&Don't show this Tip again")
        chkAllTips.Caption = .GetINIMessage("Tips", "AllTips", "&Never show Tips")
    End With
    
    'Set the control data
    lblTitle = g_strSelectedTipTitle
    lblText = g_strSelectedTipText
    
    'Resize the controls in the container picture box
    Call ResizeControls
    
    'user clicks checkboxes before we do it in code
    m_blnClickByCode = False
    
End Sub

Private Sub ResizeControls()
    '-------------------------------------------------------
    'Resizes the controls so that they all fit nice and snug
    'within the picture box
    '
    'Assumptions:   lblTitle must have been filled with the
    '               requested tip title so that the height
    '               is set by it's AutoSize property
    '-------------------------------------------------------

    Const c_lngTWIPSSPACING As Long = 120   'The spacing width/height in twips
    Dim sngControlsLeft As Single           'Left position of controls
    Dim sngControlsWidth As Single          'Width of controls

    'Place the Tip Image
    imgTip.Top = c_lngTWIPSSPACING
    imgTip.Left = c_lngTWIPSSPACING

    'Calculate the Left and Width properties for the controls
    sngControlsLeft = imgTip.Left + imgTip.Width + c_lngTWIPSSPACING
    sngControlsWidth = picWindow.ScaleWidth - sngControlsLeft - c_lngTWIPSSPACING

    'Set the Left and Width properties for the controls
    lblTitle.Left = sngControlsLeft
    lblTitle.Width = sngControlsWidth
    lineTip.X1 = sngControlsLeft
    lineTip.X2 = sngControlsLeft + sngControlsWidth
    lblText.Left = sngControlsLeft
    lblText.Width = sngControlsWidth

    'Set the Top property for the controls
    lblTitle.Top = c_lngTWIPSSPACING
    lineTip.Y1 = c_lngTWIPSSPACING + lblTitle.Height + c_lngTWIPSSPACING
    lineTip.Y2 = lineTip.Y1
    lblText.Top = lineTip.Y1 + c_lngTWIPSSPACING

    'Set the Height property for the lblText control
    lblText.Height = picWindow.ScaleHeight - lblText.Top - c_lngTWIPSSPACING

End Sub

