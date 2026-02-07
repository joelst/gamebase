VERSION 5.00
Begin VB.Form frmBigScreenshot 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Big Screenshot"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBigScreenshot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Image imgSS 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmBigScreenshot"
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

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Load()

    Dim sngTop As Single
    Dim sngLeft As Single

    g_blnLoadedfrmBigScreenshot = True
    
    'set the starting position of the window
    With Screen
        If .Width / .TwipsPerPixelX < 800 Then
            Width = g_lngFullSizePictureWidthPixels * .TwipsPerPixelX
            Height = g_lngFullSizePictureHeightPixels * .TwipsPerPixelY
        Else
            Width = g_lngZoomedPictureWidthPixels * .TwipsPerPixelX
            Height = g_lngZoomedPictureHeightPixels * .TwipsPerPixelY
        End If
    End With

    With g_xobjRegistry
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "BSSTop", sngTop, False, gc_lngINVALIDWINDOWPOS
        .GetSetting_HKCU_Fallback_HKLM "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "BSSLeft", sngLeft, False, gc_lngINVALIDWINDOWPOS
    End With
    With Screen
        If sngTop < 0 Or sngTop > .Height Or sngLeft < 0 Or sngLeft > .Width Then
            sngTop = (.Height / 2) - (Height / 2)
            sngLeft = (.Width / 2) - (Width / 2)
        End If
    End With
    Top = sngTop
    Left = sngLeft

End Sub

Public Sub ShowPicture(picBox As PictureBox, strCaption_in As String)

    Dim sngMaxTwipWidth As Single
    Dim sngMaxTwipHeight As Single
    Dim sngExtraWidth As Single
    Dim sngExtraHeight As Single
    
    Dim sngDesiredTwipWidth As Single
    Dim sngDesiredTwipHeight As Single

    On Error GoTo Err_Handler

    'get maximum size for zoomed picture
    sngExtraWidth = Width - ScaleWidth
    sngExtraHeight = Height - ScaleHeight
    With Screen
        sngMaxTwipWidth = .Width - sngExtraWidth
        sngMaxTwipHeight = .Height - sngExtraHeight
        sngDesiredTwipWidth = g_lngZoomedPictureWidthPixels * .TwipsPerPixelX
        sngDesiredTwipHeight = g_lngZoomedPictureHeightPixels * .TwipsPerPixelY
    End With
    If sngDesiredTwipWidth > sngMaxTwipWidth Then sngDesiredTwipWidth = sngMaxTwipWidth
    If sngDesiredTwipHeight > sngMaxTwipHeight Then sngDesiredTwipHeight = sngMaxTwipHeight

    With picBox
        If .Width > .Height Then
            imgSS.Width = sngDesiredTwipWidth
            imgSS.Height = sngDesiredTwipWidth * (.Height / .Width)
        Else
            imgSS.Height = sngDesiredTwipHeight
            imgSS.Width = sngDesiredTwipHeight * (.Width / .Height)
        End If
    End With

JustDisplay:

    On Error Resume Next

    'show the picture
    With imgSS
        .Picture = picBox.Image
        Width = .Width + sngExtraWidth
        Height = .Height + sngExtraHeight
    End With
    Caption = strCaption_in
    Visible = True
    If g_blnSetPicZoomForeground = True Then
        Call SetForegroundWindow(hwnd)
    End If

    Exit Sub

Err_Handler:
    'probably an overflow (picture too large)
    MsgBox Err.Description & " (" & Err.Number & "," & Err.LastDllError & ")", , gc_strAPPNAME
    imgSS.Width = picBox.Width
    imgSS.Height = picBox.Height
    Resume JustDisplay

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    'save the window position
    If WindowState = vbNormal Then
        With g_xobjRegistry
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "BSSTop", Top, False
            .SaveSetting "HKEY_CURRENT_USER", "Software\GB64\GameBase\GameBase\" & g_strGameBaseName & "\Window", "BSSLeft", Left, False
        End With
    End If
    g_blnLoadedfrmBigScreenshot = False

End Sub


Private Sub imgSS_DblClick()
    Unload Me
End Sub


