VERSION 5.00
Begin VB.Form frmRezipWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Running ..."
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmRezipWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Waiting"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for the emulator to close..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmRezipWait"
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
'     GEMUS - Gamebase EMUlator Script
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

Public m_dblProcessID As Double
Public m_lngHandle As Long
Public m_blnStopped As Boolean
Private m_blnActivated As Boolean
Private Sub cmdStop_Click()

    m_blnStopped = True
    cmdStop.Enabled = False
    DoEvents

End Sub

Private Sub Form_Activate()

    Dim blnClosed As Boolean
    Dim lngHandle As Long
    Dim lngRetVal As Long
    Dim lngExitCode As Long

    If m_blnActivated = True Then
        Exit Sub
    End If

    m_blnActivated = True
    Screen.MousePointer = vbHourglass
    DoEvents
    
    'get the process handle for the process ID
    If m_lngHandle = 0 Then
        lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, m_dblProcessID)
    Else
        lngHandle = m_lngHandle
    End If
    
    'loop until either the process finished or the user clicked the button
    Do While m_blnStopped = False
        Call GetExitCodeProcess(lngHandle, lngExitCode)
        DoEvents
        If lngExitCode <> STATUS_PENDING Then
            Exit Do
        End If
    Loop
    
    'all done
    Call CloseHandle(lngHandle)
    Screen.MousePointer = vbDefault
    Hide

End Sub

Private Sub Form_Load()

    Width = 4815 + (Width - ScaleWidth)
    Height = 975 + (Height - ScaleHeight)
    m_blnActivated = False
    m_blnStopped = False
    m_dblProcessID = 0
    m_lngHandle = 0
    cmdStop.Enabled = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        m_blnStopped = True
        cmdStop.Enabled = False
        DoEvents
    End If

End Sub


