VERSION 5.00
Begin VB.Form frmManageEmus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Emulators"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "frmManageEmus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1095
   End
   Begin Gemus.ctlEmuSelector ctlEmuSelector 
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3096
   End
End
Attribute VB_Name = "frmManageEmus"
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

Private Gemus As clsGemus
Private m_xobjFunc As GBFuncs.clsFunctions
Private m_xobjRegistry As GBFuncs.clsRegistry
Private m_strMessFile As String
Private m_strAppPathWithBackSlash As String
Private m_strOldEmus() As String
Private m_lngNumOldEmus As Long
Private m_strEmuFile As String
Private m_strGBName As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim lngTotalEmulators As Long
    Dim lngIdx As Integer
    Dim lngEmuIdx As Long
    Dim strFiles() As String
    Dim lngNumFiles As Long
    Dim strScript As String
    Dim strName As String
    Dim strExe As String
    Dim blnUseShort As Boolean
    Dim lngFileIdx As Long
    Dim blnUseDefaultTypes As Boolean
    Dim strTypes As String
    Dim lngWait As Long
    Dim blnWritten As Boolean
    Dim blnFound As Boolean

    'save the details
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    blnWritten = True
    With ctlEmuSelector
        'write the emulator names and their info sections
        lngTotalEmulators = .ListCount
        For lngIdx = 0 To (lngTotalEmulators - 1)
            'get the emulator info
            Call .GetItem(lngIdx, strName, strExe, strScript, blnUseShort, strFiles(), lngNumFiles, blnUseDefaultTypes, strTypes, lngWait)
            'write to the emulator list
            blnWritten = blnWritten And m_xobjFunc.INI_WriteSetting("Emulators", Trim$(lngIdx + 1), strName, m_strEmuFile)
            'write the emulator info
            With m_xobjFunc
                blnWritten = blnWritten And .INI_WriteSetting(strName, "UseShortFilenames", IIf((blnUseShort = True), "1", "0"), m_strEmuFile)
                blnWritten = blnWritten And .INI_WriteSetting(strName, "UseDefaultTypes", IIf((blnUseDefaultTypes = True), "1", "0"), m_strEmuFile)
                blnWritten = blnWritten And .INI_WriteSetting(strName, "SupportedTypes", strTypes, m_strEmuFile)
                blnWritten = blnWritten And .INI_WriteSetting(strName, "EmulatorPathFile", strExe, m_strEmuFile)
                blnWritten = blnWritten And .INI_WriteSetting(strName, "ScriptFile", strScript, m_strEmuFile)
                blnWritten = blnWritten And .INI_WriteSetting(strName, "WaitFinish", Trim$(lngWait), m_strEmuFile)
                If lngNumFiles > 0 Then
                    For lngFileIdx = 0 To (lngNumFiles - 1)
                        blnWritten = blnWritten And .INI_WriteSetting(strName, Trim$(lngFileIdx + 1), strFiles(lngFileIdx), m_strEmuFile)
                    Next lngFileIdx
                End If
                blnWritten = blnWritten And .INI_DeleteNumberedValues(strName, lngNumFiles + 1, m_strEmuFile)
            End With
        Next lngIdx
        'delete extra unwanted emulators
        blnWritten = blnWritten And m_xobjFunc.INI_DeleteNumberedValues("Emulators", lngTotalEmulators + 1, m_strEmuFile)
        If blnWritten = True Then
            blnFound = False
            For lngIdx = 0 To (m_lngNumOldEmus - 1)
                For lngEmuIdx = 0 To (.ListCount - 1)
                    If UCase$(m_strOldEmus(lngIdx)) = UCase$(.List(lngEmuIdx)) Then
                        blnFound = True
                        Exit For
                    End If
                Next lngEmuIdx
                If blnFound = False Then
                    m_xobjFunc.INI_DeleteSection m_strOldEmus(lngIdx), m_strEmuFile
                End If
            Next lngIdx
        Else
            MsgBox "There was an error saving the Emulator info to " & m_strEmuFile, vbExclamation, "GEMUS"
        End If
    End With

    g_blnEmusManaged = True
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub Form_Load()

    Dim lngEmuIdx As Long                       'loop index
    Dim strFiles() As String
    Const DesignScaleHeight As Single = 2415    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 3630     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set references
    Set Gemus = New clsGemus
    Set m_xobjFunc = New GBFuncs.clsFunctions
    Set m_xobjRegistry = New GBFuncs.clsRegistry

    'init vars
    m_strAppPathWithBackSlash = App.Path
    If Right$(m_strAppPathWithBackSlash, 1) <> "\" Then m_strAppPathWithBackSlash = m_strAppPathWithBackSlash & "\"
    m_strMessFile = m_strAppPathWithBackSlash & "GemusText.ini"
    m_strGBName = g_strManageEmuGBName
    m_strEmuFile = g_strManageEmuFile
    Call Gemus.Init(g_strManageEmuGBName)

    'buttons
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", m_strMessFile)
    cmdCancel.Caption = GetINIMessageFromFile("General Buttons", "Cancel", "&Cancel", m_strMessFile)
    Caption = GetINIMessageFromFile("Manage Emulators Window", "Caption", "Manage Emulators", m_strMessFile)

    'emulators
    ctlEmuSelector.SetGameBaseNameInfo m_strGBName, m_strEmuFile
    With Gemus
        m_lngNumOldEmus = .TotalEmulators
        ReDim m_strOldEmus(m_lngNumOldEmus)
        For lngEmuIdx = 0 To (.TotalEmulators - 1)
            Call .EmulatorFiles(lngEmuIdx, strFiles())
            ctlEmuSelector.AddItem .EmulatorName(lngEmuIdx), .EmulatorPathExe(lngEmuIdx), .EmulatorScript(lngEmuIdx), .EmulatorUseShort(lngEmuIdx), strFiles(), .EmulatorTotalFiles(lngEmuIdx), .EmulatorUseDefaultTypes(lngEmuIdx), .EmulatorTypes(lngEmuIdx), .EmulatorWaits(lngEmuIdx)
            m_strOldEmus(lngEmuIdx) = .EmulatorName(lngEmuIdx)
        Next lngEmuIdx
    End With
    g_blnEmusManaged = False

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'free class references
    Set Gemus = Nothing
    Set m_xobjFunc = Nothing
    Set m_xobjRegistry = Nothing

End Sub


