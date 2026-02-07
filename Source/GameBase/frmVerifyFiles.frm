VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerifyFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Available Files"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmVerifyFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraVerify 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox chkCRC32 
         Caption         =   "C&alculate CRC-32 Checksums"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.CheckBox chkExtras 
         Caption         =   "&Extras"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox chkScrnShots 
         Caption         =   "S&creenshots"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkSids 
         Caption         =   "&Sids"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkGames 
         Caption         =   "&Games"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblExtras 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Extras not found: 0/0"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblScrnShots 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Screenshots not found: 0/0"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblSids 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sids not found: 0/0"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblGames 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Games not found: 0/0"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frmVerifyFiles"
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

Private m_blnBusy As Boolean
Private m_blnDone As Boolean
Private m_blnCancelled As Boolean
Private m_strCleanup As String
Private m_strGames As String
Private m_intFileNum As Integer

Private m_lngTotalGames As Long
Private m_lngMissingGames As Long
Private m_lngTotalSids As Long
Private m_lngMissingSids As Long
Private m_lngTotalScrnShots As Long
Private m_lngMissingScrnShots As Long
Private m_lngTotalExtras As Long
Private m_lngMissingExtras As Long

Private m_CRC As clsCRC
Private Sub SetStartEnabled()

    cmdStart.Enabled = ((chkGames.Value = vbChecked) Or (chkSids.Value = vbChecked) Or (chkScrnShots.Value = vbChecked) Or (chkExtras.Value = vbChecked))

End Sub

Private Sub VerifyExtras()

    Dim lngEXidx As Long
    Dim mbRes As VbMsgBoxResult
    Dim fldName As Field
    Dim fldFilename As Field
    Dim fldID As Field
    Dim fldExtraName As Field
    Dim fldEXID As Field
    Dim fldEA As Field
    Dim fldType As Field
    Dim lngMaxEX As Long
    Dim blnFound As Boolean
    Dim strFile As String
    Dim blnBlankFilename As Boolean
    Dim strPaths As String

    On Error GoTo Failed

    If m_blnCancelled = False Then
        If (g_lngTotalExtrasPaths > 0) Then
            
            'display prompt for media
            lngMaxEX = g_lngTotalExtrasPaths - 1
            mbRes = vbOK
            For lngEXidx = 0 To lngMaxEX
                If ga_udtRegPathExtras(lngEXidx).ePathType <> e_PathLocation_FixedHD Then
                    mbRes = MsgBox(g_xobjFunc.GetINIMessage("Verify Available Files", "ExtrasDevicePrompt", "Please make sure that any CDs or removable drives housing your extras paths are inserted before continuing."), vbInformation Or vbOKCancel)
                    Exit For
                End If
            Next lngEXidx

            'verify extras
            If mbRes = vbOK Then
                strFile = ""
                m_lngMissingExtras = 0
                m_lngTotalExtras = 0
                'check extras
                lblStatus = g_xobjFunc.GetINIMessage("Verify Available Files", "VerifyingExtras", "Verifying extras") & "..."
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT Extras.EX_Id, Extras.GA_Id, Extras.Name AS ExtraName, Extras.EA, Extras.Path, Extras.Type, Games.Name FROM Extras INNER JOIN Games ON Extras.GA_Id = Games.GA_Id ORDER BY Games.Name", dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        Set fldName = !Name
                        Set fldFilename = !Path
                        Set fldID = !GA_Id
                        Set fldExtraName = !ExtraName
                        Set fldEXID = !EX_Id
                        Set fldEA = !EA
                        Set fldType = !Type
                        .MoveLast
                        .MoveFirst
                        pbStatus.Min = 0
                        pbStatus.Max = .RecordCount
                        .MoveFirst
                        Do Until .EOF Or m_blnCancelled
                            'find the file
                            blnFound = False
                            blnBlankFilename = False
                            If fldType < 3 Then
                                'file within extras path, not url
                                If fldFilename <> "" Then
                                    m_lngTotalExtras = m_lngTotalExtras + 1
                                    For lngEXidx = 0 To lngMaxEX
                                        If g_xobjFunc.FileExists(ga_udtRegPathExtras(lngEXidx).strWithBackSlash & fldFilename, False, "") Then
                                            blnFound = True
                                            Exit For
                                        End If
                                    Next lngEXidx
                                Else
                                    blnBlankFilename = True
                                    .Edit
                                    fldEA = False
                                    .Update
                                End If
                            Else
                                'url
                                blnFound = True
                            End If
                            'update the record
                            If Not blnFound Then
                                If Not blnBlankFilename Then
                                    strFile = strFile & fldEXID & vbTab & fldExtraName & vbTab & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                                    m_lngMissingExtras = m_lngMissingExtras + 1
                                End If
                                .Edit
                                fldEA = False
                                .Update
                            Else
                                .Edit
                                fldEA = True
                                .Update
                            End If
                            'update status and move to next record
                            pbStatus.Value = .AbsolutePosition
                            .MoveNext
                            DoEvents
                        Loop
                        pbStatus.Value = pbStatus.Max
                        DoEvents
                        Set fldName = Nothing
                        Set fldFilename = Nothing
                        Set fldID = Nothing
                        Set fldExtraName = Nothing
                        Set fldEXID = Nothing
                        Set fldEA = Nothing
                        Set fldType = Nothing
                    End If
                End With
                g_rsGames.Close
                Set g_rsGames = Nothing

                'write to file
                If m_lngMissingExtras > 0 Then
                    strPaths = ""
                    For lngEXidx = 0 To lngMaxEX
                        strPaths = strPaths & ga_udtRegPathExtras(lngEXidx).strWithBackSlash & vbCrLf
                    Next lngEXidx
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "ExtrasNotFound", "Extras not found") & ": " & m_lngMissingExtras & "/" & m_lngTotalExtras & vbCrLf & strPaths & vbCrLf & strFile & vbCrLf & vbCrLf
                Else
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "ExtrasNotFound", "Extras not found") & ": " & m_lngMissingExtras & "/" & m_lngTotalExtras & vbCrLf & vbCrLf & vbCrLf
                End If
            End If

        End If
    End If
    
    Exit Sub

Failed:
    MsgBox Err.Description & " (Err " & Err.Number & "," & Err.LastDllError & ")"
    m_blnCancelled = True

End Sub

Private Sub VerifyGames()
    
    Dim lngGPIdx As Long
    Dim mbRes As VbMsgBoxResult
    Dim fldName As Field
    Dim fldFilename As Field
    Dim fldID As Field
    Dim fldCCode As Field
    Dim fldFA As Field
    Dim fldSA As Field
    Dim fldCRC As Field
    Dim lngMaxGP As Long
    Dim blnFound As Boolean
    Dim strFile As String
    Dim blnBlankFilename As Boolean
    Dim strPaths As String
    Dim blnCheckCRC As Boolean
    Dim strFullFile As String
    Dim strCRC As String

    On Error GoTo Failed

    Set m_CRC = New clsCRC
    blnCheckCRC = False '(chkCRC32.Value = vbChecked)

    If m_blnCancelled = False Then
        If (g_lngTotalGamePaths > 0) And (g_blnValidGamePathExists = True) Then
            
            'display prompt for media
            lngMaxGP = g_lngTotalGamePaths - 1
            mbRes = vbOK
            For lngGPIdx = 0 To lngMaxGP
                If ga_udtRegPathGame(lngGPIdx).ePathType <> e_PathLocation_FixedHD Then
                    mbRes = MsgBox(g_xobjFunc.GetINIMessage("Verify Available Files", "GameDevicePrompt", "Please make sure that any CDs or removable drives housing your game paths are inserted before continuing."), vbInformation Or vbOKCancel)
                    Exit For
                End If
            Next lngGPIdx

            'verify games
            If mbRes = vbOK Then
                strFile = ""
                m_lngMissingGames = 0
                m_lngTotalGames = 0
                'check games
                lblStatus = g_xobjFunc.GetINIMessage("Verify Available Files", "VerifyingGames", "Verifying games") & "..."
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Name, Filename, CCode, FA, SA, CRC FROM Games ORDER BY Name", dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        Set fldName = !Name
                        Set fldFilename = !Filename
                        Set fldID = !GA_Id
                        Set fldCCode = !CCode
                        Set fldFA = !FA
                        Set fldSA = !SA
                        Set fldCRC = !CRC
                        .MoveLast
                        .MoveFirst
                        pbStatus.Min = 0
                        pbStatus.Max = .RecordCount
                        .MoveFirst
                        Do Until .EOF Or m_blnCancelled
                            'find the file
                            blnFound = False
                            blnBlankFilename = False
                            strCRC = ""
                            If fldFilename <> "" Then
                                m_lngTotalGames = m_lngTotalGames + 1
                                For lngGPIdx = 0 To lngMaxGP
                                    strFullFile = ga_udtRegPathGame(lngGPIdx).strWithBackSlash & fldFilename
                                    If g_xobjFunc.FileExists(strFullFile, False, "") Then
                                        If blnCheckCRC Then
                                            strCRC = Hex(m_CRC.CalculateFile(strFullFile))
                                        End If
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngGPIdx
                            Else
                                blnBlankFilename = True
                            End If
                            'update the record
                            If blnFound Then
                                .Edit
                                fldFA = True
                                If fldSA = True Then
                                    fldCCode = 1 'both exist
                                Else
                                    fldCCode = 2 'only game exists
                                End If
                                fldCRC = strCRC
                                .Update
                            Else
                                .Edit
                                fldFA = False
                                If fldSA = True Then
                                    fldCCode = 3 'only sid exists
                                Else
                                    fldCCode = 4 'none exist
                                End If
                                fldCRC = strCRC
                                .Update
                                If Not blnBlankFilename Then
                                    strFile = strFile & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                                    m_lngMissingGames = m_lngMissingGames + 1
                                End If
                            End If
                            'update status and move to next record
                            pbStatus.Value = .AbsolutePosition
                            .MoveNext
                            DoEvents
                        Loop
                        pbStatus.Value = pbStatus.Max
                        DoEvents
                        Set fldName = Nothing
                        Set fldFilename = Nothing
                        Set fldID = Nothing
                        Set fldCCode = Nothing
                        Set fldFA = Nothing
                        Set fldSA = Nothing
                        Set fldCRC = Nothing
                    End If
                End With
                g_rsGames.Close
                Set g_rsGames = Nothing

                'write to file
                If m_lngMissingGames > 0 Then
                    strPaths = ""
                    For lngGPIdx = 0 To lngMaxGP
                        strPaths = strPaths & ga_udtRegPathGame(lngGPIdx).strWithBackSlash & vbCrLf
                    Next lngGPIdx
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "GamesNotFound", "Games not found") & ": " & m_lngMissingGames & "/" & m_lngTotalGames & vbCrLf & strPaths & vbCrLf & strFile & vbCrLf & vbCrLf
                Else
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "GamesNotFound", "Games not found") & ": " & m_lngMissingGames & "/" & m_lngTotalGames & vbCrLf & vbCrLf & vbCrLf
                End If
            End If

        End If
    End If
    
    Set m_CRC = Nothing
    Exit Sub

Failed:
    MsgBox Err.Description & " (Err " & Err.Number & "," & Err.LastDllError & ")"
    m_blnCancelled = True
    Set m_CRC = Nothing

End Sub
Private Sub VerifyScrnShots()
    
    Dim lngSSIdx As Long
    Dim mbRes As VbMsgBoxResult
    Dim fldName As Field
    Dim fldFilename As Field
    Dim fldID As Field
    Dim lngMaxSS As Long
    Dim blnFound As Boolean
    Dim strFile As String
    Dim blnBlankFilename As Boolean
    Dim strPaths As String

    On Error GoTo Failed

    If m_blnCancelled = False Then
        If (g_lngTotalScrnShotPaths > 0) Then
            
            'display prompt for media
            lngMaxSS = g_lngTotalScrnShotPaths - 1
            mbRes = vbOK
            For lngSSIdx = 0 To lngMaxSS
                If ga_udtRegPathScrnShot(lngSSIdx).ePathType <> e_PathLocation_FixedHD Then
                    mbRes = MsgBox(g_xobjFunc.GetINIMessage("Verify Available Files", "PictureDevicePrompt", "Please make sure that any CDs or removable drives housing your picture paths are inserted before continuing."), vbInformation Or vbOKCancel)
                    Exit For
                End If
            Next lngSSIdx

            'verify screenshots
            If mbRes = vbOK Then
                strFile = ""
                m_lngMissingScrnShots = 0
                m_lngTotalScrnShots = 0
                'check screenshots
                lblStatus = g_xobjFunc.GetINIMessage("Verify Available Files", "VerifyingPictures", "Verifying pictures") & "..."
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Name, ScrnshotFilename FROM Games ORDER BY Name", dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        Set fldName = !Name
                        Set fldFilename = !ScrnShotFilename
                        Set fldID = !GA_Id
                        .MoveLast
                        .MoveFirst
                        pbStatus.Min = 0
                        pbStatus.Max = .RecordCount
                        .MoveFirst
                        Do Until .EOF Or m_blnCancelled
                            'find the file
                            blnFound = False
                            blnBlankFilename = False
                            If fldFilename <> "" Then
                                m_lngTotalScrnShots = m_lngTotalScrnShots + 1
                                For lngSSIdx = 0 To lngMaxSS
                                    If g_xobjFunc.FileExists(ga_udtRegPathScrnShot(lngSSIdx).strWithBackSlash & fldFilename, False, "") Then
                                        blnFound = True
                                        Exit For
                                    End If
                                Next lngSSIdx
                            Else
                                blnBlankFilename = True
                            End If
                            'update the record
                            If Not blnFound Then
                                If Not blnBlankFilename Then
                                    strFile = strFile & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                                    m_lngMissingScrnShots = m_lngMissingScrnShots + 1
                                End If
                            End If
                            'update status and move to next record
                            pbStatus.Value = .AbsolutePosition
                            .MoveNext
                            DoEvents
                        Loop
                        pbStatus.Value = pbStatus.Max
                        DoEvents
                        Set fldName = Nothing
                        Set fldFilename = Nothing
                        Set fldID = Nothing
                    End If
                End With
                g_rsGames.Close
                Set g_rsGames = Nothing

                'write to file
                If m_lngMissingScrnShots > 0 Then
                    strPaths = ""
                    For lngSSIdx = 0 To lngMaxSS
                        strPaths = strPaths & ga_udtRegPathScrnShot(lngSSIdx).strWithBackSlash & vbCrLf
                    Next lngSSIdx
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "PicturesNotFound", "Pictures not found") & ": " & m_lngMissingScrnShots & "/" & m_lngTotalScrnShots & vbCrLf & strPaths & vbCrLf & strFile & vbCrLf & vbCrLf
                Else
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "PicturesNotFound", "Pictures not found") & ": " & m_lngMissingScrnShots & "/" & m_lngTotalScrnShots & vbCrLf & vbCrLf & vbCrLf
                End If
            End If

        End If
    End If
    
    Exit Sub

Failed:
    MsgBox Err.Description & " (Err " & Err.Number & "," & Err.LastDllError & ")"
    m_blnCancelled = True

End Sub
Private Sub VerifySids()

    Dim mbRes As VbMsgBoxResult
    Dim fldName As Field
    Dim fldFilename As Field
    Dim fldID As Field
    Dim fldCCode As Field
    Dim fldFA As Field
    Dim fldSA As Field
    Dim blnFound As Boolean
    Dim strFile As String
    Dim strFile2 As String
    Dim blnBlankFilename As Boolean

    On Error GoTo Failed

    If m_blnCancelled = False Then
        If g_udtRegPathSid.ePathType <> e_PathLocation_Invalid Then
            
            'display prompt for media
            mbRes = vbOK
            If g_udtRegPathSid.ePathType <> e_PathLocation_FixedHD Then
                mbRes = MsgBox(g_xobjFunc.GetINIMessage("Verify Available Files", "MusicDevicePrompt", "Please make sure that the CD or removable drive housing your music files is inserted before continuing."), vbInformation Or vbOKCancel)
            End If

            'verify sids
            If mbRes = vbOK Then
                strFile = ""
                m_lngMissingSids = 0
                m_lngTotalSids = 0
                'check games table
                lblStatus = g_xobjFunc.GetINIMessage("Verify Available Files", "VerifyingTunesInGames", "Verifying tunes in Games table") & "..."
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Name, SidFilename, CCode, FA, SA FROM Games ORDER BY Name", dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        Set fldName = !Name
                        Set fldFilename = !SidFilename
                        Set fldID = !GA_Id
                        Set fldCCode = !CCode
                        Set fldFA = !FA
                        Set fldSA = !SA
                        .MoveLast
                        .MoveFirst
                        pbStatus.Min = 0
                        pbStatus.Max = .RecordCount
                        .MoveFirst
                        Do Until .EOF Or m_blnCancelled
                            'find the file
                            blnFound = False
                            blnBlankFilename = False
                            If fldFilename <> "" Then
                                m_lngTotalSids = m_lngTotalSids + 1
                                If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & fldFilename, False, "") Then
                                    blnFound = True
                                End If
                            Else
                                blnBlankFilename = True
                            End If
                            'update the record
                            If blnFound Then
                                .Edit
                                fldSA = True
                                If fldFA = True Then
                                    fldCCode = 1 'both exist
                                Else
                                    fldCCode = 3 'only sid exists
                                End If
                                .Update
                            Else
                                .Edit
                                fldSA = False
                                If fldFA = True Then
                                    fldCCode = 2 'only game exists
                                Else
                                    fldCCode = 4 'none exist
                                End If
                                .Update
                                If Not blnBlankFilename Then
                                    strFile = strFile & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                                    m_lngMissingSids = m_lngMissingSids + 1
                                End If
                            End If
                            'update status and move to next record
                            pbStatus.Value = .AbsolutePosition
                            .MoveNext
                            DoEvents
                        Loop
                        pbStatus.Value = pbStatus.Max
                        DoEvents
                        Set fldName = Nothing
                        Set fldFilename = Nothing
                        Set fldID = Nothing
                        Set fldCCode = Nothing
                        Set fldFA = Nothing
                        Set fldSA = Nothing
                    End If
                End With
                g_rsGames.Close
                Set g_rsGames = Nothing
                If strFile <> "" Then
                    strFile = g_xobjFunc.GetINIMessage("Verify Available Files", "GamesTable", "Games table") & "..." & vbCrLf & vbCrLf & strFile
                End If

                'check music table
                strFile2 = ""
                lblStatus = g_xobjFunc.GetINIMessage("Verify Available Files", "VerifyingTunesInMusic", "Verifying tunes in Music table") & "..."
                Set g_rsGames = g_dbGamebase.OpenRecordset("SELECT GA_Id, Name, Filename, SA FROM Music ORDER BY Name", dbOpenDynaset)
                With g_rsGames
                    If Not .EOF Then
                        Set fldName = !Name
                        Set fldFilename = !Filename
                        Set fldID = !GA_Id
                        Set fldSA = !SA
                        .MoveLast
                        .MoveFirst
                        pbStatus.Min = 0
                        pbStatus.Max = .RecordCount
                        m_lngTotalSids = m_lngTotalSids + .RecordCount
                        .MoveFirst
                        Do Until .EOF Or m_blnCancelled
                            'find the file
                            blnFound = False
                            blnBlankFilename = False
                            If fldFilename <> "" Then
                                If g_xobjFunc.FileExists(g_udtRegPathSid.strWithBackSlash & fldFilename, False, "") Then
                                    blnFound = True
                                End If
                            Else
                                blnBlankFilename = True
                            End If
                            'update the record
                            If blnFound Then
                                .Edit
                                fldSA = True
                                .Update
                            Else
                                .Edit
                                fldSA = False
                                .Update
                                If Not blnBlankFilename Then
                                    strFile2 = strFile2 & fldFilename & vbTab & fldID & vbTab & fldName & vbCrLf
                                End If
                                m_lngMissingSids = m_lngMissingSids + 1
                            End If
                            'update status and move to next record
                            pbStatus.Value = .AbsolutePosition
                            .MoveNext
                            DoEvents
                        Loop
                        pbStatus.Value = pbStatus.Max
                        DoEvents
                        Set fldName = Nothing
                        Set fldFilename = Nothing
                        Set fldID = Nothing
                        Set fldSA = Nothing
                    End If
                End With
                g_rsGames.Close
                Set g_rsGames = Nothing
                If strFile2 <> "" Then
                    strFile2 = g_xobjFunc.GetINIMessage("Verify Available Files", "MusicTable", "Music table") & "..." & vbCrLf & vbCrLf & strFile2
                End If
                
                'write to file
                If m_lngMissingSids > 0 Then
                    If strFile <> "" And strFile2 <> "" Then
                        Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "TunesNotFound", "Tunes not found") & ": " & m_lngMissingSids & "/" & m_lngTotalSids & vbCrLf & g_udtRegPathSid.strWithBackSlash & vbCrLf & vbCrLf & strFile & vbCrLf & strFile2 & vbCrLf & vbCrLf
                    ElseIf strFile <> "" Then
                        Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "TunesNotFound", "Tunes not found") & ": " & m_lngMissingSids & "/" & m_lngTotalSids & vbCrLf & g_udtRegPathSid.strWithBackSlash & vbCrLf & vbCrLf & strFile & vbCrLf & vbCrLf
                    Else
                        Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "TunesNotFound", "Tunes not found") & ": " & m_lngMissingSids & "/" & m_lngTotalSids & vbCrLf & g_udtRegPathSid.strWithBackSlash & vbCrLf & vbCrLf & strFile2 & vbCrLf & vbCrLf
                    End If
                Else
                    Print #m_intFileNum, g_xobjFunc.GetINIMessage("Verify Available Files", "TunesNotFound", "Tunes not found") & ": " & m_lngMissingSids & "/" & m_lngTotalSids & vbCrLf & vbCrLf & vbCrLf
                End If
            End If

        End If
    End If
    
    Exit Sub

Failed:
    MsgBox Err.Description & "  (Err " & Err.Number & "," & Err.LastDllError & ")"
    m_blnCancelled = True

End Sub

Private Sub chkCRC32_Click()

'    If chkCRC32.Value = vbChecked Then
'        If chkGames.Value = vbUnchecked Then
'            chkGames.Value = vbChecked
'        End If
'    End If

End Sub

Private Sub chkExtras_Click()
    Call SetStartEnabled
End Sub

Private Sub chkGames_Click()
    If chkGames.Value = vbUnchecked Then
        'chkCRC32.Value = vbUnchecked
    End If
    Call SetStartEnabled
End Sub


Private Sub chkScrnShots_Click()
    Call SetStartEnabled
End Sub


Private Sub chkSids_Click()
    Call SetStartEnabled
End Sub


Private Sub cmdCancel_Click()
    
    If m_blnBusy = False Then
        Unload Me
    Else
        cmdCancel.Enabled = False
        m_blnCancelled = True
    End If
    
End Sub

Private Sub cmdStart_Click()

    On Error GoTo ErrHandler

    If m_blnDone = False Then
        'prepare the window
        g_blnVerifyDone = True
        cmdStart.Enabled = False
        cmdStart.Visible = False
        cmdCancel.Default = True
        lblStatus.Visible = True
        pbStatus.Value = 0
        pbStatus.Min = 0
        pbStatus.Max = 100
        pbStatus.Visible = True
        chkGames.Enabled = False
        chkCRC32.Enabled = False
        chkSids.Enabled = False
        chkScrnShots.Enabled = False
        chkExtras.Enabled = False
        fraVerify.Enabled = False
        MousePointer = vbHourglass
        DoEvents
        'do the stuff!
        m_blnBusy = True
        'create the log file
        m_intFileNum = FreeFile()
        If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "verify.txt", False, "") Then
            Kill g_strAppPathWithBackslash & "verify.txt"
        End If
        Open g_strAppPathWithBackslash & "verify.txt" For Output As #m_intFileNum
        Print #m_intFileNum, Format$(Now, "dddd dd mmmm yyyy Hh:mm:ss") & vbCrLf & g_strGameBaseDBPath & vbCrLf & vbCrLf
        'verify and cleanup!
        If chkGames.Value = vbChecked Then
            Call VerifyGames
            lblGames = g_xobjFunc.GetINIMessage("Verify Available Files", "GamesNotFound", "Games not found") & ": " & m_lngMissingGames & "/" & m_lngTotalGames
            lblGames.Visible = (chkGames.Value = vbChecked)
        End If
        If chkSids.Value = vbChecked Then
            Call VerifySids
            lblSids = g_xobjFunc.GetINIMessage("Verify Available Files", "TunesNotFound", "Tunes not found") & ": " & m_lngMissingSids & "/" & m_lngTotalSids
            lblSids.Visible = (chkSids.Value = vbChecked)
        End If
        If chkScrnShots.Value = vbChecked Then
            Call VerifyScrnShots
            lblScrnShots = g_xobjFunc.GetINIMessage("Verify Available Files", "PicturesNotFound", "Pictures not found") & ": " & m_lngMissingScrnShots & "/" & m_lngTotalScrnShots
            lblScrnShots.Visible = (chkScrnShots.Value = vbChecked)
        End If
        If chkExtras.Value = vbChecked Then
            Call VerifyExtras
            lblExtras = g_xobjFunc.GetINIMessage("Verify Available Files", "ExtrasNotFound", "Extras not found") & ": " & m_lngMissingExtras & "/" & m_lngTotalExtras
            lblExtras.Visible = (chkExtras.Value = vbChecked)
        End If
        Close #m_intFileNum
        If m_blnCancelled Then
            Unload Me
        Else
            m_blnBusy = False
            m_blnDone = True
            pbStatus.Visible = False
            lblStatus.Visible = False
            cmdStart.Caption = g_xobjFunc.GetINIMessage("Verify Available Files", "DetailsButton", "&Details") & "..."
            cmdCancel.Caption = g_strButtonCloseText
            cmdStart.Default = True
            cmdCancel.Cancel = True
            cmdStart.Enabled = True
            cmdStart.Visible = True
            cmdStart.SetFocus
            MousePointer = vbDefault
        End If
    Else
        'all done so now it is the details button
        If g_xobjFunc.FileExists(g_strAppPathWithBackslash & "verify.txt", False, "") Then
            Call g_xobjFunc.ShellExecuteFile(g_strAppPathWithBackslash & "verify.txt", "", False)
        Else
            MsgBox g_xobjFunc.GetINIMessage("Error Messages", "Error_FileNotFound", "The file was not found!") & vbCrLf & vbCrLf & g_strAppPathWithBackslash & "verify.txt", vbExclamation, gc_strAPPNAME
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description & "  (Err " & Err.Number & "," & Err.LastDllError & ")"
    Unload Me
    
End Sub
Private Sub Form_Load()

    Const DesignScaleHeight As Single = 2415    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 6375     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'set text
    cmdCancel.Caption = g_strButtonCancelText
    With g_xobjFunc
        Caption = .GetINIMessage("Verify Available Files", "Caption", "Verify Available Files") & "..."
        chkGames.Caption = .GetINIMessage("Verify Available Files", "Games", "&Games")
        chkSids.Caption = .GetINIMessage("Verify Available Files", "Music", "&Music")
        chkScrnShots.Caption = .GetINIMessage("Verify Available Files", "Pictures", "&Pictures")
        chkExtras.Caption = .GetINIMessage("Verify Available Files", "Extras", "&Extras")
        cmdStart.Caption = .GetINIMessage("Verify Available Files", "StartButton", "S&tart")
        'chkCRC32.Caption = .GetINIMessage("Verify Available Files", "CalculateCRC", "C&alculate CRC-32 Checksums") & "   [not currently used]"
    End With

    'set initial checkmarks and OK button
    chkGames.Enabled = ((g_lngTotalGamePaths > 0) And (g_blnValidGamePathExists = True))
    'chkCRC32.Enabled = chkGames.Enabled
    chkSids.Enabled = (g_udtRegPathSid.ePathType <> e_PathLocation_Invalid)
    chkScrnShots.Enabled = (g_lngTotalScrnShotPaths > 0)
    chkExtras.Enabled = (g_lngTotalExtrasPaths > 0)
    chkGames.Value = False
    chkSids.Value = False
    chkScrnShots.Value = False
    chkExtras.Value = False
    Call SetStartEnabled

    'initial flags
    m_blnBusy = False
    m_blnDone = False
    m_blnCancelled = False
    g_blnVerifyDone = False

End Sub

