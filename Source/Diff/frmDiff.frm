VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDiff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GameBase Diff-Program"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ControlBox      =   0   'False
   Icon            =   "frmDiff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLetters 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &NONE"
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &ALL"
         Height          =   375
         Left            =   2400
         TabIndex        =   35
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "Z"
         Height          =   255
         Index           =   26
         Left            =   1680
         TabIndex        =   34
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "Y"
         Height          =   255
         Index           =   25
         Left            =   1680
         TabIndex        =   33
         Top             =   2160
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "X"
         Height          =   255
         Index           =   24
         Left            =   1680
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "W"
         Height          =   255
         Index           =   23
         Left            =   1680
         TabIndex        =   31
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "V"
         Height          =   255
         Index           =   22
         Left            =   1680
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "U"
         Height          =   255
         Index           =   21
         Left            =   1680
         TabIndex        =   29
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "T"
         Height          =   255
         Index           =   20
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "S"
         Height          =   255
         Index           =   19
         Left            =   1680
         TabIndex        =   27
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "R"
         Height          =   255
         Index           =   18
         Left            =   1680
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "Q"
         Height          =   255
         Index           =   17
         Left            =   960
         TabIndex        =   25
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "P"
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   24
         Top             =   2160
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "O"
         Height          =   255
         Index           =   15
         Left            =   960
         TabIndex        =   23
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "N"
         Height          =   255
         Index           =   14
         Left            =   960
         TabIndex        =   22
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "M"
         Height          =   255
         Index           =   13
         Left            =   960
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "L"
         Height          =   255
         Index           =   12
         Left            =   960
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "K"
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   19
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "J"
         Height          =   255
         Index           =   10
         Left            =   960
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "I"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "H"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "G"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "F"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "E"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "D"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "C"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "B"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "A"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "0-9"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the letters you want to search:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox lstGames 
      Enabled         =   0   'False
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Games.  Please wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "frmDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'              == GameBase ==
'    The Universal Retro-Gaming Frontend
'           and Database Utility
'----------------------------------------------
' Copyright (C) 1997-2004 James Burrows
' Email: gamebase@itsa.me.uk
' Homepage: http://www.bu22.com/
'----------------------------------------------
' DIFF (Finds differences between 2 databases)
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

Private Backup As String
Private Current As String
Private Template As String
Private AppPath As String
Private SQL As String
Private MSQL As String

Private Function GetDatabaseVersion(strDBFile_in As String) As Single

    Dim dbGameBase As DAO.Database
    Dim rsGameBase As Recordset
    Dim lngDBMajorVersion As Long
    Dim lngDBMinorVersion As Long
    Dim lngTestVar As Long

    lngTestVar = 1
    On Error GoTo Err_Handler

    Set dbGameBase = OpenDatabase(strDBFile_in)
    lngTestVar = 2
    Set rsGameBase = dbGameBase.OpenRecordset("Config", dbOpenSnapshot)
    With rsGameBase
        If Not .EOF Then
            .MoveFirst
            lngDBMajorVersion = !MajorVersion
            lngDBMinorVersion = !MinorVersion
            .Close
        Else
            lngDBMajorVersion = 1
            lngDBMinorVersion = 3
            .Close
        End If
    End With
    Set rsGameBase = Nothing
    GetDatabaseVersion = CSng(lngDBMajorVersion & "." & lngDBMinorVersion)
    dbGameBase.Close
    Set dbGameBase = Nothing
    Exit Function

Err_Handler:
    GetDatabaseVersion = 1.3
    If lngTestVar = 2 Then
        dbGameBase.Close
    End If
    Set dbGameBase = Nothing

End Function
Private Sub FindGames()

    Dim dbUpdate As Database
    Dim dbBackup As Database
    Dim dbCurrent As Database
    Dim rsUpdate As Recordset
    Dim rsBackup As Recordset
    Dim rsCurrent As Recordset
    Dim rsLink As Recordset
    Dim Count As Long
    Dim strPub As String
    Dim strMus As String
    Dim strPro As String
    Dim strLan As String
    Dim strCra As String
    Dim AddIt As Boolean
    
    Screen.MousePointer = vbHourglass
    cmdOK.Visible = False
    cmdOK.Enabled = False
    cmdCancel.Visible = False
    cmdCancel.Enabled = False
    lblPrompt = "Changed Games Found: 0"
    fraLetters.Visible = False
    fraLetters.Enabled = False
    pbStatus.Min = 0
    pbStatus.Value = 0
    pbStatus.Visible = True
    lblStatus.Visible = True
    DoEvents
    
    SQL = "SELECT * FROM Games WHERE ("
    For Count = 0 To 26
        If chkLetter(Count).Value = 1 Then
            If Count > 0 Then
                SQL = SQL & "(Name LIKE """ & Chr$(Count + 64) & "*"")"
            Else
                SQL = SQL & "(NOT((Name LIKE ""a*"") OR (Name LIKE ""b*"") OR (Name LIKE ""c*"") OR (Name LIKE ""d*"") OR (Name LIKE ""e*"") OR (Name LIKE ""f*"") OR (Name LIKE ""g*"") OR (Name LIKE ""h*"") OR (Name LIKE ""i*"") OR (Name LIKE ""j*"") OR (Name LIKE ""k*"") OR (Name LIKE ""l*"") OR (Name LIKE ""m*"") OR (Name LIKE ""n*"") OR (Name LIKE ""o*"") OR (Name LIKE ""p*"") OR (Name LIKE ""q*"") OR (Name LIKE ""r*"") OR (Name LIKE ""s*"") OR (Name LIKE ""t*"") OR (Name LIKE ""u*"") OR (Name LIKE ""v*"") OR (Name LIKE ""w*"") OR (Name LIKE ""x*"") OR (Name LIKE ""y*"") OR (Name LIKE ""z*"")))"
            End If
            SQL = SQL & " OR "
        End If
    Next Count
    SQL = Left$(SQL, Len(SQL) - 4)
    SQL = SQL & ") ORDER BY Name"

    Set dbUpdate = OpenDatabase(Template)
    Set dbBackup = OpenDatabase(Backup)
    Set dbCurrent = OpenDatabase(Current)
    Set rsUpdate = dbUpdate.OpenRecordset("select * from Games", dbOpenDynaset)
    Set rsBackup = dbBackup.OpenRecordset("select * from Games", dbOpenSnapshot)
    Set rsCurrent = dbCurrent.OpenRecordset(SQL, dbOpenSnapshot)

    With rsCurrent
        If .EOF Then
            MsgBox "Current Database is Empty!", vbInformation, "GameBase Diff-Program"
            rsUpdate.Close
            rsBackup.Close
            rsCurrent.Close
            Set rsUpdate = Nothing
            Set rsBackup = Nothing
            Set rsCurrent = Nothing
            Set dbUpdate = Nothing
            Set dbBackup = Nothing
            Set dbCurrent = Nothing
            End
        Else
            .MoveLast
            .MoveFirst
            pbStatus.Max = .RecordCount
            Count = 0
            Do Until .EOF
                lblStatus = !Name
                pbStatus.Value = .AbsolutePosition
                DoEvents
                rsBackup.FindFirst "GA_ID = " & Trim$(!GA_Id)
                If Not .NoMatch Then
                    If !Name <> rsBackup!Name _
                    Or !PU_Id <> rsBackup!PU_Id Or !YE_Id <> rsBackup!YE_Id _
                    Or !PR_Id <> rsBackup!PR_Id Or !MU_Id <> rsBackup!MU_Id _
                    Or !GE_Id <> rsBackup!GE_Id Or !PlayersFrom <> rsBackup!PlayersFrom _
                    Or !PlayersTo <> rsBackup!PlayersTo Or !PlayersSim <> rsBackup!PlayersSim _
                    Or !LA_Id <> rsBackup!LA_Id Or !Comment <> rsBackup!Comment _
                    Or !CR_Id <> rsBackup!CR_Id Or !V_Trainers <> rsBackup!V_Trainers _
                    Or !V_Length <> rsBackup!V_Length Or !V_TrueDriveEmu <> rsBackup!V_TrueDriveEmu _
                    Or !V_HighScoreSaver <> rsBackup!V_HighScoreSaver Or !V_LoadingScreen <> rsBackup!V_LoadingScreen _
                    Or !V_IncludedDocs <> rsBackup!V_IncludedDocs Or !V_PalNTSC <> rsBackup!V_PalNTSC _
                    Or !V_Comment <> rsBackup!V_Comment _
                    Or !Filename <> rsBackup!Filename Or !FileToRun <> rsBackup!FileToRun _
                    Or !FilenameIndex <> rsBackup!FilenameIndex Or !ScrnShotFilename <> rsBackup!ScrnShotFilename _
                    Or !SidFilename <> rsBackup!SidFilename Or !Adult <> rsBackup!Adult _
                    Or !MemoText <> rsBackup!MemoText Or !Control <> rsBackup!Control _
                    Or !Prequel <> rsBackup!Prequel Or !Sequel <> rsBackup!Sequel Or !Related <> rsBackup!Related _
                    Or !CRC <> rsBackup!CRC _
                    Then
                        AddIt = True
                    Else
                        AddIt = False
                    End If
                Else
                    AddIt = True
                End If

                If AddIt = True Then
                    Count = Count + 1
                    lblPrompt = "Changes Found: " & Trim$(Count)
                    lstGames.AddItem !Name
                    'add the record
                    Set rsLink = dbCurrent.OpenRecordset("SELECT Publisher FROM Publishers WHERE PU_Id = " & Trim$(!PU_Id), dbOpenSnapshot)
                    rsLink.MoveFirst
                    strPub = rsLink!Publisher
                    Set rsLink = Nothing
                    Set rsLink = dbCurrent.OpenRecordset("SELECT Musician AS Name FROM Musicians WHERE MU_Id = " & Trim$(!MU_Id), dbOpenSnapshot)
                    rsLink.MoveFirst
                    strMus = rsLink!Name
                    Set rsLink = Nothing
                    Set rsLink = dbCurrent.OpenRecordset("SELECT Programmer FROM Programmers WHERE PR_Id = " & Trim$(!PR_Id), dbOpenSnapshot)
                    rsLink.MoveFirst
                    strPro = rsLink!Programmer
                    Set rsLink = Nothing
                    Set rsLink = dbCurrent.OpenRecordset("SELECT Language FROM Languages WHERE LA_Id = " & Trim$(!LA_Id), dbOpenSnapshot)
                    rsLink.MoveFirst
                    strLan = rsLink!Language
                    Set rsLink = Nothing
                    Set rsLink = dbCurrent.OpenRecordset("SELECT Cracker FROM Crackers WHERE CR_Id = " & Trim$(!CR_Id), dbOpenSnapshot)
                    rsLink.MoveFirst
                    strCra = rsLink!Cracker
                    Set rsLink = Nothing
                    rsUpdate.AddNew
                    rsUpdate!GA_Id = !GA_Id
                    rsUpdate!Name = !Name
                    rsUpdate!Publisher = strPub
                    rsUpdate!YE_Id = !YE_Id
                    rsUpdate!Musician = strMus
                    rsUpdate!Programmer = strPro
                    rsUpdate!Language = strLan
                    rsUpdate!Cracker = strCra
                    rsUpdate!Genre = !GE_Id
                    rsUpdate!PlayersFrom = !PlayersFrom
                    rsUpdate!PlayersTo = !PlayersTo
                    rsUpdate!PlayersSim = !PlayersSim
                    rsUpdate!Comment = !Comment
                    rsUpdate!V_Trainers = !V_Trainers
                    rsUpdate!V_Length = !V_Length
                    rsUpdate!V_TrueDriveEmu = !V_TrueDriveEmu
                    rsUpdate!V_HighScoreSaver = !V_HighScoreSaver
                    rsUpdate!V_LoadingScreen = !V_LoadingScreen
                    rsUpdate!V_IncludedDocs = !V_IncludedDocs
                    rsUpdate!V_PalNTSC = !V_PalNTSC
                    rsUpdate!V_Comment = !V_Comment
                    rsUpdate!Filename = !Filename
                    rsUpdate!FileToRun = !FileToRun
                    rsUpdate!FilenameIndex = !FilenameIndex
                    rsUpdate!ScrnShotFilename = !ScrnShotFilename
                    rsUpdate!SidFilename = !SidFilename
                    rsUpdate!Adult = !Adult
                    rsUpdate!MemoText = !MemoText
                    rsUpdate!Prequel = !Prequel
                    rsUpdate!Sequel = !Sequel
                    rsUpdate!Related = !Related
                    rsUpdate!Control = !Control
                    rsUpdate!CRC = !CRC
                    rsUpdate.Update
                End If
                .MoveNext
            Loop
        End If
    End With
    
    rsUpdate.Close
    rsBackup.Close
    rsCurrent.Close
    Set rsUpdate = Nothing
    Set rsBackup = Nothing
    Set rsCurrent = Nothing
    Set dbUpdate = Nothing
    Set dbBackup = Nothing
    Set dbCurrent = Nothing

    Screen.MousePointer = vbDefault
    lstGames.Visible = True
    lstGames.Enabled = True
    cmdOK.Caption = "&OK"
    cmdOK.Left = (Width / 2) - (cmdOK.Width / 2)
    cmdOK.Visible = True
    cmdOK.Enabled = True
    pbStatus.Visible = False
    lblStatus.Visible = False
    lblPrompt = "Changed Games: " & Trim$(Count)
    lstGames.SetFocus

End Sub

Private Function GetFile(ByVal Filename As String, ByVal Prompt As String) As String

    Dim Folder As String

    On Error GoTo Exit_Here
    
    GetFile = ""
    
    With cDialog
        .CancelError = True
        .DialogTitle = Prompt
        .Filename = Filename
        .Filter = Filename & "|" & Filename
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNNoReadOnlyReturn
        .InitDir = App.Path
        .ShowOpen
        GetFile = .Filename
    End With
    
Exit_Here:

End Function

Private Sub chkLetter_Click(Index As Integer)

    Dim a As Long
    Dim OK As Boolean
    
    OK = False
    For a = 0 To 26
        If chkLetter(a).Value = 1 Then
            OK = True
        End If
    Next a
    
    cmdOK.Enabled = (OK)

End Sub

Private Sub cmdAll_Click()

    Dim a As Long
    
    For a = 0 To 26
        chkLetter(a).Value = 1
    Next a

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNone_Click()

    Dim a As Long
    
    For a = 0 To 26
        chkLetter(a).Value = 0
    Next a
    
End Sub

Private Sub cmdOK_Click()

    Dim x As String
    Dim e As String
    Dim g As String
    Dim n As String

    If cmdOK.Caption = "&Next >" Then
        fraLetters.Enabled = False
        fraLetters.Visible = False
        cmdOK.Caption = "&Start"
    ElseIf cmdOK.Caption = "&Start" Then
        FindGames
    Else
        If lstGames.ListCount > 0 Then
            x = ""
            Do
                x = InputBox("Type your name so the master-database maintainer knows who this update is from:", "Update Author")
                If x = "" Then
                    MsgBox "You must type a name!", vbInformation, "GameBase Diff-Program"
                ElseIf Len(x) > 255 Then
                    MsgBox "Name must not exceed 255 characters!", vbInformation, "GameBase Diff-Program"
                Else
                    Exit Do
                End If
            Loop
            
            e = ""
            Do
                e = InputBox("Type your email address so the master-database maintainer knows how to contact you for any questions:", "Update Email Address")
                If e = "" Then
                    MsgBox "You must type an email address!", vbInformation, "GameBase Diff-Program"
                ElseIf Len(e) > 255 Then
                    MsgBox "Email address must not exceed 255 characters!", vbInformation, "GameBase Diff-Program"
                Else
                    Exit Do
                End If
            Loop
            
            g = ""
            Do
                g = InputBox("Type the name of your GameBase database, so the master-database maintainer knows which database to merge your update with:", "Update GameBase Database Name")
                If g = "" Then
                    MsgBox "You must type a database name!", vbInformation, "GameBase Diff-Program"
                ElseIf Len(g) > 255 Then
                    MsgBox "Database name must not exceed 255 characters!", vbInformation, "GameBase Diff-Program"
                Else
                    Exit Do
                End If
            Loop

            n = ""
            n = InputBox("Type any Notes that you wish the master-database maintainer to see when merging this update:", "Update Notes")
            
            Dim db As Database
            Dim rs As Recordset
            Set db = OpenDatabase(Template)
            Set rs = db.OpenRecordset("Info")
            With rs
                .AddNew
                !Name = x
                !Email = e
                !GameBase = g
                !Notes = n
                !Date = Format$(Now, "dd mmm yyyy")
                .Update
                .Close
            End With
            Set rs = Nothing
            Set db = Nothing
            MsgBox "You now need to ZIP and EMAIL the file """ & Template & """ to the master-database maintainer." & vbCrLf & vbCrLf & "Do not delete this file because it is used in future updates.", vbInformation, "GameBase Diff-Program"
            If MsgBox("You should replace the BACKUP database with your CURRENT database before collecting anymore info." & vbCrLf & vbCrLf & "Do you want me to do it for you now?", vbQuestion Or vbYesNo, "GameBase Diff-Program") = vbYes Then
                Screen.MousePointer = vbHourglass
                DoEvents
                Kill Backup
                FileCopy Current, Backup
                Screen.MousePointer = vbDefault
                DoEvents
                MsgBox "Backup database successfully replaced!" & vbCrLf & vbCrLf & "You may now continue to collect info with your current database!", vbInformation, "GameBase Diff-Program"
            Else
                MsgBox "Don't forget to do it yourself then!", vbExclamation, "GameBase Diff-Program"
            End If
        End If
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    If App.PrevInstance = True Then
        End
    End If

    Dim DesignHeight As Single
    Dim DesignScaleHeight As Single
    Dim DesignWidth As Single
    Dim DesignScaleWidth As Single
    Dim Difference As Single
    Dim dbUpdate As Database
    Dim rsGames As Recordset
    Dim UpdateReady As Boolean
    Dim dbEng As DBEngine
    Dim sngVersion As Single

    'form dimensions on large/small fonts
    DesignHeight = 3720
    DesignScaleHeight = 3345
    DesignWidth = 4200
    DesignScaleWidth = 4110
    Me.Height = DesignHeight
    Me.Width = DesignWidth
    Difference = DesignScaleHeight - Me.ScaleHeight
    Me.Height = DesignHeight + Difference
    Difference = DesignScaleWidth - Me.ScaleWidth
    Me.Width = DesignWidth + Difference

    'get databases to test
    Current = GetFile("*.mdb", "Select your CURRENT CHANGED database")
    If Current = "" Then
        MsgBox "Diff-program aborted!", vbInformation, "GameBase Diff-Program"
        End
    End If
    sngVersion = GetDatabaseVersion(Current)
    If sngVersion <> 2.1 Then
        MsgBox "This database is not of the required version for this version of DIFF.EXE. Required version is 2.1. Version found is " & Format$(sngVersion, "0.0"), vbInformation, "GameBase Diff-Program"
        End
    End If
    Backup = GetFile("*.mdb", "Select the BACKED-UP PRE-CHANGED database")
    If Backup = "" Then
        MsgBox "Diff-program aborted!", vbInformation, "GameBase Diff-Program"
        End
    End If
    sngVersion = GetDatabaseVersion(Backup)
    If sngVersion <> 2.1 Then
        MsgBox "This database is not of the required version for this version of DIFF.EXE. Required version is 2.1. Version found is " & Format$(sngVersion, "0.0"), vbInformation, "GameBase Diff-Program"
        End
    End If

    'check destination database exists and is empty
    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    Template = AppPath & "update.mdb"
    If Dir(Template) = "" Then
        MsgBox "Cannot continue: """ & Template & """ Not Found!", vbInformation, "GameBase Diff-Program"
        End
    Else
        sngVersion = GetDatabaseVersion(Template)
        If sngVersion <> 2.1 Then
            MsgBox "The update.mdb file is not of the required version for this version of DIFF.EXE. Required version is 2.1. Version found is " & Format$(sngVersion, "0.0"), vbInformation, "GameBase Diff-Program"
            End
        End If
        Set dbUpdate = OpenDatabase(Template)
        Set rsGames = dbUpdate.OpenRecordset("Games", dbOpenSnapshot)
        UpdateReady = rsGames.EOF
        rsGames.Close
        Set rsGames = Nothing
        dbUpdate.Close
        Set dbUpdate = Nothing
        If Not UpdateReady Then
            If MsgBox("The UPDATE.MDB is not empty" & vbCrLf & vbCrLf & "Running the Diff-Program now will overwrite all the records in UPDATE.MDB" & vbCrLf & vbCrLf & "Do you want to overwrite these records and continue?", vbQuestion Or vbYesNo, "GameBase Diff-Program") = vbYes Then
                'clear the database
                Set dbUpdate = OpenDatabase(Template)
                dbUpdate.Execute "DELETE * FROM Games"
                dbUpdate.Execute "DELETE * FROM Info"
                Set dbUpdate = Nothing
            Else
                MsgBox "Diff-Program aborted!", vbInformation, "GameBase Diff-Program"
                End
            End If
        End If
    End If
    
    'compact the database
    If Dir(AppPath & "compact.mdb") <> "" Then
        Kill AppPath & "compact.mdb"
    End If
    Set dbEng = New DBEngine
    dbEng.CompactDatabase Template, AppPath & "compact.mdb"
    Set dbEng = Nothing
    Kill Template
    Name (AppPath & "compact.mdb") As Template
    
    'get the form ready
    lblPrompt = "Ready to find changes between:-" & vbCrLf & vbCrLf & Current & vbCrLf & vbCrLf & "and" & vbCrLf & vbCrLf & Backup & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Press Start to begin..."

End Sub

