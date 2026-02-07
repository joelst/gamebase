VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDiffList 
   Caption         =   "DIFF"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   Icon            =   "frmDiffList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   3945
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3660
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvGames 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   45
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   480
      MouseIcon       =   "frmDiffList.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmDiffList"
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
'       Borg Assimilation Console (BAC)
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

Private m_strDBName As String

Public Sub LoadList()

    Dim rsGames As DAO.Recordset
    Dim fldID As Field
    Dim fldName As Field

    sbStatus.Panels(1).Text = "Loading Game List..."
    Screen.MousePointer = vbHourglass
    DoEvents

    'get the db info
    Set rsGames = g_dbDiff.OpenRecordset("SELECT * FROM Info", dbOpenSnapshot)
    With rsGames
        .MoveFirst
        Caption = "DIFF [" & !GameBase & "]"
        lblFrom = !Name
        m_strDBName = !GameBase
        lblEmail = !Email
        lblDate = !Date
        .Close
    End With
    Set rsGames = Nothing

    'load the master database into the listbox
    With lvGames
        .Visible = False
        .Enabled = False
        .Sorted = False
        Call SendMessageAny(lvGames.hwnd, LVM_DELETEALLITEMS, ByVal 0&, ByVal 0&)
        Set rsGames = g_dbDiff.OpenRecordset("SELECT GA_Id, Name FROM Games ORDER BY Name", dbOpenSnapshot)
        If Not rsGames.EOF Then
            Set fldID = rsGames!GA_Id
            Set fldName = rsGames!Name
            rsGames.MoveFirst
            Do Until rsGames.EOF
                .ListItems.Add , "G" & fldID, fldName
                rsGames.MoveNext
            Loop
            rsGames.Close
            Set fldID = Nothing
            Set fldName = Nothing
        End If
        Set rsGames = Nothing
        .Sorted = True
        .Visible = True
        .Enabled = True
    End With
    sbStatus.Panels(1).Text = lvGames.ListItems.Count & " Games"
    Show
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

    Dim sngRegWidth As Single
    Dim sngRegHeight As Single
    Dim sngRegTop As Single
    Dim sngRegLeft As Single
    Dim blnMaximized As Boolean
    Dim sngCol1Width As Long

    'set window size/position
    With g_xobjRegistry
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListWidth", sngRegWidth, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListHeight", sngRegHeight, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListTop", sngRegTop, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListLeft", sngRegLeft, False, -1
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListMaximized", blnMaximized, False, False
        .GetSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListCol1Width", sngCol1Width, False, 3509
    End With
    If sngRegWidth < 2000 Then
        sngRegWidth = 3990
    ElseIf sngRegWidth > Screen.Width Then
        sngRegWidth = Screen.Width
    End If
    If sngRegHeight < 500 Then
        sngRegHeight = 2865
    ElseIf sngRegHeight > Screen.Height Then
        sngRegHeight = Screen.Height
    End If
    If sngRegTop < 0 Or sngRegTop > Screen.Height Or sngRegLeft < 0 Or sngRegLeft > Screen.Width Then
        sngRegTop = 0
        sngRegLeft = 0
    End If
    Move sngRegLeft, sngRegTop, sngRegWidth, sngRegHeight
    If blnMaximized Then
        WindowState = vbMaximized
    End If
    lvGames.ColumnHeaders(1).Width = sngCol1Width
    
End Sub


Private Sub Form_Resize()
    
    If ScaleWidth > 0 Then
        lvGames.Width = ScaleWidth
        sbStatus.Panels(1).Width = ScaleWidth
    End If
    If ScaleHeight > (lvGames.Top + sbStatus.Height) Then
        lvGames.Height = ScaleHeight - lvGames.Top - sbStatus.Height
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Hide
    DoEvents

    'save window size/position
    With g_xobjRegistry
        If WindowState = vbNormal Then
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListWidth", Width, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListHeight", Height, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListTop", Top, False
            .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListLeft", Left, False
        End If
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListMaximized", (WindowState = vbMaximized), False
        .SaveSetting "HKEY_LOCAL_MACHINE", "Software\GB64\BAC\Window", "DiffListCol1Width", lvGames.ColumnHeaders(1).Width, False
    End With

    g_blnDiffOpen = False
    g_strDiff = ""
    g_dbDiff.Close
    Set g_dbDiff = Nothing

End Sub


Private Sub lblEmail_Click()

    g_xobjFunc.ShellExecuteFile "mailto:" & lblEmail & "?subject=Re: Your submitted DIFF [" & m_strDBName & "] dated [" & lblDate & "]", "", False

End Sub

Private Sub lvGames_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim LItem As ListItem

    If g_blnMasterOpen = True Then
        'select the game in the other list
        With frmMasterList
            If .WindowState = vbMinimized Then
                .WindowState = vbNormal
            End If
            If ListItemExists(.lvGames, Item.Key) Then
                'set the master list to point to this game
                Set LItem = .lvGames.ListItems(Item.Key)
                .lvGames.SelectedItem = LItem
                .lvGames.SelectedItem.Selected = True
                .lvGames.SelectedItem.EnsureVisible
                Set LItem = Nothing
                Call .SelectGame(Item.Key)
            Else
                'set the master list to point to no game
                Call .SelectNoGame
            End If
        End With
    End If

End Sub

Public Sub SelectGame(strKey_in As String)

    'display the info for the game

    

End Sub
