VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRezipAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GEMUS"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the files you wish to Rezip:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmRezipAsk"
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

Private m_blnActivated As Boolean

Private Type UnzippedFileAskType
    strName As String
    strUCaseName As String
    strCRC As String
    blnChecked As Boolean
End Type
Private a_Rezip() As UnzippedFileAskType
Public lngTotalFiles As Long
Public Sub AddFile(strName_in As String, strUCaseName_in As String, strCRC_in As String, ByVal blnChecked_in As Boolean)

    ReDim Preserve a_Rezip(lngTotalFiles)
    With a_Rezip(lngTotalFiles)
        .blnChecked = blnChecked_in
        .strCRC = strCRC_in
        .strName = strName_in
        .strUCaseName = strUCaseName_in
    End With
    lngTotalFiles = lngTotalFiles + 1

End Sub


Public Sub GetFile(ByVal lngIdx_in As Long, strName_out As String, strUCaseName_out As String, strCRC_out As String, blnChecked_out As Boolean)

    With a_Rezip(lngIdx_in)
        strName_out = .strName
        strUCaseName_out = .strUCaseName
        strCRC_out = .strCRC
        blnChecked_out = .blnChecked
    End With

End Sub


Private Sub cmdOK_Click()

    Dim lngIdx As Long
    
    'save the checkmarks
    For lngIdx = 0 To (lngTotalFiles - 1)
        a_Rezip(lngIdx).blnChecked = lvFiles.ListItems("F" & Trim$(lngIdx)).Checked
    Next lngIdx

    Hide

End Sub

Private Sub Form_Activate()

    Dim lngIdx As Long
    Dim LItem As ListItem

    If m_blnActivated = True Then Exit Sub
    m_blnActivated = True

    For lngIdx = 0 To (lngTotalFiles - 1)
        With a_Rezip(lngIdx)
            Set LItem = lvFiles.ListItems.Add(, "F" & Trim$(lngIdx), .strName)
            LItem.Checked = .blnChecked
        End With
    Next lngIdx

End Sub

Private Sub Form_Load()

    Dim strAppPathWithBackSlash As String
    Dim strMessFile As String

    Width = 4575 + (Width - ScaleWidth)
    Height = 3975 + (Height - ScaleHeight)

    m_blnActivated = False

    'set the text
    strAppPathWithBackSlash = App.Path
    If Right$(strAppPathWithBackSlash, 1) <> "\" Then strAppPathWithBackSlash = strAppPathWithBackSlash & "\"
    strMessFile = strAppPathWithBackSlash & "GemusText.ini"
    cmdOK.Caption = GetINIMessageFromFile("General Buttons", "OK", "&OK", strMessFile)

    lngTotalFiles = 0
    ReDim a_Rezip(0)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Hide
    End If

End Sub

