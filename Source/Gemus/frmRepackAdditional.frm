VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepackAdditional 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repacking Archive: foo.zip"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   ControlBox      =   0   'False
   Icon            =   "frmRepackAdditional.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdNone 
      Caption         =   "&None"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&All"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblPrompt 
      Caption         =   "The following new files appeared whilst you were playing. Please select the ones you would like include in the repack:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmRepackAdditional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SetAdditionalFiles(a_strAdditional_in() As String)

    Dim LItem As ListItem
    Dim lngIdx As Long
    Dim lngTotalFiles As Long

    lngTotalFiles = UBound(a_strAdditional_in) + 1
    With lvFiles
        .Enabled = False
        .Visible = False
        For lngIdx = 0 To lngTotalFiles - 1
            Set LItem = .ListItems.Add(, , a_strAdditional_in(lngIdx))
            LItem.Checked = True
        Next lngIdx
        .Enabled = True
        .Visible = True
    End With

End Function
Public Function GetAdditionalFiles(a_strAdditional_out() As String, lngTotal_out As Long)

    Dim LItem As ListItem

    ReDim a_strAdditional_out(0)
    lngTotal_out = 0

    With lvFiles
        For Each LItem In .ListItems
            If LItem.Checked Then
                ReDim Preserve a_strAdditional_out(lngTotal_out)
                a_strAdditional_out(lngTotal_out) = LItem.Text
                lngTotal_out = lngTotal_out + 1
            End If
        Next LItem
    End With

End Function

Public Sub SetArchiveName(strArchiveName_in As String)

    Caption = "Repacking Archive: " & strArchiveName_in

End Sub

Private Sub SetListItemsChecked(ByVal blnCheckedValue_in As Boolean)

    Dim LItem As ListItem
    
    With lvFiles
        .Enabled = False
        .Visible = False
        For Each LItem In .ListItems
            LItem.Checked = blnCheckedValue_in
        Next LItem
        .Enabled = True
        .Visible = True
    End With

End Sub

Private Sub cmdAll_Click()

    Call SetListItemsChecked(True)
    lvFiles.SetFocus

End Sub

Private Sub cmdNone_Click()

    Call SetListItemsChecked(False)
    lvFiles.SetFocus

End Sub


Private Sub cmdOK_Click()

    Hide

End Sub

Private Sub Form_Activate()

    On Error GoTo Err_Handler

    lvFiles.SetFocus

Err_Handler:

End Sub

Private Sub Form_Load()

    Const DesignScaleHeight As Single = 5535    'The required scaleheight for the form
    Const DesignScaleWidth As Single = 9255     'The required scalewidth for the form

    'set the form's proper width and height
    Height = DesignScaleHeight + (Height - ScaleHeight)
    Width = DesignScaleWidth + (Width - ScaleWidth)

    'column width
    lvFiles.ColumnHeaders(1).Width = lvFiles.Width - 385

End Sub


