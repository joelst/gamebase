VERSION 5.00
Begin VB.Form frmTextInput 
   BorderStyle     =   0  'None
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmTextInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNotAllowed As String

Public Sub GetText(strDefault As String, ByVal lngMaxChars As Long, strInvalidChars As String, sngTextBoxWidth As Single, sngTop As Single, sngLeft As Single)

    Top = sngTop
    Left = sngLeft
    With txtInput
        .Width = sngTextBoxWidth
        Width = .Width
        Height = .Height
        .MaxLength = lngMaxChars
        .Text = strDefault
    End With
    strNotAllowed = strInvalidChars

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        blnGlobalGetText = True
        strGlobalGetText = txtInput
        Unload Me
        
    ElseIf KeyCode = 27 Then
        blnGlobalGetText = False
        strGlobalGetText = ""
        Unload Me
        
    End If

End Sub

Private Sub txtInput_GotFocus()
    
    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)

    If InStr(strNotAllowed, Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If

End Sub


