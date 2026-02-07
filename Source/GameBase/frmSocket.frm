VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSocket 
   BorderStyle     =   0  'None
   Caption         =   "GameBase Version Checker"
   ClientHeight    =   225
   ClientLeft      =   6345
   ClientTop       =   3630
   ClientWidth     =   2070
   Icon            =   "frmSocket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock Socket 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking for new version..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSocket"
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

Private m_strRequest As String
Private Function URLEncode(ByVal strToEncode As String) As String

    Dim lngIdx As Long
    Dim lngLen As Long
    Dim strChar As String
    Dim strDest As String

    URLEncode = ""
    lngLen = Len(strToEncode) - 1
    For lngIdx = 1 To lngLen
        strChar = Mid$(strToEncode, lngIdx, 1)
        Select Case strChar
            Case " "
                strDest = "%20"
            Case "/"
                strDest = "%2F"
            Case ":"
                strDest = "%3A"
            Case "$"
                strDest = "%24"
            Case "&"
                strDest = "%26"
            Case "%"
                strDest = "%25"
            Case ","
                strDest = "%2C"
            Case "="
                strDest = "%3D"
            Case Else
                strDest = strChar
        End Select
        URLEncode = URLEncode & strDest
    Next lngIdx

End Function

Private Sub Form_Activate()

    Dim strRequest As String
    Dim strPost As String
    Dim strHostPort As String
    Dim strProxyHostPost As String
    Dim strDir As String
    Dim strDBs As String
    Dim strEmus As String
    Dim lngTotalEmus As Long
    Dim strTemp As String
    Dim strPath As String
    Dim strFile As String
    Dim strConnectHostPort As String
    Dim lngColonPos As Long
    Dim strConnHost As String
    Dim lngConnPort As Long

    On Error GoTo Err_Handler

    'hide form
    Visible = False
    DoEvents

    'formulate the http request string
    strPost = "vcv=1&maj=" & App.Major & "&min=" & App.Minor

    'get all gb databases and emulators
    If Val(g_xobjFunc.INI_GetSetting("Startup", "SendUsageStats", "0", g_strGameBaseDotINI)) = 1 Then
        strDBs = ""
        strEmus = ""
        strDir = Dir(g_strAppPathWithBackslash, vbDirectory)
        Do While strDir <> ""
            If strDir <> "." And strDir <> ".." Then
                If (GetAttr(g_strAppPathWithBackslash & strDir) And vbDirectory) = vbDirectory Then
                    'found possible db folder
                    If g_xobjFunc.FileExists(g_strAppPathWithBackslash & strDir & "\" & strDir & ".mdb", False, "") Then
                        'it's a GB database
                        strDBs = strDBs & strDir & "|"
                        'now get a list of all emulators for the current db
                        lngTotalEmus = 0
                        Do
                            strTemp = g_xobjFunc.INI_GetSetting("Emulators", Trim$(lngTotalEmus + 1), "", g_strAppPathWithBackslash & strDir & "\" & "Emulators.ini")
                            If Trim$(strTemp) <> "" Then
                                strTemp = Trim$(g_xobjFunc.INI_GetSetting(strTemp, "EmulatorPathFile", "", g_strAppPathWithBackslash & strDir & "\" & "Emulators.ini"))
                                If strTemp <> "" Then
                                    g_xobjFunc.GetFileAndPath strTemp, strPath, strFile
                                    If strFile <> "" Then
                                        strEmus = strEmus & strFile & "|"
                                    End If
                                End If
                                lngTotalEmus = lngTotalEmus + 1
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            End If
            strDir = Dir()
        Loop
        strPost = strPost & "&databases=" & URLEncode(strDBs) & "&emulators=" & URLEncode(strEmus)
    Else
        strPost = strPost & "&databases=&emulators="
    End If

    'get the url/proxy to go to
    strHostPort = Trim$(g_xobjFunc.INI_GetSetting("Startup", "VersionCheckHostPort", "", g_strGameBaseDotINI))
    strProxyHostPost = Trim$(g_xobjFunc.INI_GetSetting("Startup", "ProxyHostPort", "", g_strGameBaseDotINI))

    'formulate the request string
    If strProxyHostPost = "" Then
        strConnectHostPort = strHostPort
        m_strRequest = "POST /gbvc.php?checkit=274277 HTTP/1.0" & vbCrLf & _
                       "Host: " & strHostPort & vbCrLf & _
                       "Content-type: application/x-www-form-urlencoded" & vbCrLf & _
                       "Content-length: " & Len(strPost) & vbCrLf & _
                       "Connection: close" & vbCrLf & vbCrLf & strPost
    Else
        strConnectHostPort = strProxyHostPost
        m_strRequest = "POST http://" & strHostPort & "/gbvc.php?checkit=274277 HTTP/1.0" & vbCrLf & _
                       "Host: " & strProxyHostPost & vbCrLf & _
                       "Content-type: application/x-www-form-urlencoded" & vbCrLf & _
                       "Content-length: " & Len(strPost) & vbCrLf & _
                       "Proxy-connection: close" & vbCrLf & _
                       "Connection: close" & vbCrLf & vbCrLf & strPost
    End If

    'seperate connection host and port
    lngColonPos = InStrRev(strConnectHostPort, ":")
    If lngColonPos > 0 Then
        strConnHost = Left$(strConnectHostPort, lngColonPos - 1)
        lngConnPort = Right$(strConnectHostPort, Len(strConnectHostPort) - lngColonPos)
    Else
        strConnHost = strConnectHostPort
        lngConnPort = 80
    End If

    'connect to server
    With Socket
        .RemoteHost = strConnHost
        .RemotePort = lngConnPort
        .Protocol = sckTCPProtocol
        .Connect
    End With

    Exit Sub

Err_Handler:
    Socket.Close
    Unload Me

End Sub

Private Sub Form_Load()

    g_blnLoadedfrmSocket = True

End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmMain.sbStatus.Panels(4) = "Done"
    DoEvents
    frmMain.sbStatus.Panels(4) = ""
    DoEvents
    g_blnLoadedfrmSocket = False

End Sub


Private Sub Socket_Connect()

    On Error GoTo Err_Handler

    Socket.SendData m_strRequest

    Exit Sub

Err_Handler:
    Socket.Close
    Unload Me

End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    Dim strResponse As String
    Dim lngStart As Long
    Dim a_strResponse() As String
    Dim a_strVersion() As String
    Dim strDownloadURL As String
    Dim strVCheckHost As String
    Dim lngMaj As String
    Dim lngMin As String
    Dim lngRev As Long
    Dim blnBeta As Boolean
    Dim blnNewVersion As Boolean

    On Error GoTo Err_Handler

    Socket.GetData strResponse
    Socket.Close
    
    a_strResponse = Split(strResponse, vbCrLf & vbCrLf, 2)
    
    If Len(a_strResponse(1)) > 0 Then
        a_strResponse = Split(a_strResponse(1), "|", 3)
        If Len(a_strResponse(0)) > 0 Then
            a_strVersion = Split(a_strResponse(0), ".", 4)
            strDownloadURL = a_strResponse(1)
            strVCheckHost = a_strResponse(2)
            lngMaj = Int(a_strVersion(0))
            lngMin = Int(a_strVersion(1))
            lngRev = Int(a_strVersion(2))
            blnBeta = (a_strVersion(3) = "B")
            If (UBound(a_strResponse) = 2) And (UBound(a_strVersion) = 3) And (Left$(UCase$(strDownloadURL), 4) = "HTTP") Then
                'save host to ini file if changed
                If UCase$(strVCheckHost) <> UCase$(Trim$(g_xobjFunc.INI_GetSetting("Startup", "VersionCheckHostPort", "", g_strGameBaseDotINI))) Then
                    Call g_xobjFunc.INI_WriteSetting("Startup", "VersionCheckHostPort", strVCheckHost, g_strGameBaseDotINI)
                End If
                'version check
                blnNewVersion = False
                If lngMaj > App.Major Then
                    blnNewVersion = True
                ElseIf lngMaj = App.Major Then
                    If lngMin > App.Minor Then
                        blnNewVersion = True
                    ElseIf lngMin = App.Minor Then
                        If lngRev > App.Revision Then
                            blnNewVersion = True
                        ElseIf lngRev = App.Revision Then
                            If (blnBeta = False) And (g_blnBETA = True) Then
                                blnNewVersion = True
                            End If
                        End If
                    End If
                End If
                If blnNewVersion = True Then
                    'new version available
                    If MsgBox("A new version of GameBase is available.  Please check the GameBase web site for more information." & vbCrLf & vbCrLf & "Would you like to visit the web site now?", vbQuestion Or vbYesNo, gc_strAPPNAME) = vbYes Then
                        g_xobjFunc.ShellExecuteFile strDownloadURL, "", False
                    End If
                End If
            End If
        End If
    End If
    
    Unload Me
    Exit Sub

Err_Handler:
    Socket.Close
    Unload Me

End Sub


Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Socket.Close
    Unload Me

End Sub


