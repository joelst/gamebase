Attribute VB_Name = "modUnzip32"
Option Explicit

'Global (nothing to do with unzipping)
Public g_strDefaultFile As String
Public g_blnZipError As Boolean

' argv
Private Type ZIPnames
    s(0 To 99) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(32800) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(256) As Byte
End Type

' DCL structure
Private Type DCLIST
    ExtractOnlyNewer As Long
    SpaceToUnderscore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    fPrivilege As Long
    Zip As String
    ExtractDir As String
End Type

' Userfunctions structure
Private Type USERFUNCTION
    DllPrnt As Long
    DLLSND As Long
    DLLREPLACE As Long
    DLLPASSWORD As Long
    DLLMESSAGE As Long
    cchComment As Integer
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumMembers As Long
End Type

' Unzip32.dll version structure
Private Type UZPVER
    structlen As Long
    flag As Long
    beta As String * 10
    date As String * 20
    zlib As String * 10
    UnZip(1 To 4) As Byte
    zipinfo(1 To 4) As Byte
    os2dll As Long
    windll(1 To 4) As Byte
End Type

' This assumes unzip32.dll is in
' your \windows\system directory!
Private Declare Function windll_unzip Lib "unzip32.dll" _
    (ByVal ifnc As Long, ByRef ifnv As ZIPnames, _
     ByVal xfnc As Long, ByRef xfnv As ZIPnames, _
     dcll As DCLIST, Userf As USERFUNCTION) As Long

' This assumes unzip32.dll is in
' your \windows\system directory!
Private Declare Sub UzpVersion2 Lib "unzip32.dll" _
    (uzpv As UZPVER)

' Private structures
Dim MYDCL As DCLIST
Dim MYUSER As USERFUNCTION
Dim MYVER As UZPVER

Global vbzipnum As Long, vbzipmes As String
Global vbzipinf As String
Global vbzipnam As ZIPnames, vbxnames As ZIPnames
Global crlf$


Function FnPtr(ByVal lngPointer_in As Long) As Long

    'Puts a function pointer in a structure
    FnPtr = lngPointer_in

End Function

Public Sub ReceiveDllMessage(ByVal lngUCSize_in As Long, ByVal lngCSize_in As Long, ByVal intCFactor_in As Integer, ByVal intMonth_in As Integer, ByVal intDay_in As Integer, ByVal intYear_in As Integer, ByVal intHour_in As Integer, ByVal intMinute_in As Integer, ByVal bytC_in As Byte, ByRef udtFName_inout As CBCh, ByRef udtMeth_inout As CBCh, ByVal lngCRC_in As Long, ByVal bytFCrypt_in As Byte)

    'Callback for unzip32.dll

    Dim s0$, xx As Long
    Dim strOut As String * 80

    On Error Resume Next
    strOut = Space(80)
    If vbzipnum = 0 Then
        Mid$(strOut, 1, 50) = "Filename:"
        Mid$(strOut, 53, 4) = "Size"
        Mid$(strOut, 62, 4) = "Date"
        Mid$(strOut, 71, 4) = "Time"
        vbzipmes = strOut + crlf
        strOut = Space(80)
    End If
    s0 = ""
    For xx = 0 To 255
        If udtFName_inout.ch(xx) = 0 Then xx = 99999 Else s0 = s0 & Chr$(udtFName_inout.ch(xx))
    Next xx
    Mid$(strOut, 1, 50) = Mid$(s0, 1, 50)
    Mid$(strOut, 51, 7) = Right$("        " + Str$(lngUCSize_in), 7)
    Mid$(strOut, 60, 3) = Right$(Str$(intDay_in), 2) + "/"
    Mid$(strOut, 63, 3) = Right$("0" + Trim$(Str$(intMonth_in)), 2) + "/"
    Mid$(strOut, 66, 2) = Right$("0" + Trim$(Str$(intYear_in)), 2)
    Mid$(strOut, 70, 3) = Right$(Str$(intHour_in), 2) + ":"
    Mid$(strOut, 73, 2) = Right$("0" + Trim$(Str$(intMinute_in)), 2)
    vbzipmes = vbzipmes + strOut + crlf
    vbzipnum = vbzipnum + 1
    
End Sub
' Callback for unzip32.dll
Function DllPrnt(ByRef fname As CBChar, ByVal X As Long) As Long
    Dim s0$, xx As Long

    ' always put this in callback routines!
    On Error Resume Next
    s0 = ""
    For xx = 0 To X
        If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
    Next xx
    vbzipinf = vbzipinf + s0
    DllPrnt = 0
End Function

' Callback for unzip32.dll
Function DllPass(ByRef s1 As Byte, X As Long, _
    ByRef s2 As Byte, _
    ByRef s3 As Byte) As Long

    ' always put this in callback routines!
    On Error Resume Next
    ' not supported - always return 1
    DllPass = 1
End Function

' Callback for unzip32.dll
Function DllRep(ByRef fname As CBChar) As Long
    Dim s0$, xx As Long

    ' always put this in callback routines!
    On Error Resume Next
    DllRep = 100 ' 100=do not overwrite - keep asking user
    s0 = ""
    For xx = 0 To 255
        If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
    Next xx
    xx = MsgBox("Overwrite " + s0 + "?", vbYesNoCancel, "VBUnzip - File already exists")
    If xx = vbNo Then Exit Function
    If xx = vbCancel Then
        DllRep = 104 ' 104=overwrite none
        Exit Function
    End If
    DllRep = 102 ' 102=overwrite 103=overwrite all
End Function

' ASCIIZ to String
Function szTrim(szString As String) As String
    Dim pos As Integer, ln As Integer

    pos = InStr(szString, Chr$(0))
    ln = Len(szString)
    Select Case pos
        Case Is > 1
            szTrim = Trim(Left(szString, pos - 1))
        Case 1
            szTrim = ""
        Case Else
            szTrim = Trim(szString)
    End Select
End Function

Public Sub UnzipToTempDIR(ByVal DestPath As String, ByVal ZipFileToUnzip As String)
    
    Dim prom%, over%, mess%, dirs%, numf&, numx&
    Dim ZipFile$, unzipdir$

    ' init global message variables
    vbzipinf = ""
    vbzipnum = 0
    ' select unzip options - change as required!
    prom = 1  ' 1=prompt to overwrite
    over = 1  ' 1=always overwrite files
    ' change the next line to do the actual unzip!
    mess = 0  ' 1=list contents of zip 0=extract
    dirs = 0  ' 1=honour zip directories
    ' select filenames if required
    '   vbzipnam.s(0) = "sfx16.dat"
    '   vbzipnam.s(1) = "sfx32.dat"
    '   vbzipnam.s(2) = "windll.h"
    '   numf = 3
    ' or just select all files
    vbzipnam.s(0) = vbNullString
    numf = 0
    ' select filenames to exclude from processing
    vbxnames.s(0) = vbNullString
    numx = 0
    ' Change the next 2 lines as required!
    ZipFile = ZipFileToUnzip
    unzipdir = DestPath
    '
    Call VBUnzip(ZipFile, unzipdir, _
        prom, over, mess, dirs, numf, numx)
    If Len(vbzipinf) > 0 Then
     '   Print "vbzipinf is:"
     '   Print vbzipinf
    End If
   ' If vbzipnum > 0 Then Print "Number of files: " + Str$(vbzipnum)
End Sub
' Main subroutine
Sub VBUnzip(fname As String, extdir As String, _
    prom As Integer, ovr As Integer, _
    mess As Integer, dirs As Integer, numfiles As Long, numxfiles As Long)
    Dim xx As Long ' , s1 As String * 20, s2 As String * 256

'    ChDrive Mid$(extdir, 1, 1)
'    ChDir extdir
    '
    MYDCL.ExtractOnlyNewer = 0      ' 1=extract only newer
    MYDCL.SpaceToUnderscore = 0     ' 1=convert space to underscore
    MYDCL.PromptToOverwrite = prom  ' 1=prompt to overwrite required
    MYDCL.fQuiet = 2                ' 2=no messages 1=less 0=all
    MYDCL.ncflag = 0                ' 1=write to stdout
    MYDCL.ntflag = 0                ' 1=test zip
    MYDCL.nvflag = mess             ' 0=extract 1=list contents
    MYDCL.nUflag = 0                ' 1=extract only newer
    MYDCL.nzflag = 0                ' 1=display zip file comment
    MYDCL.ndflag = dirs             ' 1=honour directories
    MYDCL.noflag = ovr              ' 1=overwrite files
    MYDCL.naflag = 0                ' 1=convert CR to CRLF
    MYDCL.nZIflag = 0               ' 1=Zip Info Verbose
    MYDCL.C_flag = 1                ' 1=Case insensitivity, 0=Case Sensitivity
    MYDCL.fPrivilege = 0            ' 1=ACL 2=priv
    MYDCL.Zip = fname               ' ZIP name
    MYDCL.ExtractDir = extdir       ' Extraction directory, NULL if extracting
                                    ' to current directory
    '
    MYUSER.DllPrnt = FnPtr(AddressOf DllPrnt)
    MYUSER.DLLSND = 0&
    MYUSER.DLLREPLACE = FnPtr(AddressOf DllRep)
    MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
    MYUSER.DLLMESSAGE = FnPtr(AddressOf ReceiveDllMessage)
    '
    With MYVER
        .structlen = Len(MYVER)
        .beta = Space$(9) & vbNullChar
        .date = Space$(19) & vbNullChar
        .zlib = Space$(9) & vbNullChar
    End With
    '
    Call UzpVersion2(MYVER)
    '
    xx = windll_unzip(numfiles, vbzipnam, numxfiles, vbxnames, MYDCL, MYUSER)
    If xx <> 0 Then g_blnZipError = True

End Sub
