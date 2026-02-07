; GameBase v1.3 FULL Inno Setup Installation Script File
; (c) 2010 James Burrows

[Setup]

AppName=GameBase
AppVerName=GameBase v1.3
AppPublisher=BU22
AppPublisherURL=http://www.bu22.com/
AppSupportURL=http://www.bu22.com/
AppUpdatesURL=http://www.bu22.com/
OutputBaseFilename=gamebase-setup-v1.3-full
DefaultDirName=C:\GameBase
DefaultGroupName=BU22
AllowNoIcons=yes
LicenseFile=app\License.txt
PrivilegesRequired=admin
InfoBeforeFile=app\install-notes-for-user.txt

[Tasks]

Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; MinVersion: 4,4
Name: "quicklaunchicon"; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; MinVersion: 4,4; Flags: unchecked

[Files]

;VB6 Runtimes
Source: "vb6runtimes\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vb6runtimes\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vb6runtimes\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vb6runtimes\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vb6runtimes\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vb6runtimes\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver

;Windows System Folder (NON Register)
Source: "winsys\expsrv.dll"; DestDir: "{sys}"
Source: "winsys\msjint40.dll"; DestDir: "{sys}"
Source: "winsys\msjter40.dll"; DestDir: "{sys}"
Source: "winsys\msrepl40.dll"; DestDir: "{sys}"
Source: "winsys\mswdat10.dll"; DestDir: "{sys}"
Source: "winsys\mswstr10.dll"; DestDir: "{sys}"
Source: "winsys\VB5DB.DLL"; DestDir: "{sys}"
Source: "winsys\vbajet32.dll"; DestDir: "{sys}"

;Windows System Folder (DLL Register)
Source: "winsysreg\COMDLG32.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\TABCTL32.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\msjtes40.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\msrd2x40.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\msrd3x40.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\msjet40.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "winsysreg\mswinsck.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

;DAO Folder (DLL Register)
Source: "dao\dao360.dll"; DestDir: "{dao}"; Flags: restartreplace sharedfile regserver

;App Folder (DLL Register)
Source: "app\GBFuncs.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion
Source: "app\GBSql.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion
Source: "app\Gemus.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion

;App Folder (NON Register)
Source: "app\7z.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\7z.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\BMP2GIF.EXE"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\BMP2PNG.EXE"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\Empty.mdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GameBase.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GBLauncher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GBRegRun.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GBRegRun.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GBText.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GBToolbox.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\GemusText.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\History.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\License-7-Zip.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\License.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\PNGlib.dll"; DestDir: "{app}"; Flags: ignoreversion

;scripts folder
Source: "app\Scripts\default.txt"; DestDir: "{app}\Scripts"; Flags: ignoreversion

[Registry]

Root: HKCU; Subkey: "Software\GB64"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\GB64\GameBase"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\GB64\Gemus"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\GB64\GBToolbox"; Flags: uninsdeletekey

[Icons]

Name: "{group}\The GameBase Frontend"; Filename: "{app}\GameBase.exe"
Name: "{group}\The GameBase Toolbox"; Filename: "{app}\GBToolbox.exe"
Name: "{group}\GameBase Documentation"; Filename: "{app}\Docs\index.htm"
Name: "{userdesktop}\GameBase"; Filename: "{app}\GameBase.exe"; MinVersion: 4,4; Tasks: desktopicon
Name: "{userdesktop}\GBToolbox"; Filename: "{app}\GBToolbox.exe"; MinVersion: 4,4; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\GameBase"; Filename: "{app}\GameBase.exe"; MinVersion: 4,4; Tasks: quicklaunchicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\GBToolbox"; Filename: "{app}\GBToolbox.exe"; MinVersion: 4,4; Tasks: quicklaunchicon


