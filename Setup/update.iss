; GameBase v1.3 Inno Setup Installation Script File
; (c) 2010 James Burrows
; excludes DAO 3.6, Jet 4.0 and vb runtimes
; updates from all previous versions

[Setup]

AppName=GameBase
AppVerName=GameBase v1.3
AppPublisher=BU22
AppPublisherURL=http://www.bu22.com/
AppSupportURL=http://www.bu22.com/
AppUpdatesURL=http://www.bu22.com/
OutputBaseFilename=gamebase-setup-v1.3-update
DefaultDirName={pf}\GameBase
DefaultGroupName=BU22
AllowNoIcons=yes
LicenseFile=app\License.txt
PrivilegesRequired=admin
InfoBeforeFile=app\install-notes-for-user.txt
Uninstallable=no


[Files]
;Windows System Folder (DLL Register)
Source: "winsysreg\mswinsck.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

;App Folder (DLL Register)
Source: "app\GBFuncs.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion
Source: "app\GBSql.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion
Source: "app\Gemus.dll"; DestDir: "{app}"; Flags: sharedfile regserver ignoreversion

;App Folder (NON Register)
Source: "app\7z.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "app\7z.exe"; DestDir: "{app}"; Flags: ignoreversion
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

[Run]
Filename: "{app}\GameBase.exe"; Description: "Run GameBase v1.3 now"; Flags: postinstall


