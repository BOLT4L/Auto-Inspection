[Setup]
AppName=MM auto inspection
AppVersion=1.0
DefaultDirName={pf}\MM auto inspection
DefaultGroupName=MM auto inspection
OutputDir=.
OutputBaseFilename=MM_Auto_Inspection_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=vehicle-inspection.ico
UserInfoPage=yes

[Files]
Source: "dist\\main.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\\MM auto inspection"; Filename: "{app}\\main.exe"
Name: "{group}\\Uninstall MM auto inspection"; Filename: "{uninstallexe}"

[Code]
// No serial number check 