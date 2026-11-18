; مثبّت برنامج جرد المطعم. البيانات تنحفظ بمجلد المستندات (مو بمجلد البرنامج) فما تنمسح بالتحديث أو إعادة التثبيت.
[Setup]
AppId={{6F1C2B0A-8D4E-4B7A-9C55-3A1E0D2F7B11}
AppName=جرد المطعم
AppVersion=1.9.2
AppPublisher=Restaurant Inventory
DefaultDirName={autopf}\RestaurantInventory
DefaultGroupName=جرد المطعم
OutputDir=Output
OutputBaseFilename=RestaurantInventory-Setup
Compression=lzma2
SolidCompression=yes
SetupIconFile=logo.ico
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
UninstallDisplayName=جرد المطعم
CloseApplications=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "إنشاء اختصار على سطح المكتب"; Flags: checkedonce

[Files]
Source: "..\dist\RestaurantInventory.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\جرد المطعم"; Filename: "{app}\RestaurantInventory.exe"; Tasks: desktopicon
Name: "{autoprograms}\جرد المطعم"; Filename: "{app}\RestaurantInventory.exe"

[Run]
Filename: "{app}\RestaurantInventory.exe"; Description: "تشغيل جرد المطعم الآن"; Flags: postinstall nowait skipifsilent
Filename: "{app}\RestaurantInventory.exe"; Flags: nowait; Check: Relaunch

[Code]
function Relaunch: Boolean;
begin
  Result := ExpandConstant('{param:RELAUNCH|0}') = '1';
end;
