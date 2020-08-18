; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "3PM"
#define MyAppVerName "3PM Kabalin 7.0.x"
#define MyAppPublisher "tbrSoft Internacional"
#define MyAppURL "http://www.tbrsoft.com"
#define MyAppExeName "3PM.EXE"

[Setup]
AppName={#MyAppName}
AppVerName={#MyAppVerName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
;DisableDirPage=yes
DefaultGroupName=tbrSoft
AllowNoIcons=no
OutputBaseFilename=Instalar3PM7
SetupIconFile=D:\dev\3PM kundera 70000\3pm.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "eng"; MessagesFile: "compiler:Default.isl"
Name: "bra"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "cat"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "cze"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "dan"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dut"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "fin"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "fre"; MessagesFile: "compiler:Languages\French.isl"
Name: "ger"; MessagesFile: "compiler:Languages\German.isl"
Name: "hun"; MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "ita"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "nor"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "pol"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "por"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "rus"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slo"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spa"; MessagesFile: "compiler:Languages\SpanishArg.isl"
Name: "spa"; MessagesFile: "compiler:Languages\SpanishMex.isl"
Name: "spa"; MessagesFile: "compiler:Languages\SpanishEsp.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "D:\dev\3PM kundera 70000\pub\LEER.txt"; DestDir: "{app}\PUB"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\pubMute\LEER.txt"; DestDir: "{app}\PUBMUTE"; Flags: ignoreversion

Source: "D:\dev\3PM kundera 70000\skin\3pmBaseSkin.SKIN"; DestDir: "{app}\skin"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\skin\crystal front.SKIN"; DestDir: "{app}\skin"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\skin\wood shell.SKIN"; DestDir: "{app}\skin"; Flags: ignoreversion

Source: "D:\dev\3PM kundera 70000\CREARSKIN.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\Nev.man"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "D:\dev\3PM kundera 70000\manual.doc"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\repair 3PM\repair.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\ini.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\dev\3PM kundera 70000\SKS-DLL3\Test_Teclado_TBR.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\wbemdisp.tlb"; DestDir: "{sys}\Wbem"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrerr.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrreg.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrtimer.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrfocus.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrplayer.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrSoftVumetro.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrListaRep.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrSKS3.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrjuse.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrnfo.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrFullPak.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrcaescrypto.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrprogress.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\inpout32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrnes.dll";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrun.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrX_Boton II.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "D:\dev\3PM kundera 70000\3PM.EXE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}\3PM"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName}\Manual"; Filename: "{app}\manual.doc"
Name: "{group}\{#MyAppName}\Reparar 3PM"; Filename: "{app}\repair.exe"
Name: "{group}\{#MyAppName}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

