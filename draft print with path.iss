[Setup]
AppName=Print With Path Tool
AppVersion=1.0
AppPublisher=Aashir Mehmood
LicenseFile=license.txt
InfoBeforeFile=intro.txt
WizardStyle=modern
DefaultDirName={autopf}\PrintWithPath
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=PrintWithPathInstaller
Compression=lzma
SolidCompression=yes

[Files]
Source: "PrintWithPath.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion

[Code]
var
  DevLabel: TNewStaticText;

{ --- Function to check if Word is running --- }
function IsWordRunning(): Boolean;
begin
  Result := FindWindowByClassName('OpusApp') <> 0;
end;

{ --- Check for Word before Installation starts --- }
function InitializeSetup(): Boolean;
begin
  Result := True;
  while IsWordRunning() do
  begin
    if MsgBox('Microsoft Word is currently running.'#13#10#13#10 +
              'Please close Word before continuing with the installation.',
              mbError, MB_OKCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

{ --- Check for Word before Uninstallation starts --- }
function InitializeUninstall(): Boolean;
begin
  Result := True;
  while IsWordRunning() do
  begin
    if MsgBox('Microsoft Word is running.'#13#10 +
              'Please close it to continue uninstall.'#13#10#13#10 +
              'Click OK after closing Word.',
              mbInformation, MB_OKCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

{ --- Create the Developer Credit Label --- }
procedure InitializeWizard();
begin
  DevLabel := TNewStaticText.Create(WizardForm);
  DevLabel.Parent := WizardForm;
  DevLabel.Caption := 'Developed by: Aashir Mehmood | Version 1.0';
  
  { Position it at the bottom left }
  DevLabel.Left := ScaleX(10);
  DevLabel.Top := WizardForm.ClientHeight - ScaleY(24);
  
  { Set visibility to Black and Bold }
  DevLabel.Font.Color := clBlack;
  DevLabel.Font.Style := [fsBold];
  DevLabel.Enabled := True; 
end;

{ --- Final Messages --- }
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('Installation complete!'#13#10 +
           'Please open Microsoft Word to start using the feature.',
           mbInformation, MB_OK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    MsgBox('After uninstall, please open Word once and press Ctrl+P. If it does not work, restart Word.', mbInformation, MB_OK);
  end;
end;

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\PrintWithPath.dotm"