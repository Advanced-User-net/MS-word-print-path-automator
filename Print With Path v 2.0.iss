; =========================================================================
; INNO SETUP INSTALLER SCRIPT (.iss)
; Compiles a lightweight Windows setup executable that automatically registers and
; deploys the Word Template (.dotm) directly in MS Word Startup folders.
; =========================================================================

[Setup]
AppId={{52B5A254-BDFB-4A0D-AA49-1092BDEF4D75}}
AppName=MS Word Print Path Automator
AppVersion=2.0
AppPublisher=Advanceduser.net
AppPublisherURL=https://advanceduser.net
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
DisableDirPage=yes
DefaultGroupName=MS Word Print Path Automator
DisableProgramGroupPage=yes
InfoBeforeFile=
InfoAfterFile=
OutputDir=SetupBuild
OutputBaseFilename=MSWordPrintPathSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
CreateAppDir=yes
; Icon address setup (Place 'setup.ico' in the same folder as this script before compiling)
SetupIconFile=setup.ico

[Files]
; Deploy the template containing ribbon XML + VBA macro safely in the Startup directory
Source: "PrintPathAutomator.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion restartreplace

[InstallDelete]
; Delete any overlapping or conflict copies of the template before copying the new one
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\PrintPathAutomator*.dotm"

[UninstallDelete]
; Completely wipe out the template and any potential copies or backups on uninstallation
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\PrintPathAutomator*.dotm"

[Code]
// Helper to monitor running tasks of MS Word in Windows
function IsWordRunning(): Boolean;
var
  FSWCmd: String;
  ResultCode: Integer;
begin
  // First check: rapid UI window lookup using Microsoft Word's window class name 'OpusApp'
  if FindWindowByClassName('OpusApp') <> 0 then
  begin
    Result := True;
  end
  else
  begin
    // Second check: fallback standard tasklist check in case it's a headless server process
    FSWCmd := '/C tasklist /FI "IMAGENAME eq winword.exe" | find /I "winword.exe"';
    Result := Exec('cmd.exe', FSWCmd, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
  end;
end;

// Advanced Pascal Scripting to verify if Microsoft Word is running before installation
function InitializeSetup(): Boolean;
begin
  Result := True;
  
  // Checks if Word process is running to avoid locks during template copying
  while IsWordRunning() do
  begin
    if MsgBox('Microsoft Word is currently running on your system.' #13#10 #13#10 +
      'Please close all open Word documents and windows before continuing with the installation ' +
      'so that the MS Word Startup folder can be unlocked.', 
      mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

// Advanced Pascal Scripting to verify if Microsoft Word is running before uninstallation
// This prevents in-use locks that cause "some files could be removed manually" failures
function InitializeUninstall(): Boolean;
begin
  Result := True;
  
  // Checks if Word process is running to avoid locks during template uninstallation/removal
  while IsWordRunning() do
  begin
    if MsgBox('Microsoft Word is currently running on your system.' #13#10 #13#10 +
      'Please close all open Word documents and windows before continuing with the uninstallation ' +
      'so that the Startup template can be cleanly and completely removed.', 
      mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;
