{ Copyright (C) 2021-2024 by Bill Stewart (bstewart at iname.com)

  This program is free software: you can redistribute it and/or modify it under
  the terms of the GNU General Public License as published by the Free Software
  Foundation, either version 3 of the License, or (at your option) any later
  version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
  details.

  You should have received a copy of the GNU General Public License
  along with this program. If not, see https://www.gnu.org/licenses/.

}

program startps;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}
{$R *.res}

uses
  wargcv,
  wgetopts,
  WindowsMessages,
  WindowsRegistry,
  Utility,
  windows;

const
  APP_TITLE = 'startps';
  APP_COPYRIGHT = '(C) 2021-2024 by Bill Stewart (bstewart AT iname.com)';
  PS_WIN_APPPATH_SUBKEY = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerShell.exe';
  PS_WIN_EXECUTABLE_NAME = 'powershell.exe';
  PS_CORE_APPPATH_SUBKEY = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe';
  PS_CORE_EXECUTABLE_NAME = 'pwsh.exe';

type
  TCommandLine = object
    ErrorCode: Word;
    ErrorMessage: string;
    ConfigurationName: string;
    ConsoleFileName: string;
    DebugFileName: string;
    DisableExecutionPolicy: Boolean;
    Core: Boolean;
    CorePath: string;
    Elevate: Boolean;
    Help: Boolean;
    Interactive: Boolean;
    LoadProfile: Boolean;
    Logo: Boolean;
    MTA: Boolean;
    NoExit: Boolean;
    NonInteractive: Boolean;
    NoProfile: Boolean;
    OutputFormat: string;
    Pause: Boolean;
    Quiet: Boolean;
    STA: Boolean;
    Version: string;
    Wait: Boolean;
    WindowStyle: TWindowStyle;
    WindowTitle: Boolean;
    WindowTitleText: string;
    WorkingDirectory: string;
    ScriptFileName: string;
    ScriptParameters: string;
    procedure Parse();
  end;

procedure Usage();
var
  Msg: string;
begin
  Msg := APP_TITLE + ' ' + GetFileVersion(ParamStr(0)) + ' - ' + APP_COPYRIGHT + #10
    + #10
    + 'This is free software and comes with ABSOLUTELY NO WARRANTY.' + #10
    + #10
    + 'Runs a PowerShell script or opens an interactive PowerShell window.' + #10
    + #10
    + 'Usage: ' + APP_TITLE + ' [parameters] [scriptfile [-- parameters]]' + #10
    + #10
    + 'Summary of common parameters:' + #10
    + '    --core[="path"] - Use PS Core' + #10
    + '    --disableexecutionpolicy - Disable PS execution policy' + #10
    + '    --elevate - Request to run as administrator' + #10
    + '    --quiet - Suppress error messages' + #10
    + '    --windowstyle=style - Specify a window style' + #10
    + '    --windowtitle[="text"] - Specify a window title' + #10
    + '    --workingdir="path" - Specify a working directory' + #10
    + #10
    + '--windowstyle style must be one of the following: Normal, Minimized, '
    + 'Maximized, Hidden, NormalNotActive, or MinimizedNotActive' + #10
    + #10
    + 'Parameters if running a script:' + #10
    + '    --loadprofile - Load PS profile(s) before running script' + #10
    + '    --noexit - Keep PS window open after running script' + #10
    + '    --noninteractive - Run script non-interactively' + #10
    + '    --pause - Pause window after script completes' + #10
    + '    --wait - Wait for exit and return process exit code' + #10
    + '    scriptfile - Path/filename of script to run' + #10
    + '    -- - everything after -- is script parameters' + #10
    + #10
    + 'Parameters if running interactively:' + #10
    + '  --interactive - Open interactive PS window' + #10
    + '  --noprofile - Do not load PS profile(s)' + #10
    + #10
    + 'Parameter names are case-sensitive.';
  MessageBoxW(0,  // HWND    hWnd
    PChar(Msg),   // LPCWSTR lpText
    APP_TITLE,    // LPCWSTR lpCaption
    0);           // UINT    uType
end;

procedure TCommandLine.Parse();
var
  Opts: array[1..24] of TOption;
  Opt: Char;
  I: Integer;
begin
  // Set up array of options; requires final option with empty name;
  // set Value member to specify short option match for GetLongOps
  with Opts[1] do
  begin
    Name := 'configurationname';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[2] do
  begin
    Name := 'consolefilename';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[3] do
  begin
    Name := 'core';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := 'c';
  end;
  with Opts[4] do
  begin
    Name := 'debug';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[5] do
  begin
    Name := 'disableexecutionpolicy';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'D';
  end;
  with Opts[6] do
  begin
    Name := 'elevate';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'e';
  end;
  with Opts[7] do
  begin
    Name := 'help';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'h';
  end;
  with Opts[8] do
  begin
    Name := 'interactive';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'i';
  end;
  with Opts[9] do
  begin
    Name := 'loadprofile';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[10] do
  begin
    Name := 'logo';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[11] do
  begin
    Name := 'mta';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[12] do
  begin
    Name := 'noexit';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[13] do
  begin
    Name := 'noninteractive';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'n';
  end;
  with Opts[14] do
  begin
    Name := 'noprofile';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[15] do
  begin
    Name := 'outputformat';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[16] do
  begin
    Name := 'pause';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'p';
  end;
  with Opts[17] do
  begin
    Name := 'quiet';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'q';
  end;
  with Opts[18] do
  begin
    Name := 'sta';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[19] do
  begin
    Name := 'version';
    Has_Arg := Required_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[20] do
  begin
    Name := 'wait';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'w';
  end;
  with Opts[21] do
  begin
    Name := 'windowstyle';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := 'W';
  end;
  with Opts[22] do
  begin
    Name := 'windowtitle';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := 't';
  end;
  with Opts[23] do
  begin
    Name := 'workingdirectory';
    Has_arg := Required_Argument;
    Flag := nil;
    Value := 'd';
  end;
  with Opts[24] do
  begin
    Name := '';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  // Initialize defaults
  ErrorCode := 0;
  ErrorMessage := '';
  ConfigurationName := '';
  ConsoleFileName := '';
  Core := false;
  CorePath := '';
  DebugFileName := '';
  DisableExecutionPolicy := false;
  Elevate := false;
  Help := false;
  Interactive := false;
  LoadProfile := false;
  Logo := false;
  MTA := false;
  NoExit := false;
  NonInteractive := false;
  NoProfile := false;
  OutputFormat := '';
  Pause := false;
  Quiet := false;
  STA := false;
  Version := '';
  Wait := false;
  WindowStyle := Normal;
  WindowTitle := false;
  WindowTitleText := '';
  WorkingDirectory := '';
  //OptErr := false;  // no error output from wgetopts
  repeat
    Opt := GetLongOpts('c::Dd:ehinpqt:wW::', @Opts[1], I);
    case Opt of
      'c':
      begin
        Core := true;
        if OptArg <> '' then
          CorePath := OptArg;
      end;
      'D': DisableExecutionPolicy := true;
      'd':
      begin
        if OptArg = '' then
        begin
          ErrorCode := ERROR_INVALID_PARAMETER;
          ErrorMessage := '--workingdirectory (-d) parameter requires an argument';
        end
        else
        begin
          WorkingDirectory := OptArg;
          if not DirExists(WorkingDirectory) then
          begin
            ErrorCode := ERROR_PATH_NOT_FOUND;
            ErrorMessage := 'Path not found - ''' + WorkingDirectory + '''';
          end;
        end;
      end;
      'e': Elevate := true;
      'h': Help := true;
      'i': Interactive := true;
      'n': NonInteractive := true;
      'p': Pause := true;
      'q': Quiet := true;
      'w': Wait := true;
      'W':
      begin
        case LowercaseString(OptArg) of
          'hidden': WindowStyle := Hidden;
          'normal': WindowStyle := Normal;
          'minimized': WindowStyle := Minimized;
          'maximized': WindowStyle := Maximized;
          'normalnotactive': WindowStyle := NormalNotActive;
          'minimizednotactive': WindowStyle := MinimizedNotActive;
          else
          begin
            ErrorCode := ERROR_INVALID_PARAMETER;
            ErrorMessage := '--windowstyle (-W) argument must be one of the following: ' +
              '''Hidden'', ''Normal'', ''Minimized'', ''Maximized'', ''NormalNotActive'', or ''MinimizedNotActive''';
          end;
        end;
      end;
      't':
      begin
        WindowTitle := true;
        WindowTitleText := OptArg;
      end;
      #0:
      begin
        case LowercaseString(Opts[I].Name) of
          'configurationname':
          begin
            if OptArg = '' then
            begin
              ErrorCode := ERROR_INVALID_PARAMETER;
              ErrorMessage := '--configurationname parameter requires an argument';
            end
            else
              ConfigurationName := OptArg;
          end;
          'consolefilename':
          begin
            if OptArg = '' then
            begin
              ErrorCode := ERROR_INVALID_PARAMETER;
              ErrorMessage := '--consolefilename parameter requires an argument';
            end
            else
            begin
              ConsoleFileName := OptArg;
              if not FileExists(ConsoleFileName) then
              begin
                ErrorCode := ERROR_FILE_NOT_FOUND;
                ErrorMessage := 'File not found - ''' + ConsoleFilename + '''';
              end;
            end;
          end;
          'debug':
          begin
            if OptArg <> '' then
              DebugFileName := OptArg;
          end;
          'loadprofile': LoadProfile := true;
          'logo': Logo := true;
          'mta': MTA := true;
          'noexit': NoExit := true;
          'noprofile': NoProfile := true;
          'outputformat':
          begin
            if not (SameText(OptArg, 'text') or SameText(OptArg, 'xml')) then
            begin
              ErrorCode := ERROR_INVALID_PARAMETER;
              ErrorMessage :=
                '--outputformat parameter''s argument must be one of the following: ' +
                '''Text'' or ''XML''';
            end
            else
              OutputFormat := OptArg;
          end;
          'sta': STA := true;
          'version': Version := OptArg;
        end;
      end;
      '?':
      begin
        ErrorCode := ERROR_INVALID_PARAMETER;
        ErrorMessage := 'Invalid parameter specified; use --help (-h) for usage information';
      end;
    end; //case Opt
  until Opt = EndOfOptions;
  ScriptFileName := ParamStr(OptInd);
  if ScriptFileName <> '' then
  begin
    if not FileExists(ScriptFileName) then
    begin
      ErrorCode := ERROR_FILE_NOT_FOUND;
      ErrorMessage := 'Script file not found - ''' + ScriptFileName + '''';
    end;
    ScriptParameters := GetCommandTail(GetCommandLineW(), OptInd + 1);
  end;
  if Interactive and (ScriptFileName <> '') then
  begin
    ErrorCode := ERROR_INVALID_PARAMETER;
    ErrorMessage := '--interactive (-i) and script file are mutually exclusive';
  end;
  if STA and MTA then
  begin
    ErrorCode := ERROR_INVALID_PARAMETER;
    ErrorMessage := '--mta and --sta are mutually exclusive options';
  end;
  if (not Interactive) and (ScriptFileName = '') then
  begin
    ErrorCode := ERROR_INVALID_PARAMETER;
    ErrorMessage := 'You must specify --interactive (-i) or a script file name';
  end;
end;

procedure ErrorDialog(const Msg: string; const ErrorCode: DWORD);
begin
  MessageBoxW(0,                                    // HWND    hWnd
    PChar(Msg + ' (' + IntToStr(ErrorCode) + ')'),  // LPCWSTR lpText
    APP_TITLE,                                      // LPCWSTR lpCaption
    MB_ICONERROR);                                  // UINT    uType
end;

function GetPSPath(const Core: Boolean): string;
var
  SubKeyName, FileName, Path: string;
begin
  result := '';
  if Core then
  begin
    SubKeyName := PS_CORE_APPPATH_SUBKEY;
    FileName := PS_CORE_EXECUTABLE_NAME;
  end
  else
  begin
    SubKeyName := PS_WIN_APPPATH_SUBKEY;
    FileName := PS_WIN_EXECUTABLE_NAME;
  end;
  // Try registry first
  if RegGetExpandStringValue('', HKEY_LOCAL_MACHINE, SubKeyName, '', Path) = ERROR_SUCCESS then
    result := Path
  else
  begin
    // Search Path environment variable
    Path := FindFileInPath('', FileName);
    if Path <> '' then
      result := Path;
  end;
end;

function GetConfigPolicy(const Core: Boolean): string;
var
  CfgPolicy: string;
begin
  CfgPolicy := 'function Disable-ExecutionPolicy{($c=$ExecutionContext.GetType().GetField("_context","NonPublic,Instance").GetValue($ExecutionContext)).GetType().GetField("';
  if Core then
  begin
    CfgPolicy := CfgPolicy + '<AuthorizationManager>k__BackingField';
  end
  else
  begin
    CfgPolicy := CfgPolicy + '_authorizationManager';
  end;
  result := CfgPolicy + '","NonPublic,Instance").SetValue($c,(New-Object Management.Automation.AuthorizationManager "Microsoft.PowerShell"))}';
end;

function AppendStr(const S1, S2, Delim: string): string;
begin
  if S1 = '' then
    result := S2
  else
    result := S1 + Delim + S2;
end;

var
  CommandLine: TCommandLine;   // Command line parser object
  ExecutableFileName: string;  // Path/filename of powershell.exe or pwsh.exe
  PSType: string;              // Windows PowerShell or PowerShell Core?
  Command: string;             // PowerShell command to run
  Parameters: string;          // Parameters for powershell or pwsh
  DebugFile: Text;
  DebugString: string;
  ResultCode: DWORD;

begin
  if (ParamCount = 0) or (ParamStr(1) = '/?') then
  begin
    Usage();
    exit;
  end;

  if not IsVistaOrLater() then
  begin
    ExitCode := ERROR_OLD_WIN_VERSION;
    ErrorDialog(GetWindowsMessage(ExitCode), ExitCode);
    exit;
  end;

  CommandLine.Parse();

  if CommandLine.Help then
  begin
    Usage();
    exit;
  end;

  // Fail if we got a command-line error
  ExitCode := CommandLine.ErrorCode;
  if ExitCode <> 0 then
  begin
    if not CommandLine.Quiet then
      ErrorDialog(CommandLine.ErrorMessage, CommandLine.ErrorCode);
    exit;
  end;

  // Get path/filename of pwsh.exe if specified on command line; otherwise,
  // get path/filename of pwsh.exe or powershell.exe from registry or Path
  if CommandLine.Core and (CommandLine.CorePath <> '') then
    ExecutableFileName := CommandLine.CorePath
  else
    ExecutableFileName := GetPSPath(CommandLine.Core);

  // Fail if registry and Path searches failed
  if ExecutableFileName = '' then
  begin
    if not CommandLine.Core then
      PSType := 'Windows PowerShell'
    else
      PSType := 'PowerShell Core';
    ExitCode := ERROR_FILE_NOT_FOUND;
    if not CommandLine.Quiet then
      ErrorDialog('Unable to find ' + PSType +
        ' in the registry or system Path.', ExitCode);
    exit;
  end;

  // Fail if file not found
  if not FileExists(ExecutableFileName) then
  begin
    ExitCode := ERROR_FILE_NOT_FOUND;
    if not CommandLine.Quiet then
      ErrorDialog('Unable to find file ''' + ExecutableFileName + '''.', ExitCode);
    exit;
  end;

  Command := '';

  if CommandLine.DisableExecutionPolicy then
    Command := AppendStr(AppendStr(Command, GetConfigPolicy(CommandLine.Core), ';'),
      'Disable-ExecutionPolicy', ';');

  // Set window title if requested
  if CommandLine.WindowTitle then
    Command := AppendStr(Command, '$Host.UI.RawUI.WindowTitle="' +
      ReplaceStr(CommandLine.WindowTitleText, '"', '""') + '"', ';');

  // Sanity check hidden window
  if CommandLine.WindowStyle = Hidden then
  begin
    if CommandLine.Interactive or CommandLine.Pause then
      CommandLine.WindowStyle := Normal;
  end;

  if CommandLine.ScriptFileName <> '' then
  begin
    Command := AppendStr(Command, '& "' + CommandLine.ScriptFilename + '"', ';');
    // Add parameters (if any)
    if CommandLine.ScriptParameters <> '' then
      Command := Command + ' ' + CommandLine.ScriptParameters;
    if not CommandLine.NoExit then
    begin
      // Add pause if requested
      if CommandLine.Pause then
        Command := AppendStr(Command, 'Read-Host "Press ENTER to continue"', ';');
      // Get exit code if waiting
      if CommandLine.Wait then
        Command := AppendStr(Command, 'exit $LASTEXITCODE', ';');
    end;
  end;

  // Build parameters for powershell.exe or pwsh.exe
  Parameters := '';

  // If used with Windows PowerShell, -Version parameter must be first
  if (not CommandLine.Core) and (CommandLine.Version <> '') then
    Parameters := AppendStr(Parameters, '-Version "' + CommandLine.Version + '"', ' ');

  if CommandLine.Interactive then
  begin
    if not CommandLine.Logo then
      Parameters := AppendStr(Parameters, '-NoLogo', ' ');
    if CommandLine.NoProfile then
      Parameters := AppendStr(Parameters, '-NoProfile', ' ');
    if Command <> '' then
      Parameters := AppendStr(Parameters, '-NoExit', ' ');
  end
  else
  begin
    if CommandLine.NoExit then
      Parameters := AppendStr(Parameters, '-NoExit', ' ');
    if not CommandLine.LoadProfile then
      Parameters := AppendStr(Parameters, '-NoProfile', ' ');
    if CommandLine.NonInteractive then
      Parameters := AppendStr(Parameters, '-NonInteractive', ' ');
  end;

  if CommandLine.MTA then
    Parameters := AppendStr(Parameters, '-MTA', ' ');
  if CommandLine.STA then
    Parameters := AppendStr(Parameters, '-STA', ' ');
  if CommandLine.ConfigurationName <> '' then
    Parameters := AppendStr(Parameters, '-ConfigurationName "' +
      CommandLine.ConfigurationName + '"', ' ');
  if (CommandLine.ConsoleFileName <> '') and (not CommandLine.Core) then
    Parameters := AppendStr(Parameters, '-PSConsoleFile "' +
      CommandLine.ConsoleFileName + '"', ' ');
  if CommandLine.OutputFormat <> '' then
    Parameters := AppendStr(Parameters, '-OutputFormat ' +
      CommandLine.OutputFormat, ' ');

  if Command <> '' then
    Parameters := AppendStr(Parameters, '-Command "' + ReplaceStr(Command, '"', '"""') + '"', ' ');

  {$IFDEF DEBUG}
  WriteLn();
  with CommandLine do
  begin
    WriteLn('ErrorCode:              ', ErrorCode);
    WriteLn('ErrorMessage:           ', ErrorMessage);
    WriteLn('ConfigurationName:      ', ConfigurationName);
    WriteLn('ConsoleFileName:        ', ConsoleFileName);
    WriteLn('Core:                   ', Core);
    WriteLn('CorePath:               ', CorePath);
    WriteLn('DisableExecutionPolicy: ', DisableExecutionPolicy);
    WriteLn('Elevate:                ', Elevate);
    WriteLn('Help:                   ', Help);
    WriteLn('Interactive:            ', Interactive);
    WriteLn('LoadProfile:            ', LoadProfile);
    WriteLn('Logo:                   ', Logo);
    WriteLn('MTA:                    ', MTA);
    WriteLn('NoExit:                 ', NoExit);
    WriteLn('NonInteractive:         ', NonInteractive);
    WriteLn('NoProfile:              ', NoProfile);
    WriteLn('OutputFormat:           ', OutputFormat);
    WriteLn('Pause:                  ', Pause);
    WriteLn('Quiet:                  ', Quiet);
    WriteLn('Wait:                   ', Wait);
    WriteLn('WindowTitle:            ', WindowTitle);
    WriteLn('WindowTitleText:        ', WindowTitleText);
    WriteLn('WindowStyle:            ', WindowStyle);
    WriteLn('WorkingDirectory:       ', WorkingDirectory);
    WriteLn('Version:                ', Version);
    WriteLn('ScriptFilename:         ', ScriptFilename);
    WriteLn('ScriptParameters:       ', ScriptParameters);
  end;
  WriteLn();
  WriteLn('ExecutableFileName: ', ExecutableFileName);
  WriteLn();
  WriteLn('Command: ', Command);
  WriteLn();
  WriteLn('Parameters: [', Parameters, ']');
  ResultCode := 0;
  {$ENDIF}

  if CommandLine.DebugFileName <> '' then
  begin
    Assign(DebugFile, CommandLine.DebugFileName);
    {$I-}
    if FileExists(CommandLine.DebugFileName) then
      Append(DebugFile)
    else
      Rewrite(DebugFile);
    DebugString := 'Executable: ' + ExecutableFileName + sLineBreak;
    if CommandLine.WorkingDirectory <> '' then
      DebugString := DebugString + 'Working directory: ' + CommandLine.WorkingDirectory + sLineBreak;
    DebugString := DebugString + 'Parameters: ' + Parameters;
    WriteLn(DebugFile, DebugString);
    Close(DebugFile);
    {$I+}
  end;

  {$IFNDEF DEBUG}
  if not ShellExec(ExecutableFileName,  // Executable
    Parameters,                         // Parameters
    CommandLine.WorkingDirectory,       // WorkingDirectory
    CommandLine.WindowStyle,            // WindowStyle
    CommandLine.Wait,                   // Wait
    CommandLine.Quiet,                  // Quiet
    CommandLine.Elevate,                // Elevate
    ResultCode) then                    // ResultCode
  begin
    if not CommandLine.Quiet then
      ErrorDialog(GetWindowsMessage(ResultCode), ResultCode);
  end;
  {$ENDIF}

  ExitCode := Integer(ResultCode);
end.
