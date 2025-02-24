{ Copyright (C) 2024-2025 by Bill Stewart (bstewart at iname.com)

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

unit Utility;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

interface

uses
  windows;

type
  TWindowStyle = (
    Hidden             = SW_HIDE,
    Normal             = SW_SHOWNORMAL,
    Minimized          = SW_SHOWMINIMIZED,
    Maximized          = SW_SHOWMAXIMIZED,
    NormalNotActive    = SW_SHOWNOACTIVATE,
    MinimizedNotActive = SW_SHOWMINNOACTIVE);

function AppendStr(const S1, S2, Delim: string): string;

function IntToStr(const D: DWORD): string;

function IsVistaOrLater(): Boolean;

function GetFileVersion(const FileName: string): string;

function FileExists(const FileName: string): Boolean;

function DirExists(const DirName: string): Boolean;

function FindFileInPath(const Path, FileName: string): string;

function GetParentPath(const Path: string): string;

function ShellExec(const Executable, Parameters, WorkingDirectory: string;
  const WindowStyle: TWindowStyle; const Wait, Quiet, Elevate: Boolean;
  var ResultCode: DWORD): Boolean;

function ReplaceStr(S: string; const Fnd, Repl: string): string;

implementation

const
  INVALID_FILE_ATTRIBUTES    = DWORD(-1);
  CRYPT_STRING_BASE64        = $00000001;
  CRYPT_STRING_NOCRLF        = $40000000;
  SEE_MASK_DEFAULT           = $00000000;
  SEE_MASK_CLASSNAME         = $00000001;
  SEE_MASK_CLASSKEY          = $00000003;
  SEE_MASK_IDLIST            = $00000004;
  SEE_MASK_INVOKEIDLIST      = $0000000C;
  SEE_MASK_ICON              = $00000010;
  SEE_MASK_HOTKEY            = $00000020;
  SEE_MASK_NOCLOSEPROCESS    = $00000040;
  SEE_MASK_CONNECTNETDRV     = $00000080;
  SEE_MASK_NOASYNC           = $00000100;
  SEE_MASK_FLAG_DDEWAIT      = $00000100;
  SEE_MASK_DOENVSUBST        = $00000200;
  SEE_MASK_FLAG_NO_UI        = $00000400;
  SEE_MASK_UNICODE           = $00004000;
  SEE_MASK_NO_CONSOLE        = $00008000;
  SEE_MASK_ASYNCOK           = $00100000;
  SEE_MASK_NOQUERYCLASSSTORE = $01000000;
  SEE_MASK_HMONITOR          = $00200000;
  SEE_MASK_NOZONECHECKS      = $00800000;
  SEE_MASK_WAITFORINPUTIDLE  = $02000000;
  SEE_MASK_FLAG_LOG_USAGE    = $04000000;

type
  TShellExecuteInfo = record
    cbSize:       DWORD;
    fMask:        ULONG;
    hwnd:         ULONG;
    lpVerb:       LPCWSTR;
    lpFile:       LPCWSTR;
    lpParameters: LPCWSTR;
    lpDirectory:  LPCWSTR;
    nShow:        Integer;
    hInstApp:     HINST;
    lpIDList:     LPVOID;
    lpClass:      LPCWSTR;
    hKeyClass:    HKEY;
    dwHotKey:     DWORD;
    hMonitor:     HANDLE;
    hProcess:     HANDLE;
  end;
  TStringArray = array of string;

function ShellExecuteExW(var ShellExecuteInfo: TShellExecuteInfo): BOOL; stdcall;
  external 'shell32.dll';

function AppendStr(const S1, S2, Delim: string): string;
begin
  if S1 = '' then
    result := S2
  else
  begin
    if S1[Length(S1)] <> Delim then
      result := S1 + Delim + S2
    else
      result := S1 + S2;
  end;
end;

function IsVistaOrLater(): Boolean;
var
  OVI: OSVERSIONINFO;
begin
  result := false;
  OVI.dwOSVersionInfoSize := SizeOf(OSVERSIONINFO);
  if GetVersionEx(OVI) then
    result := (OVI.dwPlatformId = VER_PLATFORM_WIN32_NT) and (OVI.dwMajorVersion >= 6);
end;

function IntToStr(const D: DWORD): string;
begin
  Str(D, result);
end;

function GetFileVersion(const FileName: string): string;
var
  VerInfoSize, Handle: DWORD;
  pBuffer: Pointer;
  pFileInfo: ^VS_FIXEDFILEINFO;
  Len: UINT;
begin
  result := '';
  VerInfoSize := GetFileVersionInfoSizeW(PChar(FileName),  // LPCWSTR lptstrFilename
    Handle);                                               // LPDWORD lpdwHandle
  if VerInfoSize > 0 then
  begin
    GetMem(pBuffer, VerInfoSize);
    if GetFileVersionInfoW(PChar(FileName),  // LPCWSTR lptstrFilename
      Handle,                                // DWORD   dwHandle
      VerInfoSize,                           // DWORD   dwLen
      pBuffer) then                          // LPVOID  lpData
    begin
      if VerQueryValueW(pBuffer,  // LPCVOID pBlock
        '\',                      // LPCWSTR lpSubBlock
        pFileInfo,                // LPVOID  *lplpBuffer
        Len) then                 // PUINT   puLen
      begin
        with pFileInfo^ do
        begin
          result := IntToStr(HiWord(dwFileVersionMS)) + '.' +
            IntToStr(LoWord(dwFileVersionMS)) + '.' +
            IntToStr(HiWord(dwFileVersionLS));
        end;
      end;
    end;
    FreeMem(pBuffer);
  end;
end;

function FileExists(const FileName: string): Boolean;
var
  Attrs: DWORD;
begin
  Attrs := GetFileAttributesW(PChar(FileName));  // LPCWSTR lpFileName
  result := (Attrs <> INVALID_FILE_ATTRIBUTES) and
    ((Attrs and FILE_ATTRIBUTE_DIRECTORY) = 0);
end;

function DirExists(const DirName: string): Boolean;
var
  Attrs: DWORD;
begin
  Attrs := GetFileAttributesW(PChar(DirName));  // LPCWSTR lpFileName
  result := (Attrs <> INVALID_FILE_ATTRIBUTES) and
    ((Attrs and FILE_ATTRIBUTE_DIRECTORY) <> 0);
end;

function FindFileInPath(const Path, FileName: string): string;
var
  pPath, pFileName: PChar;
  NumChars, BufSize: DWORD;
begin
  result := '';
  if Path = '' then
    pPath := nil
  else
    pPath := PChar(Path);
  NumChars := SearchPathW(pPath,  // LPCWSTR lpPath
    PChar(FileName),              // LPCWSTR lpFileName
    nil,                          // LPCWSTR lpExtension
    0,                            // DWORD   nBufferLength
    nil,                          // LPWSTR  lpBuffer
    nil);                         // LPWSTR  lpFilePart
  if NumChars > 0 then
  begin
    BufSize := NumChars * SizeOf(Char);
    GetMem(pFileName, BufSize);
    if SearchPathW(pPath,  // LPCWSTR lpPath
      PChar(FileName),     // LPCWSTR lpFileName
      nil,                 // LPCWSTR lpExtension
      NumChars,            // DWORD   nBufferLength
      pFileName,           // LPWSTR  lpBuffer
      nil) > 0 then        // LPWSTR  lpFilePart
    begin
      result := pFileName;
    end;
    FreeMem(pFileName);
  end;
end;

function GetParentPath(const Path: string): string;
var
  L, I: Integer;
begin
  result := Path;
  L := Length(Path);
  if L > 1 then
  begin
    for I := L downto 1 do
      if Path[I] = '\' then
      begin
        result := Copy(Path, 1, I - 1);
        break;
      end;
  end;
end;

function ShellExec(const Executable, Parameters, WorkingDirectory: string;
  const WindowStyle: TWindowStyle; const Wait, Quiet, Elevate: Boolean;
  var ResultCode: DWORD): Boolean;
var
  SEI: TShellExecuteInfo;
begin
  FillChar(SEI, SizeOf(SEI), 0);
  SEI.cbSize := SizeOf(SEI);
  if Wait then
    SEI.fMask := SEI.fMask or SEE_MASK_NOCLOSEPROCESS;
  if Quiet then
    SEI.fMask := SEI.fMask or SEE_MASK_FLAG_NO_UI;
  if Elevate then
    SEI.lpVerb := 'runas'
  else
    SEI.lpVerb := 'open';
  SEI.lpFile := PChar(Executable);
  if Parameters <> '' then
    SEI.lpParameters := PChar(Parameters)
  else
    SEI.lpParameters := nil;
  if WorkingDirectory <> '' then
    SEI.lpDirectory := PChar(WorkingDirectory)
  else
    SEI.lpDirectory := nil;
  SEI.nShow := Integer(WindowStyle);
  result := ShellExecuteExW(SEI);  // SEIW *pExecInfo
  if result then
  begin
    if Wait then
    begin
      result := WaitForSingleObject(SEI.hProcess,  // HANDLE hHandle
        INFINITE) <> WAIT_FAILED;                  // DWORD  dwMilliseconds
      if result then
      begin
        result := GetExitCodeProcess(SEI.hProcess,  // HANDLE  hProcess
          ResultCode);                              // LPDWORD lpExitCode
        if not result then
          ResultCode := GetLastError();
      end
      else
        ResultCode := GetLastError();
    end
    else
      ResultCode := 0;
    CloseHandle(SEI.hProcess);  // HANDLE hObject
  end
  else
    ResultCode := GetLastError();
end;

// Returns the number of times Substring appears in S
function CountSubstring(const Substring, S: string): Integer;
var
  P: Integer;
begin
  result := 0;
  P := Pos(Substring, S, 1);
  while P <> 0 do
  begin
    Inc(result);
    P := Pos(Substring, S, P + Length(Substring));
  end;
end;

// Splits S into the Dest array using Delim as a delimiter
procedure StrSplit(S, Delim: string; var Dest: TStringArray);
var
  I, P: Integer;
begin
  I := CountSubstring(Delim, S);
  // If no delimiters, then Dest is a single-element array
  if I = 0 then
  begin
    SetLength(Dest, 1);
    Dest[0] := S;
    exit;
  end;
  SetLength(Dest, I + 1);
  for I := 0 to Length(Dest) - 1 do
  begin
    P := Pos(Delim, S);
    if P > 0 then
    begin
      Dest[I] := Copy(S, 1, P - 1);
      Delete(S, 1, P + Length(Delim) - 1);
    end
    else
      Dest[I] := S;
  end;
end;

function ReplaceStr(S: string; const Fnd, Repl: string): string;
var
  A: TStringArray;
  I: Integer;
begin
  StrSplit(S, Fnd, A);
  if Length(A) = 1 then
  begin
    result := S;
    exit;
  end;
  S := A[0];
  for I := 1 to Length(A) - 1 do
  begin
    S := S + Repl + A[I];
  end;
  result := S;
end;

begin
end.
