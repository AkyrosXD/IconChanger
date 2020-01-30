
// Source: http://www.informit.com/articles/article.aspx?p=26940&seqNum=4

unit WinShell;

interface

uses SysUtils, Windows, Registry, ActiveX, ShlObj;

type
 EShellOleError = class(Exception);

 TShellLinkInfo = record
  PathName: string;
  Arguments: string;
  Description: string;
  WorkingDirectory: string;
  IconLocation: string;
  IconIndex: integer;
  ShowCmd: integer;
  HotKey: word;
 end;

 TSpecialFolderInfo = record
  Name: string;
  ID: Integer;
 end;

const
 SpecialFolders: array[0..29] of TSpecialFolderInfo = (
  (Name: 'Alt Startup'; ID: CSIDL_ALTSTARTUP),
  (Name: 'Application Data'; ID: CSIDL_APPDATA),
  (Name: 'Recycle Bin'; ID: CSIDL_BITBUCKET),
  (Name: 'Common Alt Startup'; ID: CSIDL_COMMON_ALTSTARTUP),
  (Name: 'Common Desktop'; ID: CSIDL_COMMON_DESKTOPDIRECTORY),
  (Name: 'Common Favorites'; ID: CSIDL_COMMON_FAVORITES),
  (Name: 'Common Programs'; ID: CSIDL_COMMON_PROGRAMS),
  (Name: 'Common Start Menu'; ID: CSIDL_COMMON_STARTMENU),
  (Name: 'Common Startup'; ID: CSIDL_COMMON_STARTUP),
  (Name: 'Controls'; ID: CSIDL_CONTROLS),
  (Name: 'Cookies'; ID: CSIDL_COOKIES),
  (Name: 'Desktop'; ID: CSIDL_DESKTOP),
  (Name: 'Desktop Directory'; ID: CSIDL_DESKTOPDIRECTORY),
  (Name: 'Drives'; ID: CSIDL_DRIVES),
  (Name: 'Favorites'; ID: CSIDL_FAVORITES),
  (Name: 'Fonts'; ID: CSIDL_FONTS),
  (Name: 'History'; ID: CSIDL_HISTORY),
  (Name: 'Internet'; ID: CSIDL_INTERNET),
  (Name: 'Internet Cache'; ID: CSIDL_INTERNET_CACHE),
  (Name: 'Network Neighborhood'; ID: CSIDL_NETHOOD),
  (Name: 'Network Top'; ID: CSIDL_NETWORK),
  (Name: 'Personal'; ID: CSIDL_PERSONAL),
  (Name: 'Printers'; ID: CSIDL_PRINTERS),
  (Name: 'Printer Links'; ID: CSIDL_PRINTHOOD),
  (Name: 'Programs'; ID: CSIDL_PROGRAMS),
  (Name: 'Recent Documents'; ID: CSIDL_RECENT),
  (Name: 'Send To'; ID: CSIDL_SENDTO),
  (Name: 'Start Menu'; ID: CSIDL_STARTMENU),
  (Name: 'Startup'; ID: CSIDL_STARTUP),
  (Name: 'Templates'; ID: CSIDL_TEMPLATES));

function CreateShellLink(const AppName, Desc: string; Dest: Integer): string;

function GetSpecialFolderPath(Folder: Integer; CanCreate: Boolean): string;

procedure GetShellLinkInfo(const LinkFile: WideString; var SLI: TShellLinkInfo);

procedure SetShellLinkInfo(const LinkFile: WideString; const SLI: TShellLinkInfo);

implementation

uses
  ComObj;

function GetSpecialFolderPath(Folder: Integer; CanCreate: Boolean): string;
var
  FilePath: widestring;
  PFilePath: PWideChar; //array[0..MAX_PATH] of char;
begin
  SetLength(FilePath, MAX_PATH);
  PFilePath := Addr(FilePath[1]);
 { Get path of selected location }
  SHGetSpecialFolderPathW(0, PFilePath, Folder, CanCreate);
  Result := FilePath;
end;

function CreateShellLink(const AppName, Desc: string; Dest: Integer): string;
{ Creates a shell link for application or document specified in }
{ AppName with description Desc. Link will be located in folder }
{ specified by Dest, which is one of the string constants shown }
{ at the top of this unit. Returns the full path name of the  }
{ link file. }
var
  SL: IShellLink;
  PF: IPersistFile;
  LnkName: WideString;
begin
  OleCheck(CoCreateInstance(CLSID_ShellLink, nil, CLSCTX_INPROC_SERVER, IShellLink, SL));
 { The IShellLink implementer must also support the IPersistFile }
 { interface. Get an interface pointer to it. }
  PF := SL as IPersistFile;
  OleCheck(SL.SetPath(PChar(AppName))); // set link path to proper file
  if Desc <> '' then
    OleCheck(SL.SetDescription(PChar(Desc))); // set description
 { create a path location and filename for link file }
  LnkName := GetSpecialFolderPath(Dest, True) + '\' + ChangeFileExt(AppName, 'lnk');
  PF.Save(PWideChar(LnkName), True);     // save link file
  Result := LnkName;
end;

procedure GetShellLinkInfo(const LinkFile: WideString; var SLI: TShellLinkInfo);
{ Retrieves information on an existing shell link }
var
  SL: IShellLink;
  PF: IPersistFile;
  FindData: TWin32FindData;
  AStr: array[0..MAX_PATH] of char;
begin
  OleCheck(CoCreateInstance(CLSID_ShellLink, nil, CLSCTX_INPROC_SERVER, IShellLink, SL));
 { The IShellLink implementer must also support the IPersistFile }
 { interface. Get an interface pointer to it. }
  PF := SL as IPersistFile;
 { Load file into IPersistFile object }
  OleCheck(PF.Load(PWideChar(LinkFile), STGM_READ));
 { Resolve the link by calling the Resolve interface function. }
  OleCheck(SL.Resolve(0, SLR_ANY_MATCH or SLR_NO_UI));
 { Get all the info! }
  with SLI do
  begin
    OleCheck(SL.GetPath(AStr, MAX_PATH, FindData, SLGP_SHORTPATH));
    PathName := AStr;
    OleCheck(SL.GetArguments(AStr, MAX_PATH));
    Arguments := AStr;
    OleCheck(SL.GetDescription(AStr, MAX_PATH));
    Description := AStr;
    OleCheck(SL.GetWorkingDirectory(AStr, MAX_PATH));
    WorkingDirectory := AStr;
    OleCheck(SL.GetIconLocation(AStr, MAX_PATH, IconIndex));
    IconLocation := AStr;
    OleCheck(SL.GetShowCmd(ShowCmd));
    OleCheck(SL.GetHotKey(HotKey));
  end;
end;

procedure SetShellLinkInfo(const LinkFile: WideString; const SLI: TShellLinkInfo);
{ Sets information for an existing shell link }
var
  SL: IShellLink;
  PF: IPersistFile;
begin
  OleCheck(CoCreateInstance(CLSID_ShellLink, nil, CLSCTX_INPROC_SERVER, IShellLink, SL));
 { The IShellLink implementer must also support the IPersistFile }
 { interface. Get an interface pointer to it. }
  PF := SL as IPersistFile;
 { Load file into IPersistFile object }
  OleCheck(PF.Load(PWideChar(LinkFile), STGM_SHARE_DENY_WRITE));
 { Resolve the link by calling the Resolve interface function. }
  OleCheck(SL.Resolve(0, SLR_ANY_MATCH or SLR_UPDATE or SLR_NO_UI));
 { Set all the info! }
  with SLI, SL do
  begin
    OleCheck(SetPath(PChar(PathName)));
    OleCheck(SetArguments(PChar(Arguments)));
    OleCheck(SetDescription(PChar(Description)));
    OleCheck(SetWorkingDirectory(PChar(WorkingDirectory)));
    OleCheck(SetIconLocation(PChar(IconLocation), IconIndex));
    OleCheck(SetShowCmd(ShowCmd));
    OleCheck(SetHotKey(HotKey));
  end;
  PF.Save(PWideChar(LinkFile), True);  // save file
end;

end.

