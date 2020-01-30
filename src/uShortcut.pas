unit uShortcut;

interface

uses
  WinShell, System.SysUtils, Vcl.Graphics, ShlObj, ComObj, ActiveX, IOUtils,
  Winapi.UrlMon, DateUtils;

function ChangeIcon(TargetLink, NewImageLocation: string): string;

function ChangeIconFromURL(TargetLink, ImageURL: string): string;

implementation

function ImageToBitmap(Filename: string): TBitmap;
var
  SourceImage: TWICImage;
begin
  if not TFile.Exists(Filename) then
  begin
    Exit(nil);
  end;
  SourceImage := TWICImage.Create;
  try
    SourceImage.LoadFromFile(Filename);
  except
    try
      TFile.Delete(Filename);
    except
    end;
    SourceImage.Free;
    Exit(nil);
  end;
  Result := TBitmap.Create;
  Result.Assign(SourceImage);
end;

function CreateTempDirectories(var BitmapsDir, ImagesDir: string): string;
var
  ProgramTempDirectory, BitmapsDirectory, ImagesDirectory: string;
begin
  ProgramTempDirectory := TPath.Combine(TPath.GetTempPath, 'IconChanger');
  if not TDirectory.Exists(ProgramTempDirectory) then
  begin
    if not CreateDir(ProgramTempDirectory) then
    begin
      Exit('Unable to create directory "' + ProgramTempDirectory + '"');
    end;
  end;
  BitmapsDirectory := TPath.Combine(ProgramTempDirectory, 'Bitmaps');
  if not TDirectory.Exists(BitmapsDirectory) then
  begin
    if not CreateDir(BitmapsDirectory) then
    begin
      Exit('Unable to create directory "' + BitmapsDirectory + '"');
    end;
  end;
  ImagesDirectory := TPath.Combine(ProgramTempDirectory, 'Images');
  if not TDirectory.Exists(ImagesDirectory) then
  begin
    if not CreateDir(ImagesDirectory) then
    begin
      Exit('Unable to create directory "' + ImagesDirectory + '"');
    end;
  end;
  BitmapsDir := BitmapsDirectory;
  ImagesDir := ImagesDirectory;
  Result := string.Empty;
end;

function SaveBitmapImage(BitmapImage: TBitmap; SaveLocation: string): Boolean;
begin
  try
    BitmapImage.SaveToFile(SaveLocation);
  except
    Exit(False);
  end;
  Result := True;
end;

function CurrentDateTimeStr: string;
var
  CurrYear, CurrMonth, CurrDay, CurrHour, CurrMinute, CurrSecond: string;
begin
  CurrYear := IntToStr(YearOf(Now));
  CurrMonth := IntToStr(MonthOfTheYear(Now));
  CurrDay := IntToStr(DayOfTheMonth(Now));
  CurrHour := IntToStr(HourOfTheDay(Now));
  CurrMinute := IntToStr(MinuteOfTheHour(Now));
  CurrSecond := IntToStr(SecondOfTheMinute(Now));
  Result := CurrYear + CurrMonth + CurrDay + '_' + CurrHour + CurrMinute + CurrSecond;
end;

function DownloadImage(ImageURL: string): string;
var
  BitmapsDir, ImagesDir, ImageName, ImagePath: string;
  StatusCode: Integer;
begin
  if CreateTempDirectories(BitmapsDir, ImagesDir) <> string.Empty then
  begin
    Exit(string.Empty);
  end;
  ImageName := 'image_' + CurrentDateTimeStr + '.png';
  ImagePath := TPath.Combine(ImagesDir, ImageName);
  try
    StatusCode := URLDownloadToFile(nil, PChar(ImageURL), PChar(ImagePath), 0, nil);
    if StatusCode <> 0 then
    begin
      Exit(string.Empty);
    end;
  except
    Exit(string.Empty);
  end;
  Result := ImagePath
end;

function ChangeIcon(TargetLink, NewImageLocation: string): string;
var
  CreateDirError, BitmapsDir, ImagesDir, BitmapFileName, BitmapLocation: string;
  BitmapImage: TBitmap;
  shlTest: TShellLinkInfo;
begin
  if not TFile.Exists(TargetLink) then
  begin
    Exit('The selected shortcut could not be found!');
  end;
  CreateDirError := CreateTempDirectories(BitmapsDir, ImagesDir);
  if CreateDirError <> string.Empty then
  begin
    Exit(CreateDirError);
  end;
  BitmapImage := ImageToBitmap(NewImageLocation);
  if BitmapImage = nil then
  begin
    Exit('The selected image is invalid or could not be found!');
  end;
  BitmapFileName := ExtractFileName(TPath.GetFileNameWithoutExtension(NewImageLocation)) + '.bmp';
  BitmapLocation := TPath.Combine(BitmapsDir, BitmapFileName);
  if not SaveBitmapImage(BitmapImage, BitmapLocation) then
  begin
    Exit('Failed to save Bitmap file.');
  end;
  GetShellLinkInfo(TargetLink, shlTest);
  shlTest.IconLocation := BitmapLocation;
  try
    SetShellLinkInfo(TargetLink, shlTest);
  except
    Exit('Invalid Image!');
  end;
  Result := string.Empty;
end;

function ChangeIconFromURL(TargetLink, ImageURL: string): string;
var
  ImageLocation: string;
begin
  ImageLocation := DownloadImage(ImageURL);
  if ImageLocation = string.Empty then
  begin
    Exit('Error while downloading the image from the URL!');
  end;
  Result := ChangeIcon(TargetLink, ImageLocation);
end;

end.
