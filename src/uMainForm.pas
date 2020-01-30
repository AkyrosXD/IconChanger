unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, uShortcut, Vcl.ComCtrls;

type
  TMainForm = class(TForm)
    lbl1: TLabel;
    edtShotcutLocation: TEdit;
    btnBrowseShortcut: TButton;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    grpTargetImage: TGroupBox;
    edtImagePath: TEdit;
    edtImageURL: TEdit;
    btnSelectImage: TButton;
    btnChangeShortcut: TButton;
    btnClear: TButton;
    statCopyright: TStatusBar;
    btnClear2: TButton;
    procedure btnBrowseShortcutClick(Sender: TObject);
    procedure btnSelectImageClick(Sender: TObject);
    procedure btnChangeShortcutClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btnClear2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.btnBrowseShortcutClick(Sender: TObject);
var
  dlg: TOpenDialog;
  shortcutPath: string;
begin
  dlg := TOpenDialog.Create(nil);
  try
    dlg.Filter := 'Link Files (*.lnk)|*.lnk';
    if dlg.Execute(Handle) then
    begin
      shortcutPath := dlg.FileName;
    end;
  finally
    dlg.Free;
  end;
  edtShotcutLocation.Text := shortcutPath;
end;

procedure TMainForm.btnChangeShortcutClick(Sender: TObject);
var
  ShortcutPath, ImagePath, ImageURL, OutputMessage: string;
  uYesNo: Integer;
begin
  ShortcutPath := edtShotcutLocation.Text;
  ImagePath := edtImagePath.Text;
  ImageURL := edtImageURL.Text;
  if (ImagePath <> string.Empty) and (ImageURL <> string.Empty)  then
  begin
    Application.MessageBox('You must select an image or enter an image URL only!', 'Error', MB_ICONERROR);
    Exit;
  end;
  uYesNo := Application.MessageBox('Once you change the icon, it cannot be changed back. Are you sure that you want to continue?', 'Confirmation', MB_YESNO + MB_ICONWARNING);
  if(uYesNo = IDNO) then
  begin
    Exit;
  end;
  if (ImagePath <> string.Empty) and (ImageURL = string.Empty) then
  begin
    OutputMessage := ChangeIcon(ShortcutPath, ImagePath);
  end
  else if (ImagePath = string.Empty) and (ImageURL <> string.Empty) then
  begin
    OutputMessage := ChangeIconFromURL(ShortcutPath, ImageURL);
  end
  else
  begin
    Application.MessageBox('Please select an image or enter an image URL!', 'Error', MB_ICONERROR);
    Exit;
  end;
  if OutputMessage <> string.Empty then
  begin
    Application.MessageBox(PWideChar(OutputMessage), 'Error', MB_ICONERROR);
    Exit;
  end;
  Application.MessageBox('Icon changed successfully!', 'Information', MB_ICONASTERISK);
end;

procedure TMainForm.btnClear2Click(Sender: TObject);
begin
  edtImageURL.Text := string.Empty;
end;

procedure TMainForm.btnClearClick(Sender: TObject);
begin
  edtImagePath.Text := string.Empty;
end;

procedure TMainForm.btnSelectImageClick(Sender: TObject);
var
  dlg: TOpenDialog;
  imagePath: string;
begin
  dlg := TOpenDialog.Create(nil);
  try
    dlg.Filter := 'Image Filers|*.jpg;*.jpeg;*.png';
    if dlg.Execute(Handle) then
    begin
      imagePath := dlg.FileName;
    end;
  finally
    dlg.Free;
  end;
  edtImagePath.Text := imagePath;
end;

end.
