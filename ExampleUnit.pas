unit ExampleUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Mask, Vcl.ComCtrls, Vcl.Imaging.pngimage,
  JvExMask, JvToolEdit,
  Vcl.DragFiles, // https://github.com/ps147258/others_vcl/blob/master/Vcl.DragFiles.pas
  Winapi.FileThumbnail;

type
  TExampleForm1 = class(TForm)
    JvFilenameEdit1: TJvFilenameEdit;
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    CheckBox1: TCheckBox;
    Edit1: TEdit;
    UpDown1: TUpDown;
    Button1: TButton;
    Panel1: TPanel;
    Image1: TImage;
    Image2: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
    DropFiles: TDropFiles;
    procedure OnDropEvent(Sender: TObject; WinControl: TWinControl);
    procedure UpdateExtractSize;
    function GetPixels(SizeType: TIconSizeType): TSize; inline;
    function GetSize(SizeType: TIconSizeType): Word; inline;
    function ExecShellIcon(SizeType: TIconSizeType): Boolean;
    function ExecThumbnail(SizeType: TIconSizeType): Boolean;
    function ExecFileIcon(Index: Integer; SizeType: TIconSizeType): Boolean;
    function ExecSysImageListIcon(SizeType: TIconSizeType): Boolean;
  public
    { Public declarations }
  end;

var
  ExampleForm1: TExampleForm1;

implementation

//uses
//  Debug;

{$R *.dfm}

procedure TExampleForm1.FormCreate(Sender: TObject);
begin
  UpdateExtractSize;
end;

procedure TExampleForm1.FormShow(Sender: TObject);
begin
  DropFiles := TDropFiles.Create(Panel1, OnDropEvent);
end;

procedure TExampleForm1.OnDropEvent(Sender: TObject; WinControl: TWinControl);
begin
  JvFilenameEdit1.FileName := TDropFiles(Sender).First;
  Button1.Click;
end;

procedure TExampleForm1.UpdateExtractSize;
var
  Mode: TIconSizeType;
  Modes: TIconSizeTypes;
  function TypicallyStr(M: TIconSizeType): string;
  begin
    Result := Format('%s[%d]', [M.Str, M.TypicallySize]);
  end;
  function ActualStr(M: TIconSizeType): string;
  var
    Size: TSize;
  begin
    Size := M.ActualPixels;
    Result := Format('%s[%dx%d]', [M.Str, Size.cx, Size.cy]);
  end;
  procedure SetExtractStr(M: TIconSizeType; const S: string);
  var
    Strings: TStrings;
  begin
    Strings := RadioGroup2.Items;
    if Integer(M) < Strings.Count then
      Strings[Integer(M)] := S
    else
      Strings.Add(S);
  end;
begin
  if CheckBox1.Checked then
  begin
    Modes := CatchImageListAll;
    for Mode := _IconBase to _IconBest do
      if Mode in Modes then
        SetExtractStr(Mode, ActualStr(Mode))
      else
        SetExtractStr(Mode, TypicallyStr(Mode));
    Exit;
  end;

  for Mode := _IconBase to _IconBest do
    SetExtractStr(Mode, TypicallyStr(Mode));
end;

procedure TExampleForm1.CheckBox1Click(Sender: TObject);
begin
  UpdateExtractSize;
end;

procedure TExampleForm1.Button1Click(Sender: TObject);
var
  Succeed: Boolean;
  SizeType: TIconSizeType;
begin
  SizeType := TIconSizeType(RadioGroup2.ItemIndex);
  case RadioGroup1.ItemIndex of
    0: Succeed := ExecShellIcon(SizeType);
    1: Succeed := ExecThumbnail(SizeType);
    2: Succeed := ExecFileIcon(UpDown1.Position, SizeType);
    3: Succeed := ExecSysImageListIcon(SizeType);
    else Succeed := False;
  end;
  if not Succeed then
    Image1.Picture.Assign(nil)
end;

function TExampleForm1.GetPixels(SizeType: TIconSizeType): TSize;
begin
  if CheckBox1.Checked then
    Result := SizeType.ActualPixels
  else
    Result := SizeType.TypicallyPixels;
end;

function TExampleForm1.GetSize(SizeType: TIconSizeType): Word;
  function MaxPixels(S: TSize): Word; inline;
  begin
    if S.cy > S.cx then
      Result := S.cy
    else
      Result := S.cx;
  end;
begin
  if CheckBox1.Checked then
    Result := MaxPixels(SizeType.ActualPixels)
  else
    Result := SizeType.TypicallySize;
end;

function TExampleForm1.ExecShellIcon(SizeType: TIconSizeType): Boolean;
var
  Icon: TIcon;
begin
  Result := GetShellIcon(Icon, JvFilenameEdit1.FileName, GetPixels(SizeType));
  if not Result then Exit;
  try
    Caption := Format('ShellIcon %dx%d', [Icon.Width, Icon.Height]);
    Image1.Picture.Assign(Icon);
  finally
    Icon.Free;
  end;
end;

function TExampleForm1.ExecThumbnail(SizeType: TIconSizeType): Boolean;
var
  Bitmap: TBitmap;
begin
  Result := GetThumbnail(Bitmap, JvFilenameEdit1.FileName, GetPixels(SizeType));
  if not Result then Exit;
  try
    Caption := Format('Thumbnail %dx%d', [Icon.Width, Icon.Height]);
    Image1.Picture.Assign(Bitmap);
  finally
    Bitmap.Free;
  end;
end;

function TExampleForm1.ExecFileIcon(Index: Integer; SizeType: TIconSizeType): Boolean;
var
  Icon: TIcon;
begin
  Result := GetFileIcon(Icon, JvFilenameEdit1.FileName, Index, GetSize(SizeType)) <> ImageSize_Non;
  if not Result then Exit;
  try
    Caption := Format('FileIcon[%d] %dx%d', [Index, Icon.Width, Icon.Height]);
    Image1.Picture.Assign(Icon);
  finally
    Icon.Free;
  end;
end;

function TExampleForm1.ExecSysImageListIcon(SizeType: TIconSizeType): Boolean;
var
  Icon: TIcon;
begin
  Result := GetSysImageListIcon(Icon, JvFilenameEdit1.FileName, SizeType);
  if not Result then Exit;
  try
    Caption := Format('SysImageListIcon %s %dx%d', [SizeType.Id, Icon.Width, Icon.Height]);
    Image1.Picture.Assign(Icon);
  finally
    Icon.Free;
  end;
end;

end.
