# FileThumbnail
Get file thumbnail, file icon or system file associated icon.

Sample:

![Example_FileThumbnail](https://github.com/ps147258/FileThumbnail/assets/34940706/260c599b-9496-4ac2-8c4a-920d965e4a77)

Additional component:
* JEDI: https://github.com/project-jedi
  * TJvFilenameEdit

Additional unit:
* Vcl.DragFiles: ​​https://github.com/ps147258/others_vcl/blob/master/Vcl.DragFiles.pas
  * TDropFiles

Example:
```
uses FileThumbnail;

procedure GetFileThumb;
var
  Thumb: TFileThumb;
  Icon: TIcon;
begin
  if GetShellIcon(Icon, JvFilenameEdit1.FileName, _IconBest.TypicallyPixels) then
    try
      Image1.Picture.Assign(Icon);
    finally
      Icon.Free;
    end
  else
    Image1.Picture.Assign(nil);
end;
```
