program Example;

uses
  Vcl.Forms,
  ExampleUnit in 'ExampleUnit.pas' {ExampleForm1},
  Winapi.FileThumbnail in 'Winapi.FileThumbnail.pas';

{$R *.res}

begin
  {$IFDEF DEBUG}
  ReportMemoryLeaksOnShutdown := True;
  {$ENDIF}
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TExampleForm1, ExampleForm1);
  Application.Run;
end.
