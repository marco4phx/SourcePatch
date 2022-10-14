program SrcPatch;

uses
  Vcl.Forms,
  Sorcio in 'Sorcio.pas' {Form1},
  units.StrReplace in 'units.StrReplace.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
