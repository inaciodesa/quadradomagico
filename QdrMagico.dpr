program QdrMagico;

uses
  Forms,
  UnitPrincipal in 'UnitPrincipal.pas' {frmPrincipal},
  sobreUnit in 'sobreUnit.pas' {frmSobre};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Quadro Mágico';
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(TfrmSobre, frmSobre);
  Application.Run;
end.
