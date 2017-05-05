unit sobreUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Buttons;

type
  TfrmSobre = class(TForm)
    Image1: TImage;
    txtSobre: TMemo;
    BitBtn1: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSobre: TfrmSobre;

implementation

{$R *.DFM}

procedure percorreArquivoTexto ( nomeDoArquivo: String );
var arq: TextFile;
linha: String;
begin
AssignFile ( arq, nomeDoArquivo );
Reset ( arq );
// ReadLn ( arq, linha );
while not Eof ( arq ) do
begin
  ReadLn ( arq, linha );
  frmsobre.txtsobre.lines.add(linha);
end;
CloseFile ( arq );
end;

procedure TfrmSobre.BitBtn1Click(Sender: TObject);
begin
  close;
end;


procedure TfrmSobre.FormCreate(Sender: TObject);
begin
percorreArquivoTexto('qrMagic.ini');
end;

end.
