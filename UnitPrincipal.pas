unit UnitPrincipal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Grids, Buttons, ComCtrls, ComObj;

type
  TfrmPrincipal = class(TForm)
    Grade: TStringGrid;
    Panel1: TPanel;
    Label1: TLabel;
    txtNum: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Status: TStatusBar;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    Image1: TImage;
    procedure SpeedButton1Click(Sender: TObject);
    Procedure magic(n : integer);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure txtNumKeyPress(Sender: TObject; var Key: Char);
    procedure GerarExcel(Consulta: TStringGrid);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
   LV = nil;

   Type
    Tint = integer;
    Tcel = Tint;
    Tpont_L = ^TLinha;
    Tpont_C = ^TColuna;
    TLinha = Record
                valor : Tcel;
                col : Tpont_C;
                prox : Tpont_L;
             end;
    TColuna = Record
                valor : Tcel;
                Prox : Tpont_C;
              end;

Var n, max, centro : integer;
    M  : Tpont_L;
    frmPrincipal: TfrmPrincipal;

implementation

uses sobreUnit;

{$R *.DFM}

procedure InsereLinha(var P : TPont_L;dados : TCel);
var aux, L : Tpont_L;
begin
  new(aux);
  aux^.valor := dados;
  aux^.col := LV;
  aux^.prox := LV;
  if P <> lv then
     begin
       L := P;
       while L^.prox <> lv do
             L := l^.prox;
       L^.prox := aux;
     end else P := aux;
end;


procedure InsereColuna(var P : TPont_L;dados : TCel);
var aux, L : Tpont_C;
begin
  new(aux);
  aux^.valor := dados;
  aux^.prox := LV;
  if P^.col <> LV then
     begin
       L := P^.col;
       while L^.prox <> LV do
             L := L^.prox;
       L^.prox := aux;
     end else P^.col := aux;
end;

Function PrimeiraLinha(n : integer) : Tpont_L;
var
    Linha : TPont_L;
    i, CelAnt : integer;
    celula : Tcel;
begin
  Linha := lv;
  celula := n * centro;
  InsereLinha(Linha,celula);
  CelAnt := celula;
  for i := 2 to centro do
     begin
       celula := CelAnt - (n + 2);
       CelAnt := celula;
       InsereLinha(Linha,celula);
     end;
  celula := (max - 1);
  CelAnt := celula;
  InsereLinha(Linha,celula);
  for i := (centro + 1) to (centro + (centro - 2)) do
    begin
       celula := CelAnt - (n + 2);
       CelAnt := celula;
       InsereLinha(linha,celula);
    end;
    PrimeiraLinha := Linha;
end;

Procedure GeraMatrizImpar(M : TPont_L);
var
    valant, i : integer;
    celula : tcel;
begin
  while M <> LV do
  begin
     ValAnt := M^.valor;
     for i := 2 to n do
       begin
         if (valAnt mod n) = 0 then
             celula := ValAnt + 1
         else celula := ValAnt + (n + 1);
         if celula > max then
         celula := celula - max;
         ValAnt := celula;
         InsereColuna(M,celula);
       end;
     M := M^.prox;
  end;
end;

{ Imprime a matriz gerada na grade }
Procedure Imprime(Matriz : Tpont_L);
var
   i, j : integer;
   celula : Tcel;
   col : Tpont_C;
begin
        i := -1;
        while matriz <> lv do
         begin
           inc(i);
           celula := Matriz^.valor;
           Col := Matriz^.col;
           frmprincipal.Grade.cells[i,0] := inttostr(celula);
           j := 1;
           while col <> lv do
            begin
               celula := col^.valor;
               frmprincipal.Grade.cells[i,j] := inttostr(celula);
               col := col^.prox;
               inc(j);
           end;
         Matriz := Matriz^.prox;
         end;
end;

{ Preenche a grade com valores binarios }
Function GradeInicial(n : integer) : Tpont_L;
var i, j, valor, valor1, valor2 : integer;
    G : Tpont_L;
begin
G := Lv;
for i := 1 to n do
    begin
           valor1 := (i mod 4) div 2;
           valor2 := (1 mod 4) div 2;
           if valor2 =  valor1 then
              valor := 1
           else valor := 0;
           InsereLinha(G,valor);
    end;
GradeInicial := G;
j := 0;
while G <> lv do
    begin
       inc(j);
       for i := 2 to n do
       begin
           valor1 := (i mod 4) div 2;
           valor2 := (j mod 4) div 2;
           if valor2 = valor1 then
              valor := 1
           else valor := 0;
           InsereColuna(G,valor);
       end;
    G := G^.prox;
    end;
end;

{ ** Matriz para numeros pares: Quanto o n e multiplo de 4 ** }
Procedure GeraMatrizPar1(Matriz : Tpont_L);
var Aux : Tpont_L;
    col : Tpont_C;
    i, valor, contc : integer;
begin
Aux := Matriz;
  for i := 1 to n do
      begin
        if Matriz^.valor = 1 then
           Matriz^.valor := (n * n) + 1 - i
        else Matriz^.valor := i;
        Matriz := Matriz^.prox;
      end;

ContC := 1;
While Aux <> lv do
begin
    col := aux^.col;
    for i := 1 to n-1 do
      begin
           valor := n * (i) + ContC;
           if col^.valor = 1 then
              col^.valor := (n * n) + 1 - valor
           else col^.valor := valor;
           col := col^.prox; ;
      end;
    Aux := Aux^.prox;
    inc(ContC);
end;
end;

{ ** Insere o valor no inicio da Matriz (utilizado na MatrizPar2 ** }

Procedure InsereNoInicio(var Mat : Tpont_L;valor : integer);
var aux : Tpont_L;
begin
  new(aux);
  aux^.valor := valor;
  aux^.prox := Mat;
  aux^.col := lv;
  Mat := aux;
end;

Function Inverte(Mat : Tpont_L) : Tpont_L;
var aux : Tpont_L;
    a : Tpont_L;
begin
   a := lv;
   while Mat <> lv do
        begin
          new(aux);
          aux^.valor := mat^.valor;
          aux^.prox := a;
          aux^.col := mat^.col;
          a := aux;
          Mat := Mat^.prox;
        end;
   Inverte := a;

end;

{ Busca um valor com base na linha e coluna indicado }
Function busca(Matriz : Tpont_L;Linha : integer; coluna : integer) : Tcel;
var i, j : integer;
    col : Tpont_c;
begin
if coluna > 1 then
   for j := 2 to Coluna do
      Matriz := Matriz^.prox;
if Linha = 1 then
    busca := Matriz^.valor
else begin
   col := matriz^.col;
   for i := 2 to Linha  - 1 do
        col := col^.prox;
   busca := col^.valor;
end;
end;

{ Dado a linha e a coluna cadastra nessas posicoes o valor }
Procedure preenche(Matriz : Tpont_L;Linha, coluna, valor : integer);
var i, j : integer;
    col : Tpont_C;
begin
if coluna > 1 then
   for j := 1 to Coluna - 1 do
      Matriz := Matriz^.prox;
if Linha = 1 then
    Matriz^.valor := valor
else begin
   col := matriz^.col;
   for i := 2 to Linha - 1 do
        col := col^.prox;
   col^.valor := valor;
end;
end;

Procedure ListaDeColunas(var P : Tpont_C;valor : integer);
var aux, L : Tpont_C;
begin
  new(aux);
  aux^.valor := valor;
  aux^.prox := LV;
  if P <> lv then
     begin
       L := P;
       while L^.prox <> lv do
             L := L^.prox;
       L^.prox := aux;
     end
  else P := aux;
end;


{ Troca os numeros de posicao }
Procedure TrocaValores(M : Tpont_L;n : integer);
var i, j, k, L, p,         { indices }
    ant : integer;  { valores }
    col : Tpont_C;
begin
  p := n div 2;
  k := (n - 2) div 4;
  L := (n - k + 2);
  col := lv;
  { Gera uma lista com o numero das colunas que serao modificadas }

  for i := 1 to k do
      ListaDeColunas(col,i);

  if L <= n then
     for i := L to n do
         ListaDeColunas(col,i);
  { Troca os valos das colunas que estao na lista col }
  { ==> Primeira troca de valores }
  while col <> lv do
  begin
    j := col^.valor;
    for i := 1 to p do
       begin
          ant := busca(m,i,j);
          preenche(m,i,j,(busca(m,i+p,j)));
          preenche(m,i+p,j,ant);
       end;
    col := col^.prox;
  end;
  i := k + 1;
  col := lv;
  ListaDeColunas(col,1); ListaDeColunas(Col,i);

  { Segunda troca de valores }
  while col <> lv do
   begin
     j := col^.valor;
     ant := busca(m,i,j);
     preenche(m,i,j,(busca(m,i+p,j)));
     preenche(m,i+p,j,ant);
     col := col^.prox;
   end;
end;


{ Gera a matriz par de numeros que nao sao multiplos de 4 }
Procedure GeraMatrizPar2(var M : Tpont_L);
var i, j, p : integer;
    aux, Mval : TPont_L;
    col : Tpont_C;
begin
  p := n;
  n := n div 2;
  max := n * n;
  centro := (n div 2) + 1;
  aux := PrimeiraLinha(n);
  GeraMatrizImpar(aux);
  M := inverte(aux);
  Mval := M;
  aux := M;

  {Preenche a primeira linha da matriz}
  for i := 1 to n do
    begin
       inserelinha(aux,Mval^.valor + (2 * (n * n)));
       aux := aux^.prox;
       Mval := Mval^.prox;
    end;

  { Preenche o terceiro quadrante }
  aux := M;
  Mval := M;
  for i := 1 to n do
    begin
      inserecoluna(aux,Mval^.valor + (3 * (n * n)));
      col := Mval^.col;
      for j := 1 to n - 1 do
      begin
        insereColuna(aux,col^.valor + (3 * (n * n)));
        col := col^.prox;
      end;
      aux := aux^.prox;
      Mval := Mval^.prox;
    end;
  { Preenche o segundo e quarto qudrante }
  aux := M;
  Mval := M;
  while aux^.col <> lv do
     aux := aux^.prox;
  for i := 1 to n do
  begin
    col := Mval^.col;
    for j := 1 to (p - 1) do
    begin
       if j < n
         then insereColuna(aux,col^.valor + 2 * (n * n))
         else insereColuna(aux,col^.valor - 2 * (n * n));
       col := col^.prox;
    end;
    aux := aux^.prox;
    Mval := Mval^.prox;
  end;
  n := p;
  TrocaValores(M,n);
end;

Procedure TfrmPrincipal.magic(n : integer);
begin
if (n > 0) and (n <> 2) then
begin
  if n = 1 then
     begin
       grade.Cells[0,0] := '1';
    end
  else if (n mod 2) <> 0 then
    begin
      M := PrimeiraLinha(n);
      GeraMatrizImpar(M);
      Imprime(M);
    end
  else if (n mod 4) = 0 then
     begin
       M := GradeInicial(n);
       GeraMatrizPar1(M);
       Imprime(M);
     end
  else
     begin
       GeraMatrizPar2(M);
       Imprime(M);
     end;
end else
  begin
    showmessage ('Digite valores maiores que 1 e diferentes de 2');
  end;
end;

function IsInteger(TestaString: String) : boolean;
 begin
  try
  StrToInt(TestaString);
  except
  On EConvertError do Result := False;
  else
  Result := True;
  end;
end;

{ Limpa os dados armazenados no Gride}
Procedure Limpa(gr : TStringGrid);
var i : integer;
begin
for i := 0 to gr.RowCount do
    gr.rows[i].Clear;
gr.RowCount := 0;
gr.colCount := 0;
end;

procedure TfrmPrincipal.SpeedButton1Click(Sender: TObject);
var constante : integer;
begin
Limpa(grade);
if IsInteger(txtNum.text)  then
begin
 n := strtoInt(txtNum.text);
 max := n * n;
 centro := (n div 2) + 1;
 constante := (n * (n * n + 1)) div 2;
if n > 100 then
begin
  if application.messagebox('Valores muito elevados podem travar o sistema' + chr(13) + 'Deseja continuar ?','Valor elevado',MB_YESNO) = mryes then
   begin
      magic(n);
      Grade.ColCount := n;
      Grade.RowCount := n;
   end;
end else
      begin
         magic(n);
         Grade.ColCount := n;
         Grade.RowCount := n;
      end;
Status.panels[1].text := 'Matriz de ordem: ' + inttostr(n) + '    ' + 'Constante Mágica: ' + inttostr(constante) + '    '+ ' Numeros variando de 1 a ' + inttostr(N * N);
end else showmessage('Digite apenas numeros inteiros');
txtNum.text := '';
end;

procedure TfrmPrincipal.SpeedButton2Click(Sender: TObject);
begin
application.terminate;
end;

procedure TfrmPrincipal.FormActivate(Sender: TObject);
begin
  txtNum.SetFocus;
end;

procedure TfrmPrincipal.txtNumKeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(13) then
      SpeedButton1Click(sender);
end;

procedure TFrmPrincipal.GerarExcel(Consulta: TStringGrid);
var coluna, linha: integer;
    excel: variant;
begin
try
  excel := CreateOleObject('Excel.Application');
  excel.Workbooks.add(1);
except Application.MessageBox ('Versão do Ms-Excel'+
  'Incompatível','Erro',MB_OK+MB_ICONEXCLAMATION);
end;
try
for coluna := 0 to Consulta.colCount do
   for linha := 0 to Consulta.rowCount do
      excel.cells [linha+2,coluna+2] := consulta.cells[coluna, linha];
excel.columns.AutoFit;
excel.visible:=true;
except
  Application.MessageBox ('Aconteceu um erro desconhecido durante a conversão'+
  'da tabela para o Ms-Excel','Erro',MB_OK+MB_ICONEXCLAMATION);
end;
end;


procedure TfrmPrincipal.SpeedButton3Click(Sender: TObject);
begin
  GerarExcel(grade);
end;

procedure TfrmPrincipal.SpeedButton4Click(Sender: TObject);
begin
 frmsobre.showmodal;
end;

end.

