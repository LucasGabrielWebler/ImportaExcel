unit uDividir;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TfrmDividir = class(TForm)
    Label1: TLabel;
    edtChar: TEdit;
    btnDividir: TButton;
    procedure FormShow(Sender: TObject);
    procedure btnDividirClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDividir: TfrmDividir;

implementation

uses
  importa_excel, uUtil;

{$R *.dfm}

procedure TfrmDividir.btnDividirClick(Sender: TObject);
var
  i,j, idxChar: Integer;
  temp: string;
  dividiu: Boolean;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  //-------------------------------
  //Modificar função inserir coluna

  //Aumentar colCount
  frmPrinc.StringGrid1.ColCount := frmPrinc.StringGrid1.ColCount + 1;
  //Receber ultima coluna
  i := frmPrinc.StringGrid1.ColCount;

  //Flag para dizer se dividiu ou não
  dividiu := False;

  //Percorrer até a coluna atual copiando valores das células para direita
  while i>=frmPrinc.StringGrid1.Col do
  begin
    for j := 0 to frmPrinc.StringGrid1.RowCount do
    begin
      temp := frmPrinc.StringGrid1.Cells[i,j];
      //Se for a coluna, verificar caracter divisor e copia metade pra cima
      if i = frmPrinc.StringGrid1.Col then
      begin
        //Procurar caracter divisor
        idxChar := temp.IndexOf(edtChar.Text);
        if idxChar > 0 then
        begin
          //Escrever campo da direita (após caracter divisor)
          frmPrinc.StringGrid1.Cells[i+1,j] := Copy(temp, idxChar+1+Length(edtChar.Text), MaxInt);
          //Escrever campo da esquerda (até o caracter divisor)
          frmPrinc.StringGrid1.Cells[i,j] := Copy(temp, 0, idxChar);
          //Setar flag dividiu
          dividiu := True;
        end
        //Se não achou o caracter, coloca em branco na direita
        else
        begin
          frmPrinc.StringGrid1.Cells[i+1,j] := '';
        end;
      end
      //Sen escreve na direita
      else
        frmPrinc.StringGrid1.Cells[i,j] := frmPrinc.StringGrid1.Cells[i-1,j];
    end;
    i:= i-1;
  end;

  //Se não dividiu nenhuma celula, remover coluna criada em branco
  if not dividiu then
  begin
    DeleteCol(frmPrinc.StringGrid1, frmPrinc.StringGrid1.Col+1);
  end;
end;

procedure TfrmDividir.FormShow(Sender: TObject);
begin
  edtChar.Text := '/';
end;

end.
