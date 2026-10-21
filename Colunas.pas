unit Colunas;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Clipbrd;

type
  TfrmColunas = class(TForm)
    LabelTipoImp: TLabel;
    ListColunas: TListBox;
    procedure MostrarColunas();
    procedure FormShow(Sender: TObject);
    procedure ListColunasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmColunas: TfrmColunas;

function ColunasAceitas(tipo: string): TStringList;

implementation

{$R *.dfm}

//Colunas aceitas por tipo de importação (tela Colunas e aviso de colunas não reconhecidas)
//Quem chama libera a lista
function ColunasAceitas(tipo: string): TStringList;
begin
  Result := TStringList.Create;

  //Colunas de cada tipo de importação, agrupadas em sequência lógica (seguindo as tabelas do banco)
  if tipo = 'Clie/Forn' then
  begin
    //Identificação
    Result.Add('CODI');
    Result.Add('EMPRESA');
    Result.Add('TIPOCAD');
    Result.Add('ATIVO');
    Result.Add('NOME');
    Result.Add('NOME_FANT');
    Result.Add('GRUPO');
    //Documentos
    Result.Add('CPF_CNPJ');
    Result.Add('CPF');
    Result.Add('CNPJ');
    Result.Add('RG');
    Result.Add('INSC');
    Result.Add('INSCR_PRODUTOR');
    //Dados pessoais
    Result.Add('DATA_NASC');
    Result.Add('SEXO');
    Result.Add('ESTA_CIVI');
    Result.Add('NOME_PAI');
    Result.Add('NOME_MAE');
    //Endereço
    Result.Add('ENDE');
    Result.Add('COMP');
    Result.Add('BAIR');
    Result.Add('CIDA');
    Result.Add('UF');
    Result.Add('CEP');
    Result.Add('PROX');
    //Contato
    Result.Add('FONE');
    Result.Add('FONE2');
    Result.Add('FAX');
    Result.Add('MAIL');
    //Trabalho
    Result.Add('FIRM');
    Result.Add('TRABALHA_DESDE');
    Result.Add('CARG');
    Result.Add('SALA');
    Result.Add('ENDE_FIRM');
    Result.Add('BAIR_FIRM');
    Result.Add('CIDA_FIRM');
    Result.Add('UF_FIRM');
    Result.Add('CEP_FIRM');
    Result.Add('FONE_FIRM');
    //Cônjuge
    Result.Add('CONJ');
    Result.Add('DATA_CONJ');
    Result.Add('CONJ_FIRM');
    Result.Add('CONJ_CARG');
    Result.Add('CONJ_SALA');
    //Outros
    Result.Add('REFE_COME');
    Result.Add('OBS');
  end
  else if tipo = 'Produtos' then
  begin
    //Identificação
    Result.Add('EMPR');
    Result.Add('CODI');
    Result.Add('DESCR');
    Result.Add('DESCR2');
    Result.Add('REFE');
    Result.Add('REFE_ORIGINAL');
    Result.Add('CODI_BARRA');
    Result.Add('CODI_BARRA_COM');
    Result.Add('ATIVO');
    //Classificação
    Result.Add('TIPO');
    Result.Add('GRUP');
    Result.Add('SUB_GRUP');
    Result.Add('DEPARTAMENTO');
    Result.Add('MARCA');
    Result.Add('COLECAO');
    //Unidade, medidas e fiscal
    Result.Add('UNID');
    Result.Add('FATOR_CONV');
    Result.Add('PESL');
    Result.Add('PESB');
    Result.Add('NCM');
    Result.Add('CEST');
    //Estoque
    Result.Add('QTD');
    Result.Add('EST');
    Result.Add('MAX');
    Result.Add('QTD_MINIMA');
    Result.Add('LOCALIZACAO');
    //Fornecedor
    Result.Add('FORN');
    //Custos e preços
    Result.Add('CUSTO');
    Result.Add('FRETE');
    Result.Add('IPI');
    Result.Add('ICMS');
    Result.Add('PIS');
    Result.Add('COFINS');
    Result.Add('CUSTO_REAL');
    Result.Add('CUSTO_MEDIO');
    Result.Add('MARGEM');
    Result.Add('PRECO_PRAZO');
    Result.Add('PRECO_VISTA');
    //Tributação
    Result.Add('CSOSN');
    Result.Add('CSOSN_ESTA');
    Result.Add('CSOSN_INTER');
    Result.Add('CST');
    Result.Add('CST_ESTA');
    Result.Add('CST_INTER');
    Result.Add('CST_ESTADUAL');
    Result.Add('CST_INTERESTADUAL');
    Result.Add('CST_ESTA_CF');
    Result.Add('CST_INTER_CF');
    Result.Add('CST_ESTA_PRODUTOR');
    Result.Add('ALIQ_ICMS');
    Result.Add('REDU_ESTA');
    Result.Add('REDU_INTER');
    Result.Add('CST_IPI');
    Result.Add('ALIQ_IPI');
    Result.Add('CST_PIS');
    Result.Add('ALIQ_PIS');
    Result.Add('CST_COFINS');
    Result.Add('ALIQ_COFINS');
    //Outros
    Result.Add('OBS');
  end
  else if tipo = 'Grupos' then
  begin
    Result.Add('CODI');
    Result.Add('EMPR');
    Result.Add('DESCR');
  end
  else if tipo = 'SubGrupos' then
  begin
    Result.Add('CODI');
    Result.Add('EMPR');
    Result.Add('DESCR');
  end
  else if tipo = 'Marcas' then
  begin
    Result.Add('CODI');
    Result.Add('DESCR');
  end
  else if tipo = 'Títulos a Pagar' then
  begin
    //Identificação
    Result.Add('CODI');
    Result.Add('EMPR');
    Result.Add('FORN');
    //Datas e valores
    Result.Add('DATA');
    Result.Add('VENC');
    Result.Add('VALO');
    Result.Add('SALD');
    Result.Add('HIST');
    //Cobrança e financeiro
    Result.Add('OPER');
    Result.Add('CART');
    Result.Add('LOCA_COBR');
    Result.Add('C_FUNC');
    //Baixa
    Result.Add('DATA_BAIXA');
  end
  else if tipo = 'Títulos a Receber' then
  begin
    //Identificação
    Result.Add('CODI');
    Result.Add('EMPR');
    Result.Add('CLIE');
    //Datas e valores
    Result.Add('DATA');
    Result.Add('VENC');
    Result.Add('VALO');
    Result.Add('SALD');
    Result.Add('HIST');
    //Cobrança e financeiro
    Result.Add('OPER');
    Result.Add('CART');
    Result.Add('LOCA_COBR');
    Result.Add('C_FUNC');
    //Baixa
    Result.Add('DATA_BAIXA');
  end;
end;

procedure TfrmColunas.MostrarColunas();
var
  lista: TStringList;
begin
  ListColunas.Items.Clear;
  lista := ColunasAceitas(LabelTipoImp.Caption);
  try
    ListColunas.Items.AddStrings(lista);
  finally
    lista.Free;
  end;
end;

procedure TfrmColunas.FormShow(Sender: TObject);
begin
  MostrarColunas;
end;


procedure TfrmColunas.ListColunasKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   //RECONHECER CTRL+C
   if ((Shift = [ssCtrl]) and (Key = 67)) then
   begin
     ClipBoard.AsText := ListColunas.Items[ListColunas.ItemIndex];
   end;
end;

end.
