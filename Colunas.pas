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
    procedure DesenhaColuna(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmColunas: TfrmColunas;

function ColunasAceitas(tipo: string): TStringList;

implementation

uses
  uTema;

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
    Result.Add('CSOSN_ESTADUAL');
    Result.Add('CSOSN_ESTA_CF');
    Result.Add('CSOSN_ESTA_PRODUTOR');
    Result.Add('CSOSN_INTER');
    Result.Add('CSOSN_INTERESTADUAL');
    Result.Add('CSOSN_INTER_CF');
    Result.Add('CST');
    Result.Add('CST_ESTA');
    Result.Add('CST_ESTADUAL');
    Result.Add('CST_ESTA_CF');
    Result.Add('CST_ESTA_PRODUTOR');
    Result.Add('CST_INTER');
    Result.Add('CST_INTERESTADUAL');
    Result.Add('CST_INTER_CF');
    Result.Add('ALIQ_ICMS');
    Result.Add('ALIQ_ICMS_ESTA');
    Result.Add('ALIQ_ICMS_ESTA_CF');
    Result.Add('ALIQ_ICMS_ESTA_PRODUTOR');
    Result.Add('REDU');
    Result.Add('REDU_ESTA');
    Result.Add('REDU_ESTA_CF');
    Result.Add('REDU_ESTA_PRODUTOR');
    //Result.Add('REDU_INTER'); //Desativada: campo interestadual não é mais usado
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
  else if tipo = 'Transportadoras' then
  begin
    //Cadastro principal da transportadora (tabela TRAN), na ordem da tela do ERP
    Result.Add('CODI');
    Result.Add('ATIVO');
    Result.Add('CNPJ');
    Result.Add('CPF');
    Result.Add('CPF_CNPJ');
    Result.Add('INSC');
    Result.Add('NOME');
    Result.Add('ENDE');
    Result.Add('CIDA');
    Result.Add('UF');
    Result.Add('CODI_CIDA');
    Result.Add('FONE');
    Result.Add('PLACA');
    Result.Add('PLACA2');
    Result.Add('PLACA3');
    Result.Add('EMAIL');
    Result.Add('OBS');
    Result.Add('OBSERVACAO');
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

//Texto exibido na tela Colunas: CST, CSOSN, ALIQ_ICMS e REDU mostram a quais campos se aplicam
//(o grupo e, recuadas embaixo dele, as colunas avulsas); as demais colunas aparecem só com o nome
function TextoColuna(nome: string): string;
var
  base: string;
begin
  Result := nome;
  //Alíquota de ICMS e redução: ALIQ_ICMS e REDU valem para os 3 campos estaduais
  if (nome = 'ALIQ_ICMS') or (nome = 'REDU') then begin Result := nome + #9 + 'os 3 estaduais'; Exit; end;
  if (nome = 'ALIQ_ICMS_ESTA') or (nome = 'ALIQ_ICMS_ESTA_CF') or (nome = 'ALIQ_ICMS_ESTA_PRODUTOR') then begin Result := '     ' + nome + #9 + 'avulso'; Exit; end;
  if (nome = 'REDU_ESTA') or (nome = 'REDU_ESTA_CF') or (nome = 'REDU_ESTA_PRODUTOR') then begin Result := '     ' + nome + #9 + 'avulso'; Exit; end;
  if Pos('CSOSN', nome) = 1 then base := 'CSOSN'
  else if Pos('CST', nome) = 1 then base := 'CST'
  else Exit;
  if nome = base then Result := nome + #9 + 'os 5 campos'
  else if nome = base + '_ESTA' then Result := nome + #9 + '3 estaduais'
  else if nome = base + '_INTER' then Result := nome + #9 + '2 interestaduais'
  else if (nome = base + '_ESTADUAL') or (nome = base + '_ESTA_CF') or (nome = base + '_ESTA_PRODUTOR') or
          (nome = base + '_INTERESTADUAL') or (nome = base + '_INTER_CF') then Result := '     ' + nome + #9 + 'avulso';
end;

procedure TfrmColunas.MostrarColunas();
var
  lista: TStringList;
  i: Integer;
begin
  ListColunas.Items.Clear;
  lista := ColunasAceitas(LabelTipoImp.Caption);
  try
    //Desenho próprio: descrição das colunas de CST/CSOSN alinhada e em cinza
    ListColunas.Style := lbOwnerDrawFixed;
    ListColunas.OnDrawItem := DesenhaColuna;
    for i := 0 to lista.Count-1 do
      //Sinônimos aceitos na importação, mas que não aparecem na lista (ex.: OBS das transportadoras = OBSERVACAO)
      if not ((LabelTipoImp.Caption = 'Transportadoras') and (lista[i] = 'OBS')) then
        ListColunas.Items.Add(TextoColuna(lista[i]));
  finally
    lista.Free;
  end;
end;

//Desenha o nome da coluna e, para CST/CSOSN, a descrição alinhada à direita do nome (texto separado por #9)
procedure TfrmColunas.DesenhaColuna(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  texto, descr: string;
begin
  texto := ListColunas.Items[Index];
  descr := '';
  if Pos(#9, texto) > 0 then
  begin
    descr := Copy(texto, Pos(#9, texto)+1, MaxInt);
    texto := Copy(texto, 1, Pos(#9, texto)-1);
  end;
  //Mesmas faixas de cor da planilha (azul e amarelo alternados); item selecionado mantém o destaque padrão
  if not (odSelected in State) then
  begin
    //Primeira coluna com a mesma cor da primeira linha de dados da planilha
    ListColunas.Canvas.Brush.Color := CorFaixa(Index + 1);
    ListColunas.Canvas.Font.Color := CorTexto;
  end;
  ListColunas.Canvas.FillRect(Rect);
  ListColunas.Canvas.TextOut(Rect.Left + 2, Rect.Top, texto);
  if descr <> '' then
  begin
    if not (odSelected in State) then ListColunas.Canvas.Font.Color := CorDescricao;
    ListColunas.Canvas.TextOut(Rect.Left + 190, Rect.Top, descr);
  end;
end;

procedure TfrmColunas.FormShow(Sender: TObject);
begin
  MostrarColunas;
end;


procedure TfrmColunas.ListColunasKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  temp: string;
begin
   //RECONHECER CTRL+C
   if ((Shift = [ssCtrl]) and (Key = 67)) then
   begin
     if ListColunas.ItemIndex < 0 then Exit;
     temp := ListColunas.Items[ListColunas.ItemIndex];
     if Pos(#9, temp) > 0 then temp := Copy(temp, 1, Pos(#9, temp)-1);
     ClipBoard.AsText := Trim(temp);
   end;
end;

end.
