unit importa_excel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.Buttons, ComObj, IniFiles,
  Vcl.FileCtrl, Data.DBXFirebird, Data.DB, Data.SqlExpr, importando, OleAuto, Vcl.Menus, System.StrUtils,
  empresa, Colunas, uSubstituir, uUtil, uProcurar, Vcl.ExtCtrls, uDividir;

type
  TfrmPrinc = class(TForm)
    opnDadosOrigem: TOpenDialog;
    StringGrid1: TStringGrid;
    opnDadosDestino: TOpenDialog;
    conDestino: TSQLConnection;
    SaveDialog1: TSaveDialog;
    Menu: TMainMenu;
    t1: TMenuItem;
    Editar1: TMenuItem;
    mnuCabecalho: TMenuItem;
    Limpar: TMenuItem;
    mnuLimpaClieForn: TMenuItem;
    mnuLimpaTransportadoras: TMenuItem;
    mnuLimpaGrupos: TMenuItem;
    mnuLimpaSubGrupos: TMenuItem;
    mnuLimpaMarcas: TMenuItem;
    mnuLimpaProdutos: TMenuItem;
    mnuLimpaGrades: TMenuItem;
    mnuLimpaTituP: TMenuItem;
    mnuLimpaTituR: TMenuItem;
    mnuTema: TMenuItem;
    mnuBackup: TMenuItem;
    mnuBackupAgora: TMenuItem;
    mnuBackupPasta: TMenuItem;
    mnuAdicionarColuna: TMenuItem;
    mnuAdicionarLinha: TMenuItem;
    mnuDeletarColuna: TMenuItem;
    mnuDeletarLinha: TMenuItem;
    mnuDadosEmpr: TMenuItem;
    mnuColunas: TMenuItem;
    conOrigem: TSQLConnection;
    N1: TMenuItem;
    mnuSubstituir: TMenuItem;
    mnuProcurar: TMenuItem;
    mnuDividir: TMenuItem;
    pnlTop: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lblColUpdate: TLabel;
    lblTableName: TLabel;
    btnLoadOrigem: TBitBtn;
    FilePath: TEdit;
    btnAbrirOrigem: TBitBtn;
    SelectImport: TComboBox;
    ButImport: TBitBtn;
    DBPath: TEdit;
    ButOpenDB: TBitBtn;
    StartLine: TEdit;
    btnTXT: TBitBtn;
    edtTableName: TEdit;
    btnSalvar: TBitBtn;
    BtnDestinoPadrao: TButton;

    function quantidadeEmpresas(colEmpr: Integer): Integer;
    procedure btnAbrirOrigemClick(Sender: TObject);
    procedure btnLoadOrigemClick(Sender: TObject);
    procedure BtnOpenDB(Sender: TObject);
    procedure StringGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ButImportClick(Sender: TObject);
    procedure btnSalvarClick(Sender: TObject);
    procedure mnuCabecalhoClick(Sender: TObject);
    procedure mnuLimpaClieFornClick(Sender: TObject);
    procedure mnuLimpaTransportadorasClick(Sender: TObject);
    procedure mnuLimpaGruposClick(Sender: TObject);
    procedure mnuLimpaSubGruposClick(Sender: TObject);
    procedure mnuLimpaMarcasClick(Sender: TObject);
    procedure mnuLimpaProdutosClick(Sender: TObject);
    procedure mnuLimpaGradesClick(Sender: TObject);
    procedure mnuLimpaTituPClick(Sender: TObject);
    procedure mnuLimpaTituRClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mnuTemaItemClick(Sender: TObject);
    procedure mnuBackupAgoraClick(Sender: TObject);
    procedure mnuBackupPastaClick(Sender: TObject);
    procedure mnuAdicionarColunaClick(Sender: TObject);
    procedure mnuAdicionarLinhaClick(Sender: TObject);
    procedure mnuDeletarColunaClick(Sender: TObject);
    procedure mnuDeletarLinhaClick(Sender: TObject);
    procedure mnuDadosEmprClick(Sender: TObject);
    procedure mnuColunasClick(Sender: TObject);
    procedure btnTXTClick(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure mnuSubstituirClick(Sender: TObject);
    function getProdCodUpdate(line: Integer) : string;
    procedure SelectImportChange(Sender: TObject);
    procedure mnuProcurarClick(Sender: TObject);
    procedure mnuDividirClick(Sender: TObject);
    procedure BtnDestinoPadraoClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    function StringGridToArray(Grid: TStringGrid): Integer;
    function ColunasCST(colunas: array of string): Integer;
    function cadastraClieForn(colClieForn,dadosClieForn: string): Integer;
    class function buscaCidade(Cidade, UF: string): String;
    function temCodTituloP(Codigo: String): Boolean;
    function temCodTituloR(Codigo: String): Boolean;
    function temGrupo(Codigo: String): Boolean;
    function temSubGrup(Codigo: String): Boolean;
    function temMarca(Codigo: String): Boolean;
    function getCodiClieForn(clieforn: String): Integer;
  end;

var
  //Variaveis globais
  frmPrinc: TfrmPrinc;
  colUpdate: Array of string;
  colUpdateCount: Integer;
  gridTemp: Array of Array of string;
  qtdEmpr: Integer;
  status: Integer;
  butContinue: Integer;

implementation

uses
  Data.DBXCommon, uImportaClieForn, uImportaProduto, uImportaGrupo, uImportaSubGrupo,
  uImportaMarca, uImportaTituP, uImportaTituR, uImportaCUSTOM, uImportaTran, uImportaGrade, uTema, uBackup;

const
  //Recálculo das margens dos produtos (sem estouro numérico e com o sinal correto)
  //MARGEM1: margem do preço a prazo sobre o custo real (negativa quando o preço é menor que o custo)
  SQL_MARGEM1 = 'UPDATE prod_custos pc SET pc.cust_margem1 = CAST((pc.cust_preco_prazo - pc.cust_custo_real)*100 AS NUMERIC(18,2)) / CAST(pc.cust_custo_real AS DOUBLE PRECISION) WHERE pc.cust_preco_prazo > 0 AND pc.cust_custo_real > 0;';
  //MARGEM2: desconto do preço a vista sobre o preço a prazo (negativa quando a vista é maior)
  SQL_MARGEM2 = 'UPDATE prod_custos pc SET pc.cust_margem2 = CAST((pc.cust_preco_prazo - pc.cust_preco_vista)*100 AS NUMERIC(18,2)) / CAST(pc.cust_preco_prazo AS DOUBLE PRECISION) WHERE pc.cust_preco_prazo > 0;';

//Função para definir a quantidade de empresas
function TfrmPrinc.quantidadeEmpresas(colEmpr: Integer): Integer;
var
  i, max: Integer;
begin
  max := 1;
  for i := StrToInt(StartLine.Text) to StringGrid1.RowCount-1 do begin
    //Vazio ou inválido conta como empresa 1 (o valor inválido para a importação na própria linha)
    if StrToIntDef(Trim(StringGrid1.Cells[colEmpr,i]),1) > max  then begin
      max := StrToIntDef(Trim(StringGrid1.Cells[colEmpr,i]),1);
    end;
  end;
  //Banco FDB: não passar da maior empresa cadastrada
  //(empresa inexistente para a importação somente na própria linha)
  if UpperCase( ExtractFileExt(DBPath.Text) ) = '.FDB' then
    if max > StrToIntDef(querySelect('SELECT MAX(codi) FROM dados_empre'), 1) then
      max := StrToIntDef(querySelect('SELECT MAX(codi) FROM dados_empre'), 1);
  Result := max;
end;


//Botão para selecionar arquivo
procedure TfrmPrinc.btnAbrirOrigemClick(Sender: TObject);
var
  arquivo : String;

begin
  if opnDadosOrigem.Execute then
  begin
    arquivo := ExtractFilePath(Application.ExeName);
    FilePath.Text := opnDadosOrigem.FileName;
  end;
  btnLoadOrigem.SetFocus;
end;


//Botão para selecionar arquivo
procedure TfrmPrinc.BtnOpenDB(Sender: TObject);
var
  arquivo : String;

begin
  if opnDadosDestino.Execute then
  begin
    conDestino.Connected := False;
    arquivo := ExtractFilePath(Application.ExeName);
    DBPath.Text := opnDadosDestino.FileName;
    conDestino.Params.Values['DataBase'] := DBPath.Text;
  end;

end;


//Botão para Carregar arquivo Excel na StringGrid
procedure TfrmPrinc.btnLoadOrigemClick(Sender: TObject);
var
  i: integer;
  fileExt :string;

begin
  //Limpar StringGrid
  StringGrid1.ColCount := 1;
  StringGrid1.RowCount := 1;

  //Carregar extensão do arquivo
  fileExt := LowerCase(ExtractFileExt(FilePath.Text));

  //Carregar arquivo de acordo com a extensão
  if (fileExt='.xls') or (fileExt='.xlsx') then
  begin
    //Carregar Excel na StringGrid
    Xls_To_StringGrid(StringGrid1, FilePath.Text);
  end
  else if (fileExt='.csv') then
  begin
    //Carregar CSV na StringGrid
    CSV_To_StringGrid(StringGrid1, FilePath.Text);
  end
  else if (fileExt='.fdb') then
  begin
    //Carregar FDB no conOrigem
    ShowMessage('Funcionalidade de carregar base ADMERP ainda não implementada');
    conOrigem.Params.Values['DataBase'] := FilePath.Text;
  end
  else
  begin
    //Mensagem de extenção não suportada
    ShowMessage('Extenção não suportada. ( '+fileExt+' )');
  end
  ;

  //Remover linhas em branco
  RemoveWhiteRows(StringGrid1);

  //Remover espaços no inicio e fim das strings
  RemoveSpaces(StringGrid1);

  //Redimensionar colunas
  for i := 0 to (StringGrid1.ColCount - 1) do
    AutoSizeCol(StringGrid1, i);

end;


//Função para cadastrar cliente/fornecedor no banco de dados. Retorna Gen_ID
function TfrmPrinc.cadastraClieForn(colClieForn,dadosClieForn: string): Integer;
var
  queryTemp: TSQLQuery;

begin
  Result := -1;
  queryTemp := TSQLQuery.Create(nil);
  try
    try
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;

      //Desativar Trigger das cidades
      queryTemp.CommandText := 'ALTER TRIGGER clieforn_biu0 INACTIVE;';
      queryTemp.ExecSQL;
      try
        //Executar INSERT
        queryTemp.CommandText := 'INSERT INTO clieforn ('+ colClieForn +') VALUES ' + '(' + dadosClieForn + ');';
        queryTemp.ExecSQL;
      finally
        //Reativar Trigger das cidades, mesmo se o INSERT der erro
        //(ExecuteDirect para manter o INSERT no CommandText da mensagem de erro)
        frmPrinc.conDestino.ExecuteDirect('ALTER TRIGGER clieforn_biu0 ACTIVE;');
      end;

      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT c.codi FROM clieforn c WHERE c.codi = gen_id(gen_clieforn_id,0);');
      queryTemp.Open;
      if not queryTemp.IsEmpty then
        Result := queryTemp.FieldByName('CODI').AsInteger;

    except
      on e: exception do
      begin
        ShowMessage('Erro SQL: '+e.message+sLineBreak+queryTemp.CommandText);
      end;
    end;
  finally
    queryTemp.Free;
    frmPrinc.conDestino.Close;
  end;
end;


//Função para buscar a cidade no banco.
{
  A combinação ´class function´ na declaração de um método diz ao compilador que
  aquele método pode ser chamado a partir da própria classe, sem a necessidade
  de se instanciar um objeto dela.
}
class function TfrmPrinc.buscaCidade(Cidade, UF: string): String;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  //Arquivo TXT/SQL: subconsulta resolvida na execução do script (apóstrofo duplicado para o SQL)
  then Result := '(SELECT FIRST 1 CID_CODI FROM CIDADE WHERE CID_DESC = '''+stringreplace(Cidade, '''', '''''',[rfReplaceAll])+''' AND CID_UF = '''+UF+''')'
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT * FROM CIDADE WHERE CID_DESC = :PDESC AND CID_UF = :PUF');
      queryTemp.ParamByName('PDESC').AsString := Cidade;
      queryTemp.ParamByName('PUF').AsString := UF;
      queryTemp.Open;
    finally
      if queryTemp.IsEmpty then
      begin
        Result := '';
      end
      else begin
        Result := queryTemp.FieldByName('CID_CODI').AsString;
      end;
      queryTemp.Close;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RECONHECER SE JA EXISTE O CODIGO DO TITULO PAGAR OU NAO
function TfrmPrinc.temCodTituloP(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := False
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT tp.codi FROM titup tp WHERE tp.codi = :PCODI');
      queryTemp.ParamByName('PCODI').AsString := Codigo;
      queryTemp.Open;

      if queryTemp.IsEmpty = True then
        Result := False
      else
        Result := True;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RECONHECER SE JA EXISTE O CODIGO DO TITULO RECEBER OU NAO
function TfrmPrinc.temCodTituloR(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := False
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT tr.codi FROM titur tr WHERE tr.codi = :PCODI');
      queryTemp.ParamByName('PCODI').AsString := Codigo;
      queryTemp.Open;

      if queryTemp.IsEmpty = True then
        Result := False
      else
        Result := True;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RECONHECER SE JA EXISTE O CODIGO DO GRUPO OU NAO
function TfrmPrinc.temGrupo(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := False
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT g.codi FROM grup_prod g WHERE g.codi = :PCODI');
      queryTemp.ParamByName('PCODI').AsString := Codigo;
      queryTemp.Open;

      if queryTemp.IsEmpty = True then
        Result := False
      else
        Result := True;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RECONHECER SE JA EXISTE O CODIGO DO SUB GRUPO OU NAO
function TfrmPrinc.temSubGrup(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := False
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT sg.codi FROM sub_grup_prod sg WHERE sg.codi = :PCODI');
      queryTemp.ParamByName('PCODI').AsString := Codigo;
      queryTemp.Open;

      if queryTemp.IsEmpty = True then
        Result := False
      else
        Result := True;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RECONHECER SE JA EXISTE O CODIGO DA MARCA OU NAO
function TfrmPrinc.temMarca(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := False
  else
  begin
    try
      queryTemp := TSQLQuery.Create(nil);
      frmPrinc.conDestino.Open;
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT m.codi FROM marca m WHERE m.codi = :PCODI');
      queryTemp.ParamByName('PCODI').AsString := Codigo;
      queryTemp.Open;

      if queryTemp.IsEmpty = True then
        Result := False
      else
        Result := True;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//FUNÇÃO PARA RETORNAR CODIGO DO CLIE/FORN PELO NOME
function TfrmPrinc.getCodiClieForn(clieforn: String): Integer;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := 0
  else
  begin
    try
      try
        queryTemp := TSQLQuery.Create(nil);
        frmPrinc.conDestino.Open;
        queryTemp.SQLConnection := frmPrinc.conDestino;
        queryTemp.SQL.Clear;
        //queryTemp.SQL.Add('select c.codi from clieforn c where c.nome = :PNOME');
        //queryTemp.ParamByName('PNOME').AsString := clieforn;
        queryTemp.CommandText := 'SELECT c.codi FROM clieforn c WHERE c.nome = ' + clieforn + ';';
        queryTemp.ExecSQL;
        queryTemp.Open;

        Result := queryTemp.FieldByName('CODI').AsInteger;

      except
        on e: exception do
        begin
          ShowMessage('Erro SQL: '+e.message+sLineBreak+queryTemp.CommandText);
        end;
      end;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
  end;
end;


//Função para encontrar o codigo do produto quando é Update e não usa CODI
function TfrmPrinc.getProdCodUpdate(line: Integer) : String;
var
  i,j: Integer;
  str, temp: string;
begin
  str := '';
  for j := 0 to colUpdateCount-1 do begin
    //Testa se alguma das colunas marcadas como update pode ser substituida pelo codigo
    if LowerCase(colUpdate[j]) = 'refe' then begin
      i := BuscaColuna(StringGrid1,'refe');
      if i <> -1 then begin
        temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,line]));
        temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);

        if str <> '' then str := str + ' AND ';
        str := str + 'refe = '+QuotedStr(temp);
      end;
    end;
    if LowerCase(colUpdate[j]) = 'refe_original' then begin
      i := BuscaColuna(StringGrid1,'refe_original');
      if i <> -1 then begin
        temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,line]));
        temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);

        if str <> '' then str := str + ' AND ';
        str := str + 'refe_original = '+QuotedStr(temp);
      end;
    end;
    if LowerCase(colUpdate[j]) = 'codi_barra' then begin
      i := BuscaColuna(StringGrid1,'codi_barra');
      if i <> -1 then begin
        temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,line]));
        temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);

        if str <> '' then str := str + ' AND ';
        str := str + 'codi_barra = '+QuotedStr(temp);
      end;
    end;
    if LowerCase(colUpdate[j]) = 'codi_barra_com' then begin
      i := BuscaColuna(StringGrid1,'codi_barra_com');
      if i <> -1 then begin
        temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,line]));
        temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);

        if str <> '' then str := str + ' AND ';
        str := str + 'codi_barra_com = '+QuotedStr(temp);
      end;
    end;
  end;
  if str = '' then Result := ''
  else Result := '(SELECT codi FROM prod WHERE ' +str+ ' )';
end;


//FIM DAS FUNÇÕES ANTES DA IMPORTAÇÃO
//------------------------------------------------------------------------------


//IMPORTAR DADOS
procedure TfrmPrinc.ButImportClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  temp, temp2, max: String;
  i, k: integer;
  fileTXT: TextFile;
  lista: TStringList;

begin
  //Incicialmente, testar se existem colunas com mesmo nome
  if checkCol(StringGrid1)=False then Exit;
  //Se não tiver colunas iguais, segue importação.

  //Verificar se não selecionou um tipo de imporatação, finaliza
  if SelectImport.Text = 'Tipo de Importação' then begin
    ShowMessage('Selecione um tipo de Importação.');
    Exit;
  end;

  //Verificar se selecionou um destino para importar
  if DBPath.Text = 'Caminho do destino (TXT, SQL) ou DADOS (.FDB)' then begin
    ShowMessage('Selecione um destino para a Importação.');
    Exit;
  end;

  //Se for CUSTOM, verificar se nome da tabela esta preenchido
  if (SelectImport.Text = 'CUSTOM') and (edtTableName.Text = '') then
  begin
    ShowMessage('Informe o nome da Tabela para importação CUSTOM.');
    Exit;
  end;

  //Coluna da empresa: EMPRESA e EMPR são aceitas em todos os tipos de importação
  //(Clie/Forn usa EMPRESA; Produtos, Títulos, Grupos e SubGrupos usam EMPR)
  if (SelectImport.Text <> 'CUSTOM') and (SelectImport.Text <> 'Marcas') and (SelectImport.Text <> 'Transportadoras') then
  begin
    if SelectImport.Text = 'Clie/Forn' then begin temp := 'EMPRESA'; temp2 := 'EMPR'; end
    else begin temp := 'EMPR'; temp2 := 'EMPRESA'; end;
    if (BuscaColuna(StringGrid1,temp) <> -1) and (BuscaColuna(StringGrid1,temp2) <> -1) then
    begin
      ShowMessage('Utilize somente uma coluna de empresa (EMPRESA ou EMPR), não as duas.');
      Exit;
    end;
    //Sinônimo: passa a coluna para o nome usado por este tipo de importação
    i := BuscaColuna(StringGrid1,temp2);
    if i <> -1 then
    begin
      StringGrid1.Cells[i,0] := temp;
      for k := 0 to colUpdateCount-1 do
        if UpperCase(colUpdate[k]) = temp2 then colUpdate[k] := temp;
    end;
  end;

  //Colunas que não pertencem ao tipo de importação (ex.: nome digitado errado) não são importadas:
  //avisa quais são e pergunta se deseja continuar
  if SelectImport.Text <> 'CUSTOM' then
  begin
    lista := ColunasAceitas(SelectImport.Text);
    try
      temp := '';
      for i := 1 to StringGrid1.ColCount-1 do
        if (Trim(StringGrid1.Cells[i,0]) <> '') and (lista.IndexOf(UpperCase(Trim(StringGrid1.Cells[i,0]))) = -1) then
          temp := temp + '   ' + Trim(StringGrid1.Cells[i,0]) + sLineBreak;
    finally
      lista.Free;
    end;
    if temp <> '' then
      if Mensagem('As colunas abaixo não são reconhecidas para ' + SelectImport.Text + ' e NÃO serão importadas:' + sLineBreak + sLineBreak +
                  temp + sLineBreak + 'Confira o nome das colunas em Carregar > Colunas.' + sLineBreak + 'Deseja continuar a importação?',
                  mtWarning, [mbYes, mbNo], ['Continuar', 'Cancelar'], 'Colunas não reconhecidas') <> mrYes then
        Exit;
  end;

  //Se for Clie/Forn, não pode usar CPF_CNPJ junto com as colunas CPF ou CNPJ
  if (SelectImport.Text = 'Clie/Forn') and (BuscaColuna(StringGrid1,'cpf_cnpj') <> -1) and
     ((BuscaColuna(StringGrid1,'cpf') <> -1) or (BuscaColuna(StringGrid1,'cnpj') <> -1)) then
  begin
    ShowMessage('Utilize a coluna CPF_CNPJ ou as colunas CPF/CNPJ, não as duas formas juntas.');
    Exit;
  end;

  //Clie/Forn: referência comercial não pode ser chave de Update
  if (SelectImport.Text = 'Clie/Forn') and (VerificaUpdate('refe_come') = 1) then
  begin
    ShowMessage('A coluna REFE_COME não pode ser marcada como Update (chave). Desmarque a coluna para continuar.');
    Exit;
  end;

  //Clie/Forn: inclusão exige o nome (razão social), na coluna NOME ou NOME_FANT
  if (SelectImport.Text = 'Clie/Forn') and (colUpdateCount <= 0) and
     (BuscaColuna(StringGrid1,'nome') = -1) and (BuscaColuna(StringGrid1,'nome_fant') = -1) then
  begin
    ShowMessage('Informe a coluna NOME (razão social) para incluir clientes e fornecedores.');
    Exit;
  end;

  //Transportadoras: inclusão exige a coluna NOME; a cidade vem por CIDA + UF ou por CODI_CIDA, não pelas duas
  if (SelectImport.Text = 'Transportadoras') and (colUpdateCount <= 0) and (BuscaColuna(StringGrid1,'nome') = -1) then
  begin
    ShowMessage('Informe a coluna NOME para incluir transportadoras.');
    Exit;
  end;
  if (SelectImport.Text = 'Transportadoras') and (BuscaColuna(StringGrid1,'codi_cida') <> -1) and (BuscaColuna(StringGrid1,'cida') <> -1) then
  begin
    ShowMessage('Informe a cidade pelas colunas CIDA e UF ou pela coluna CODI_CIDA, não pelas duas formas juntas.');
    Exit;
  end;
  if (SelectImport.Text = 'Transportadoras') and (BuscaColuna(StringGrid1,'cpf_cnpj') <> -1) and
     ((BuscaColuna(StringGrid1,'cpf') <> -1) or (BuscaColuna(StringGrid1,'cnpj') <> -1)) then
  begin
    ShowMessage('Utilize a coluna CPF_CNPJ ou as colunas CPF/CNPJ, não as duas formas juntas.');
    Exit;
  end;

  if (SelectImport.Text = 'Transportadoras') and (BuscaColuna(StringGrid1,'obs') <> -1) and (BuscaColuna(StringGrid1,'observacao') <> -1) then
  begin
    ShowMessage('Utilize somente uma coluna de observação (OBS ou OBSERVACAO), não as duas.');
    Exit;
  end;

  //Se for Produtos, não pode usar CSOSN/CST junto com as colunas separadas (estadual/interestadual)
  //CSOSN: cada campo só pode vir de uma coluna (CSOSN = os 5 campos, CSOSN_ESTA = os 3 estaduais,
  //CSOSN_INTER = os 2 interestaduais e uma coluna para cada campo)
  if (SelectImport.Text = 'Produtos') and
     ((ColunasCST(['csosn','csosn_esta','csosn_estadual']) > 1) or
      (ColunasCST(['csosn','csosn_esta','csosn_esta_cf']) > 1) or
      (ColunasCST(['csosn','csosn_esta','csosn_esta_produtor']) > 1) or
      (ColunasCST(['csosn','csosn_inter','csosn_interestadual']) > 1) or
      (ColunasCST(['csosn','csosn_inter','csosn_inter_cf']) > 1)) then
  begin
    ShowMessage('Há mais de uma coluna de CSOSN para o mesmo campo.' + sLineBreak + sLineBreak +
                'CSOSN vale para os 5 campos, CSOSN_ESTA para os 3 estaduais e CSOSN_INTER para os 2 interestaduais.' + sLineBreak +
                'Use CSOSN_ESTADUAL, CSOSN_ESTA_CF, CSOSN_ESTA_PRODUTOR, CSOSN_INTERESTADUAL e CSOSN_INTER_CF somente nos campos que as outras colunas não cobrem.');
    Exit;
  end;
  //CST de ICMS: cada campo só pode vir de uma coluna (CST = os 5 campos, CST_ESTA = os 3 estaduais,
  //CST_INTER = os 2 interestaduais e uma coluna para cada campo)
  if (SelectImport.Text = 'Produtos') and
     ((ColunasCST(['cst','cst_esta','cst_estadual']) > 1) or
      (ColunasCST(['cst','cst_esta','cst_esta_cf']) > 1) or
      (ColunasCST(['cst','cst_esta','cst_esta_produtor']) > 1) or
      (ColunasCST(['cst','cst_inter','cst_interestadual']) > 1) or
      (ColunasCST(['cst','cst_inter','cst_inter_cf']) > 1)) then
  begin
    ShowMessage('Há mais de uma coluna de CST para o mesmo campo.' + sLineBreak + sLineBreak +
                'CST vale para os 5 campos, CST_ESTA para os 3 estaduais e CST_INTER para os 2 interestaduais.' + sLineBreak +
                'Use CST_ESTADUAL, CST_ESTA_CF, CST_ESTA_PRODUTOR, CST_INTERESTADUAL e CST_INTER_CF somente nos campos que as outras colunas não cobrem.');
    Exit;
  end;

  //Alíquota de ICMS: ALIQ_ICMS vale para os 3 campos estaduais, não pode vir junto com ALIQ_ICMS_ESTA, ALIQ_ICMS_ESTA_CF ou ALIQ_ICMS_ESTA_PRODUTOR
  if (SelectImport.Text = 'Produtos') and (BuscaColuna(StringGrid1,'aliq_icms') <> -1) and
     (ColunasCST(['aliq_icms_esta','aliq_icms_esta_cf','aliq_icms_esta_produtor']) > 0) then
  begin
    ShowMessage('Há mais de uma coluna de alíquota de ICMS para o mesmo campo.' + sLineBreak + sLineBreak +
                'ALIQ_ICMS vale para os 3 campos estaduais (estadual, consumidor final e produtor).' + sLineBreak +
                'Use ALIQ_ICMS_ESTA, ALIQ_ICMS_ESTA_CF e ALIQ_ICMS_ESTA_PRODUTOR somente sem a coluna ALIQ_ICMS.');
    Exit;
  end;
  //Redução: REDU vale para os 3 campos estaduais, não pode vir junto com REDU_ESTA, REDU_ESTA_CF ou REDU_ESTA_PRODUTOR
  if (SelectImport.Text = 'Produtos') and (BuscaColuna(StringGrid1,'redu') <> -1) and
     (ColunasCST(['redu_esta','redu_esta_cf','redu_esta_produtor']) > 0) then
  begin
    ShowMessage('Há mais de uma coluna de redução para o mesmo campo.' + sLineBreak + sLineBreak +
                'REDU vale para os 3 campos estaduais (estadual, consumidor final e produtor).' + sLineBreak +
                'Use REDU_ESTA, REDU_ESTA_CF e REDU_ESTA_PRODUTOR somente sem a coluna REDU.');
    Exit;
  end;

  //Títulos: colunas obrigatórias (cliente/fornecedor, vencimento e valor ou saldo)
  if (SelectImport.Text = 'Títulos a Pagar') or (SelectImport.Text = 'Títulos a Receber') then
  begin
    if SelectImport.Text = 'Títulos a Pagar' then temp := 'forn' else temp := 'clie';
    if BuscaColuna(StringGrid1,temp) = -1 then
    begin
      ShowMessage('Informe a coluna ' + UpperCase(temp) + ' (código ou nome do ' +
        IfThen(temp = 'forn', 'fornecedor', 'cliente') + ') para importar ' + SelectImport.Text + '.');
      Exit;
    end;
    if BuscaColuna(StringGrid1,'venc') = -1 then
    begin
      ShowMessage('Informe a coluna VENC (vencimento) para importar ' + SelectImport.Text + '.');
      Exit;
    end;
    if (BuscaColuna(StringGrid1,'valo') = -1) and (BuscaColuna(StringGrid1,'sald') = -1) then
    begin
      ShowMessage('Informe a coluna VALO (valor) ou SALD (saldo) para importar ' + SelectImport.Text + '.');
      Exit;
    end;
    //Reinicia o controle de códigos e de cadastros da importação
    if SelectImport.Text = 'Títulos a Pagar' then cImportaTituP.Inicia
    else cImportaTituR.Inicia;
  end;

  //Grades: PROD e LINHA são obrigatórias; QTD, EST e MAX são o saldo final de cada combinação (não usa Update)
  if SelectImport.Text = 'Grades' then
  begin
    if colUpdateCount > 0 then
    begin
      ShowMessage('A importação de Grades não usa colunas Update.' + sLineBreak + sLineBreak +
                  'QTD, EST e MAX são o saldo final de cada combinação e o importador lança só a diferença.' + sLineBreak +
                  'Desmarque as colunas Update para continuar.');
      Exit;
    end;
    if BuscaColuna(StringGrid1,'prod') = -1 then
    begin
      ShowMessage('Informe a coluna PROD (código do produto) para importar Grades.');
      Exit;
    end;
    if BuscaColuna(StringGrid1,'linha') = -1 then
    begin
      ShowMessage('Informe a coluna LINHA (ex.: tamanho) para importar Grades.');
      Exit;
    end;
    if (BuscaColuna(StringGrid1,'qtd') <> -1) and ((BuscaColuna(StringGrid1,'est') <> -1) or (BuscaColuna(StringGrid1,'max') <> -1)) then
    begin
      ShowMessage('Utilize a coluna QTD (estoque fiscal e físico) ou as colunas EST (fiscal) e MAX (físico), não as duas formas juntas.');
      Exit;
    end;
  end;

  //Produtos: verifica se a importação usa o generator (linhas sem CODI)
  if SelectImport.Text = 'Produtos' then
    cImportaProduto.Inicia(StringGrid1, StrToIntDef(StartLine.Text, 1));

  //Backup antes da primeira importação da empresa (pasta de backup ainda sem cópias)
  if not BackupAntesDaImportacao(NomeTipoBackup(SelectImport.Text, edtTableName.Text, colUpdateCount > 0)) then
    Exit;


  //Status se esta OK ou se tem erro, setado como OK
  status := 1;

  //Flag se o botão continuar foi clicado
  butContinue := 0;

  try
    try
      //Criar tela de loading
      frmImportando.Show;
      frmImportando.Label2.Font.Color := CorTexto;

      //Carregar inicio da StartLine
      if not IsNumeric(StartLine.Text) then begin
        StartLine.Text := '1';
      end;
      if StrToInt(StartLine.Text) <= 0 then begin
        StartLine.Text := '1';
      end;
      if StrToInt(StartLine.Text) > StringGrid1.RowCount then begin
        ShowMessage('Inicio maior que o número máximo de linhas: '+IntToStr(StringGrid1.RowCount));
        Exit;
      end;

      //Buscar quantidade de empresas
      i:=BuscaColuna(StringGrid1,'empr');
      if (i<>-1) then
      begin
        qtdEmpr := quantidadeEmpresas(i);
      end
      else qtdEmpr := 1;

      //Produtos: criar a PROD_AJUS antes das movimentações na ITENS,
      //assim o NUME dos itens (gen_id(gen_prod_ajus_id,0)) é o código deste ajuste
      //No Update só cria se alguma coluna de estoque gerar movimentação
      //(coluna de estoque marcada como Update é apenas condição, não movimenta)
      if (SelectImport.Text = 'Produtos') and
         ((colUpdateCount <= 0) or
          ((BuscaColuna(StringGrid1,'qtd') <> -1) and (VerificaUpdate('qtd') = 0)) or
          ((BuscaColuna(StringGrid1,'est') <> -1) and (VerificaUpdate('est') = 0)) or
          ((BuscaColuna(StringGrid1,'max') <> -1) and (VerificaUpdate('max') = 0))) then
      begin
        frmImportando.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
        if UpperCase( ExtractFileExt(DBPath.Text) ) = '.FDB' then begin
          SQL := TSQLDataSet.Create(Nil);
          conDestino.Open;
          try
            SQL.SQLConnection := conDestino;
            SQL.CommandText := 'INSERT INTO prod_ajus (codi,data,obs) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE,''IMPORTADO PELO IMPORTA EXCEL'');';
            SQL.ExecSQL;
          finally
            SQL.Free;
            conDestino.Close;
          end;
        end
        else if (UpperCase( ExtractFileExt(DBPath.Text) ) = '.TXT') or
                (UpperCase( ExtractFileExt(DBPath.Text) ) = '.SQL')
        then begin
          AssignFile(fileTXT, DBPath.Text);
          if not FileExists(DBPath.Text) then ReWrite(fileTXT)
          else append(fileTXT);
          try
            WriteLn(fileTXT, '----------Comandos da PROD_AJUS----------');
            WriteLn(fileTXT, 'INSERT INTO prod_ajus (codi,data,obs) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE,''IMPORTADO PELO IMPORTA EXCEL'');');
            WriteLn(fileTXT, 'COMMIT WORK;');
          finally
            CloseFile(fileTXT);
          end;
        end;
      end;

      //Grades: generators e controle do ajuste de estoque da importação
      if SelectImport.Text = 'Grades' then
        cImportaGrade.Inicia;

      for k := StrToInt(StartLine.Text) to StringGrid1.RowCount-1 do
      begin

        //Atualizar StartLine
        StartLine.Text := IntToStr(k);

        //Se clicou em cancelar, quebra o laço das linhas e finaliza importação.
        //if frmImportando.Active=False then break; - Usando Active, se minimizar a tela cancela.
        if frmImportando.Visible=False then break;


        //Atualizar Status
        frmImportando.atualizaItens(k,StringGrid1.RowCount-1);

        //----------------------------------------------------------------------
        //----------------------------------------------------------------------
        //Importar Clientes e Fornecedores
        if SelectImport.Text = 'Clie/Forn' then
        begin
          cImportaClieForn.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end
        //----------------------------------------------------------------------------
        //Importar Produtos
        else if SelectImport.Text = 'Produtos' then
        begin
          cImportaProduto.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Grupos
        else if SelectImport.Text = 'Grupos' then
        begin
          cImportaGrupo.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar SubGrupos
        else if SelectImport.Text = 'SubGrupos' then
        begin
          cImportaSubGrupo.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Marcas
        else if SelectImport.Text = 'Marcas' then
        begin
          cImportaMarca.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Transportadoras
        else if SelectImport.Text = 'Transportadoras' then
        begin
          cImportaTran.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Grades
        else if SelectImport.Text = 'Grades' then
        begin
          cImportaGrade.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Títulos a Pagar
        else if SelectImport.Text = 'Títulos a Pagar' then
        begin
          cImportaTituP.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //Importar Títulos a Receber
        else if SelectImport.Text = 'Títulos a Receber' then
        begin
          cImportaTituR.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end

        //----------------------------------------------------------------------------
        //OUTRAS OPÇÕES DE IMPORTAÇÃO COLOCAR AQUI


        //----------------------------------------------------------------------------
        //Importar CUSTOM
        else if SelectImport.Text = 'CUSTOM' then
        begin
          cImportaCUSTOM.ImportaRegistro(k, StringGrid1);

          //Verificar se deu erro
          if status = 0 then
            Break; //Quebra o for
        end
        ;


      //Fim do For das Linhas
      end;

      //------------------------------------------------------------------------------
      //COMANDOS PÓS IMPORTAÇÃO

      //Usando banco de dados FDB
      if UpperCase( ExtractFileExt(DBPath.Text) ) = '.FDB' then begin
        //Abrir conexoes
        SQL := TSQLDataSet.Create(Nil);
        conDestino.Open;
        SQL.SQLConnection := conDestino;

        if SelectImport.Text='Clie/Forn' then
        begin
          //Arrumar Generator dos Clientes e Fornecedores
          max := querySelect('SELECT MAX(codi) FROM clieforn');
          if StrToInt(max) > 0 then
          begin
            frmImportando.atualizaStatus('Alterando generator do Clie/Forn.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_CLIEFORN_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;
        end

        else if SelectImport.Text='Produtos' then
        begin
          //Arrumar Generator dos Produtos
          max := querySelect('SELECT MAX(codi) FROM prod');
          if StrToInt(max) > 0 then
          begin
            frmImportando.atualizaStatus('Alterar Generator do Produto.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;

          //Se for INSERT
          if colUpdateCount <= 0 then begin
            //Recalcular as margens de toda a base pelos preços (o preço pela MARGEM já foi
            //calculado somente nos produtos da planilha, sem alterar o preço a vista informado)
            frmImportando.atualizaStatus('Ajustando MARGENS.');
            SQL.CommandText := SQL_MARGEM1;
            SQL.ExecSQL;
            SQL.CommandText := SQL_MARGEM2;
            SQL.ExecSQL;
            //PROD_AJUS é criada antes da importação dos registros
          end
          //Se for UPDATE
          else begin
            //PROD_AJUS é criada antes da importação dos registros

            //Recalcular as margens se a planilha tiver custo, preço ou margem
            //(o preço pela MARGEM já foi calculado somente nos produtos da planilha)
            if (BuscaColuna(StringGrid1,'margem') <> -1) or (BuscaColuna(StringGrid1,'preco_prazo') <> -1) or
               (BuscaColuna(StringGrid1,'preco_vista') <> -1) or (BuscaColuna(StringGrid1,'custo_real') <> -1) or
               (BuscaColuna(StringGrid1,'custo') <> -1) then
            begin
              frmImportando.atualizaStatus('Ajustando MARGENS.');
              SQL.CommandText := SQL_MARGEM1;
              SQL.ExecSQL;
              SQL.CommandText := SQL_MARGEM2;
              SQL.ExecSQL;
            end;
          end;

        end

        else if SelectImport.Text='Grupos' then
        begin
          //Arrumar Generator dos Grupos
          max := querySelect('SELECT MAX(codi) FROM grup_prod');
          if StrToInt(max) > 0 then
          begin
            frmImportando.atualizaStatus('Alterar Generator dos Grupos.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_GRUP_PROD_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;
        end

        else if SelectImport.Text='SubGrupos' then
        begin
          //Arrumar Generator dos SubGrupos
          max := querySelect('SELECT MAX(codi) FROM sub_grup_prod');
          if StrToInt(max) > 0 then
          begin
            frmImportando.atualizaStatus('Alterar Generator dos SubGrupos.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_SUB_GRUP_PROD_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;
        end

        else if SelectImport.Text='Transportadoras' then
        begin
          //Arrumar Generator das Transportadoras
          max := querySelect('SELECT MAX(codi) FROM tran');
          if StrToIntDef(max, 0) > 0 then
          begin
            frmImportando.atualizaStatus('Alterar Generator das Transportadoras.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_TRAN_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;
        end

        else if SelectImport.Text='Marcas' then
        begin
          //Arrumar Generator das MARCAS
          max := querySelect('SELECT MAX(codi) FROM marca');
          if StrToInt(max) > 0 then
          begin
            frmImportando.atualizaStatus('Alterar Generator das Marcas.');
            SQL.CommandText := 'ALTER SEQUENCE GEN_MARCA_ID RESTART WITH ' + max + ';';
            SQL.ExecSQL;
          end;
        end

        else if SelectImport.Text='Grades' then
        begin
          //Completar o estoque da grade (zero) de todos os produtos com grade
          frmImportando.atualizaStatus('Completando o estoque das grades.');
          SQL.CommandText := cImportaGrade.SQLMatriz;
          SQL.ExecSQL;
        end

        ;
        //Fechar conexoes
        SQL.Free;
        conDestino.Close;
      end

      //Usando arquivo TXT ou SQL
      else if (UpperCase( ExtractFileExt(DBPath.Text) ) = '.TXT') or
              (UpperCase( ExtractFileExt(DBPath.Text) ) = '.SQL')
      then begin
        //Carregar arquivo TXT ou SQL
        AssignFile(fileTXT, DBPath.Text);
        if not FileExists(DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        WriteLn(fileTXT, '----------Comandos da PÓS-IMPORTAÇÃO----------');

        if SelectImport.Text='Clie/Forn' then
        begin
          frmImportando.atualizaStatus('Alterando generator do Clie/Forn.');
          WriteLn(fileTXT, 'SELECT gen_id(gen_clieforn_id, ABS((SELECT MAX(CODI) FROM clieforn) - (SELECT gen_id(gen_clieforn_id,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        else if SelectImport.Text='Produtos' then
        begin
          //Se for INSERT
          if colUpdateCount <= 0 then begin
            //Arrumar Generator dos Produtos
            frmImportando.atualizaStatus('Alterar Generator do Produto.');
            WriteLn(fileTXT, 'SELECT gen_id(gen_prod_id, ABS((SELECT MAX(CODI) FROM prod) - (SELECT gen_id(gen_prod_id,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
            WriteLn(fileTXT, 'COMMIT WORK;');

            //Recalcular as margens de toda a base pelos preços (o preço pela MARGEM já foi
            //calculado somente nos produtos da planilha, sem alterar o preço a vista informado)
            frmImportando.atualizaStatus('Ajustando MARGENS.');
            WriteLn(fileTXT, SQL_MARGEM1);
            WriteLn(fileTXT, 'COMMIT WORK;');
            WriteLn(fileTXT, SQL_MARGEM2);
            WriteLn(fileTXT, 'COMMIT WORK;');
            //PROD_AJUS é criada antes da importação dos registros
          end

          //Se for UPDATE
          else begin
            //PROD_AJUS é criada antes da importação dos registros

            //Recalcular as margens se a planilha tiver custo, preço ou margem
            //(o preço pela MARGEM já foi calculado somente nos produtos da planilha)
            if (BuscaColuna(StringGrid1,'margem') <> -1) or (BuscaColuna(StringGrid1,'preco_prazo') <> -1) or
               (BuscaColuna(StringGrid1,'preco_vista') <> -1) or (BuscaColuna(StringGrid1,'custo_real') <> -1) or
               (BuscaColuna(StringGrid1,'custo') <> -1) then
            begin
              frmImportando.atualizaStatus('Ajustando MARGENS.');
              WriteLn(fileTXT, SQL_MARGEM1);
              WriteLn(fileTXT, 'COMMIT WORK;');
              WriteLn(fileTXT, SQL_MARGEM2);
              WriteLn(fileTXT, 'COMMIT WORK;');
            end;
          end;
        end

        else if SelectImport.Text='Grupos' then
        begin
          //Arrumar Generator dos Grupos
          frmImportando.atualizaStatus('Alterar Generator dos Grupos.');
          WriteLn(fileTXT, 'SELECT gen_id(gen_grup_prod_id, abs((SELECT MAX(CODI) FROM grup_prod) - (SELECT gen_id(gen_grup_prod_id,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        else if SelectImport.Text='SubGrupos' then
        begin
          //Arrumar Generator dos SubGrupos
          frmImportando.atualizaStatus('Alterar Generator dos SubGrupos.');
          WriteLn(fileTXT, 'SELECT gen_id(GEN_SUB_GRUP_PROD_ID, ABS((SELECT MAX(CODI) FROM sub_grup_prod) - (SELECT gen_id(GEN_SUB_GRUP_PROD_ID,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        else if SelectImport.Text='Transportadoras' then
        begin
          //Arrumar Generator das Transportadoras
          frmImportando.atualizaStatus('Alterar Generator das Transportadoras.');
          WriteLn(fileTXT, 'SELECT gen_id(GEN_TRAN_ID, ABS((SELECT MAX(CODI) FROM tran) - (SELECT gen_id(GEN_TRAN_ID,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        else if SelectImport.Text='Marcas' then
        begin
          //Arrumar Generator das MARCAS
          frmImportando.atualizaStatus('Alterar Generator das Marcas.');
          WriteLn(fileTXT, 'SELECT gen_id(GEN_MARCA_ID, ABS((SELECT MAX(CODI) FROM marca) - (SELECT gen_id(GEN_MARCA_ID,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        else if SelectImport.Text='Grades' then
        begin
          //Completar o estoque da grade (zero) de todos os produtos com grade
          frmImportando.atualizaStatus('Completando o estoque das grades.');
          WriteLn(fileTXT, cImportaGrade.SQLMatriz);
          WriteLn(fileTXT, 'COMMIT WORK;');
        end

        ;
        //Fechar arquivo
        CloseFile(fileTXT);
      end;

    except
      on e: exception do
      begin
        ShowMessage('Erro Interno: '+e.message+sLineBreak);
        status := 0;
      end;
    end;
  finally
    frmImportando.fim(status);

  end;

  //Backup depois da importação concluída (não pergunta se parou com erro ou se foi cancelada)
  if (status = 1) and frmImportando.Visible then
    BackupDepoisDaImportacao(NomeTipoBackup(SelectImport.Text, edtTableName.Text, colUpdateCount > 0));
end;


procedure TfrmPrinc.BtnDestinoPadraoClick(Sender: TObject);
var
  informacaoPadrao: string;
begin
  informacaoPadrao := DBPath.Text;

  conDestino.Connected := False;
  DBPath.Text := 'C:\ADMERP\DADOS.FDB';
  conDestino.Params.Values['DataBase'] := DBPath.Text;

  try
    conDestino.Connected := True;
    ShowMessage('Conectado ao banco de destino');
  except
    on E: Exception do
    begin
      DBPath.Text := informacaoPadrao;
      conDestino.Params.Values['DataBase'] := informacaoPadrao;
      ShowMessage('Erro ao conectar: ' + E.Message);
    end;
  end;
end;

//Função para salvar StringGrid em um arquivo Excel
function SaveAsExcelFile(stringGrid: TstringGrid; FileName: string): Boolean;
const
  xlWBATWorksheet = -4167;
var
  Row, Col: Integer;
  GridPrevFile: string;
  XLApp, Sheet: OLEVariant;

begin
  Screen.Cursor := crHourGlass;
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.Workbooks.Add(xlWBatWorkSheet);
    Sheet := XLApp.Workbooks[1].WorkSheets[1];

    Sheet.Name := 'String Grid';

    for col := 1 to stringGrid.ColCount - 1 do
    begin
      for row := 0 to stringGrid.RowCount - 1 do
      begin
        try
          Sheet.Cells[row + 1, col].NumberFormat := '@'; //Text
          Sheet.Cells[row + 1, col] := stringGrid.Cells[col, row];
        except
          on E:Exception do
          begin
            Mensagem('Erro gravando Excel. Verifique o arquivo resultante e se estiver com problemas tente salvar em CSV.'+#13+E.Message, mtCustom,[mbOK], ['Ok'], 'Erro salvando Excel.');
            Break;
          end;
        end;
      end;
    end;
    try
      Sheet.Columns.Autofit;
      XLApp.Workbooks[1].SaveAs(FileName);
      Result := True;
    except
      Result := False
    end;
  finally
    Screen.Cursor := crDefault;
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
    end;
  end;
end;


//Função para salvar StringGrid em um arquivo CSV
function SaveAsCSVFile(Grid: TstringGrid; FileName: string):Boolean;
var
  i, j : Integer;
  CSV : TStrings;
  stream : string;

begin
  Result := False;
  //Criar StringList
  CSV := TStringList.Create;

  Screen.Cursor := crHourGlass;
  CSV.Delimiter := ';';

  Try
    for i := 0 to Grid.RowCount - 1 do
    begin
      if Grid.ColCount >= 2 then
      begin
        stream := Grid.Cells[1,i];
      end
      else begin
        Exit;
      end;

      for j := 2 to Grid.ColCount -1 do
      begin
        stream := stream + ';' + Grid.Cells[j,i];
      end;
      CSV.Add(stream);
    end;

    try
      //Salvar no CSV
      CSV.SaveToFile(FileName);
      Result := True;
    except
      Result := False;
    end;
  Finally
    Screen.Cursor := crDefault;
    CSV.Free;
  End;
end;


//Botão Salvar StringGrid em planilha
procedure TfrmPrinc.btnSalvarClick(Sender: TObject);
var
  fileExt: string;

begin
  //Carregar extensão do arquivo
  fileExt := ExtractFileExt(FilePath.Text);

  //Sugerir extensão inicial
  SaveDialog1.Filter := 'EXCEL files (*.xlsx)|*.XLSX|CSV files (*.csv)|*.CSV';
  SaveDialog1.DefaultExt := 'xlsx';
  if (fileExt='.csv') then
  begin
    SaveDialog1.Filter := 'CSV files (*.csv)|*.CSV|EXCEL files (*.xlsx)|*.XLSX|';
    SaveDialog1.DefaultExt := 'csv';
  end;

  if SaveDialog1.Execute then
  begin
    //Carregar extensão do arquivo
    fileExt := LowerCase(ExtractFileExt(SaveDialog1.FileName));
    if fileExt='' then
    begin
      fileExt := '.'+LowerCase(SaveDialog1.DefaultExt);
    end;

    //Salvar arquivo de acordo com a extensão
    if (fileExt='.xls') or (fileExt='.xlsx') then
    begin
      //Salvar StringFrid em Excel
      if SaveAsExcelFile(StringGrid1, SaveDialog1.FileName) then begin
        ShowMessage(SaveDialog1.FileName+sLineBreak+'StringGrid salva com sucesso!');
      end
      else ShowMessage(SaveDialog1.FileName+sLineBreak+'Erro ao salvar StringGrid!');
    end
    else if (fileExt='.csv') then
    begin
      //Salvar StringGrid em CSV
      if SaveAsCSVFile(StringGrid1, SaveDialog1.FileName) then begin
        ShowMessage(SaveDialog1.FileName+sLineBreak+'StringGrid salva com sucesso!');
      end
      else ShowMessage(SaveDialog1.FileName+sLineBreak+'Erro ao salvar StringGrid!');
    end;

  end;
end;


procedure TfrmPrinc.btnTXTClick(Sender: TObject);
begin
  //Sugerir extensão inicial
  SaveDialog1.Filter := 'Text File(*.txt)|*.TXT|SQL File (*.sql)|*.SQL|';
  SaveDialog1.DefaultExt := 'txt';

  if SaveDialog1.Execute then
  begin
    DBPath.Text := SaveDialog1.FileName;
    conDestino.Params.Values['DataBase'] := '';
  end;

end;

//Função para salvar StringGrid em um array de string
function TfrmPrinc.StringGridToArray(Grid: TStringGrid): Integer;
var
  i,j: integer;
begin
  //Redimensionar array
  SetLength(gridTemp,Grid.RowCount);
  for i := 0 to Grid.RowCount-1 do
  begin
    SetLength(gridTemp[i],Grid.ColCount);
  end;

  //Copiar da StringGrid para o array
  for i := 0 to Grid.RowCount-1 do
  begin
    for j := 0 to Grid.ColCount-1 do
    begin
      gridTemp[i,j] := Grid.Cells[j,i];
    end;
  end;

end;


//Mostrar quias as colunas estão disponíveis para importar
procedure TfrmPrinc.mnuColunasClick(Sender: TObject);
begin
  if SelectImport.Text = 'Tipo de Importação' then
  begin
    ShowMessage('Selecione o Tipo de Importação primeiro!');
  end
  else begin
    //Criar tela de colunas
    frmColunas.LabelTipoImp.Caption := SelectImport.Text;
    frmColunas.LabelTipoImp.Left := (frmColunas.Width - frmColunas.LabelTipoImp.Width ) div 2;
    frmColunas.Show;
  end;
end;


//Carregar Cabeçalho de outra tabela nesta tabela
procedure TfrmPrinc.mnuCabecalhoClick(Sender: TObject);
var
  arquivo: string;

begin
  if opnDadosOrigem.Execute then
  begin
    ExtractFilePath(Application.ExeName);
    arquivo :=  opnDadosOrigem.FileName;

    StringGridToArray(StringGrid1);
    XlsHeaderLoad(StringGrid1,arquivo);

  end;
end;


//Quantas das colunas informadas existem na planilha (usado na validação das colunas de CST, CSOSN, alíquota de ICMS e redução)
function TfrmPrinc.ColunasCST(colunas: array of string): Integer;
var
  j: Integer;
begin
  Result := 0;
  for j := Low(colunas) to High(colunas) do
    if BuscaColuna(StringGrid1, colunas[j]) <> -1 then Inc(Result);
end;


//Confirmar antes de apagar os dados da base de destino (os menus Limpar apagam tudo e zeram os generators)
function ConfirmaLimpeza(dados: string): Boolean;
begin
  Result := False;
  if UpperCase(ExtractFileExt(frmPrinc.DBPath.Text)) <> '.FDB' then
  begin
    ShowMessage('Selecione uma base de dados (.FDB) como destino para limpar os dados.');
    Exit;
  end;
  Result := Mensagem('Serão APAGADOS TODOS os dados de ' + dados + ' da base:' + sLineBreak + sLineBreak +
                     frmPrinc.DBPath.Text + sLineBreak + sLineBreak +
                     'Essa operação não pode ser desfeita. Deseja continuar?',
                     mtWarning, [mbYes, mbNo], ['Apagar', 'Cancelar'], 'Limpar dados') = mrYes;
end;


//Limpar dados de clientes e fornecedores do banco
procedure TfrmPrinc.mnuLimpaClieFornClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  tr: TDBXTransaction;
begin
  if not ConfirmaLimpeza('clientes e fornecedores') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  try
    conDestino.Open;
    SQL.SQLConnection := conDestino;

    //Tudo em uma transação: se algum vínculo impedir, nada é alterado
    tr := conDestino.BeginTransaction;
    try
      //Carteiras (TIPO_CART) apontam para o participante da movimentação financeira (base zerada usa o cadastro 1)
      SQL.CommandText := 'UPDATE tipo_cart SET codi_participante_movi_fina = NULL WHERE codi_participante_movi_fina IS NOT NULL;';
      SQL.ExecSQL;

      SQL.CommandText := 'DELETE FROM clieforn;';
      SQL.ExecSQL;

      conDestino.CommitFreeAndNil(tr);
    except
      on e: Exception do
      begin
        conDestino.RollbackFreeAndNil(tr);
        ShowMessage('Não foi possível limpar os clientes e fornecedores. Nenhum dado foi alterado.' + sLineBreak + sLineBreak +
                    'Verifique se há registros vinculados a eles (ex.: títulos a pagar/receber).' + sLineBreak + sLineBreak + e.Message);
        Exit;
      end;
    end;

    SQL.CommandText := 'ALTER SEQUENCE GEN_CLIEFORN_ID RESTART WITH 0;';
    SQL.ExecSQL;

    ShowMessage('Limpado dados de clientes e fornecedores.');
  finally
    //Fechar conexoes
    SQL.Free;
    conDestino.Close;
  end;
end;


//Limpar transportadoras do banco
procedure TfrmPrinc.mnuLimpaTransportadorasClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('transportadoras') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  try
    conDestino.Open;
    SQL.SQLConnection := conDestino;

    //Transportadora usada em MDF-e, CT-e, romaneio ou orçamento não pode ser apagada: nada é alterado
    try
      SQL.CommandText := 'DELETE FROM tran;';
      SQL.ExecSQL;
    except
      on e: Exception do
      begin
        ShowMessage('Não foi possível limpar as transportadoras. Nenhum dado foi alterado.' + sLineBreak + sLineBreak +
                    'Verifique se há registros vinculados a elas (ex.: MDF-e, CT-e, romaneio de carga ou orçamento).' + sLineBreak + sLineBreak + e.Message);
        Exit;
      end;
    end;

    SQL.CommandText := 'ALTER SEQUENCE GEN_TRAN_ID RESTART WITH 0;';
    SQL.ExecSQL;

    ShowMessage('Limpado dados de transportadoras.');
  finally
    //Fechar conexoes
    SQL.Free;
    conDestino.Close;
  end;
end;


//Limpar dados de grupos do banco
procedure TfrmPrinc.mnuLimpaGruposClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('grupos') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM grup_prod;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_grup_prod_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de grupos.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de marcas do banco
procedure TfrmPrinc.mnuLimpaMarcasClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('marcas') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM marca;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_MARCA_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de marcas.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de produtos do banco
//Limpar grades: os produtos são mantidos, só perdem a grade
procedure TfrmPrinc.mnuLimpaGradesClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  tr: TDBXTransaction;
begin
  if not ConfirmaLimpeza('grades (cadastro das grades, estoque por tamanho e cor e os ajustes feitos pela importação de grades; os produtos são mantidos, sem grade)') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  try
    conDestino.Open;
    SQL.SQLConnection := conDestino;

    //Tudo em uma transação: se algum vínculo impedir, nada é alterado
    tr := conDestino.BeginTransaction;
    try
      //Produtos continuam cadastrados, sem grade
      SQL.CommandText := 'UPDATE prod SET grade_grad_id = NULL WHERE grade_grad_id IS NOT NULL;';
      SQL.ExecSQL;
      //Lançamentos de estoque da grade (vendas e pedidos já feitos só perdem o vínculo com o tamanho/cor)
      SQL.CommandText := 'DELETE FROM grade_lancamento;';
      SQL.ExecSQL;
      SQL.CommandText := 'DELETE FROM grad_ajus;';
      SQL.ExecSQL;
      SQL.CommandText := 'DELETE FROM grade_estoque;';
      SQL.ExecSQL;
      SQL.CommandText := 'DELETE FROM grade_prod;';
      SQL.ExecSQL;
      //Ajustes da importação de grades: os ITENS saem junto (trigger da PROD_AJUS) e o estoque total do produto
      //volta ao que era antes da importação
      SQL.CommandText := 'DELETE FROM prod_ajus WHERE CAST(obs AS VARCHAR(200)) = ''' + OBS_AJUSTE + ''';';
      SQL.ExecSQL;
      //Cadastro das grades (tamanhos e cores saem junto)
      SQL.CommandText := 'DELETE FROM grade;';
      SQL.ExecSQL;

      conDestino.CommitFreeAndNil(tr);
    except
      on e: Exception do
      begin
        conDestino.RollbackFreeAndNil(tr);
        ShowMessage('Não foi possível limpar as grades. Nenhum dado foi alterado.' + sLineBreak + sLineBreak +
                    'Verifique se há registros vinculados a elas (ex.: integração com e-commerce).' + sLineBreak + sLineBreak + e.Message);
        Exit;
      end;
    end;

    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_LINHA_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_COLUNA_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_ESTOQUE_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_LANCAMENTO_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_LANCAMENTO_ITENS_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_AJUS_ID RESTART WITH 0;';
    SQL.ExecSQL;
    SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_PROD_ID RESTART WITH 0;';
    SQL.ExecSQL;

    ShowMessage('Limpado dados de grades.');
  finally
    //Fechar conexoes
    SQL.Free;
    conDestino.Close;
  end;
end;

procedure TfrmPrinc.mnuLimpaProdutosClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('produtos (com estoque, ajustes e estoque das grades)') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  //Lançamentos de estoque das grades apontam para o produto (o estoque da grade é apagado junto com o produto)
  SQL.CommandText := 'DELETE FROM grade_lancamento;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM grad_ajus;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM prod;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM prod_ajus;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM itens;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_AJUS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_ESTOQUE_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_LANCAMENTO_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_LANCAMENTO_ITENS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_GRADE_PROD_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_ITENS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_MVA_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ADICIONAIS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_AJUS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_CUSTOS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_FORN_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ICMS_ST_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_TRIBUTOS_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de produtos.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de subgrupos do banco
procedure TfrmPrinc.mnuLimpaSubGruposClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('subgrupos') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM sub_grup_prod;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_sub_grup_prod_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de subgrupos.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de Titulos a pagar do banco
procedure TfrmPrinc.mnuLimpaTituPClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('títulos a pagar') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM titup;';
  SQL.ExecSQL;

  SQL.CommandText := 'DELETE FROM btitup;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de Títulos a Pagar.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de Titulos a receber do banco
procedure TfrmPrinc.mnuLimpaTituRClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  if not ConfirmaLimpeza('títulos a receber') then Exit;

  //Abrir conexoes
  SQL := TSQLDataSet.Create(Nil);
  conDestino.Open;
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM titur;';
  SQL.ExecSQL;

  SQL.CommandText := 'DELETE FROM btitur;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de Títulos a Receber.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


procedure TfrmPrinc.mnuProcurarClick(Sender: TObject);
begin
  frmProcurar.Show;
end;

procedure TfrmPrinc.mnuSubstituirClick(Sender: TObject);
begin
  frmSubstituir.Show;
end;


//Evento ao apertar Botões do Teclado na StringGrid
procedure TfrmPrinc.StringGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i,j,but: integer;
  temp: string;
begin

   //RECONHECER CTRL+H
   if ((Shift = [ssCtrl]) and (Key = $48)) then
   begin
    mnuSubstituir.Click;
   end;

   //RECONHECER CTRL+F
   if ((Shift = [ssCtrl]) and (Key = $46)) then
   begin
    mnuProcurar.Click;
   end;

   //RECONHECER CTRL+D
   if ((Shift = [ssCtrl]) and (Key = $44)) then
   begin
    mnuDividir.Click;
   end;

   //RECONHECER CTRL+Z
   if ((Shift = [ssCtrl]) and (Key = 90)) then
   begin
    if Length(gridTemp) > 1 then
    begin
      StringGrid1.RowCount := Length(gridTemp);
      StringGrid1.ColCount := Length(gridTemp[0]);

      for i := 0 to Length(gridTemp)-1 do
      begin
        for j := 0 to Length(gridTemp[0])-1 do
        begin
          StringGrid1.Cells[j,i] := gridTemp[i,j];
        end;
      end;
    end;
    //Redimensionar colunas
    for i := 0 to StringGrid1.ColCount - 1 do
      AutoSizeCol(StringGrid1, i);
   end;

  //Se apertar 'Del'
  if (Key = VK_DELETE) then
  begin
    but := Mensagem('Deletar linha ou coluna', mtCustom, [mbYes, mbNo],['Linha','Coluna'], 'Deletar');
    if (but = 6) then
    begin
      //ShowMessage('Deletar Linha');
      StringGridToArray(StringGrid1);
      DeleteRow(StringGrid1, StringGrid1.Row);
    end
    else if (but = 7) then
    begin
      //ShowMessage('Deletar Coluna');
      StringGridToArray(StringGrid1);
      DeleteCol(StringGrid1, StringGrid1.Col);
      //Redimensionar colunas
      for i := 0 to StringGrid1.ColCount - 1 do
        AutoSizeCol(StringGrid1, i);
    end;
  end;

  //Teclas para voltar as linhas fixas
  if StringGrid1.Col=0 then
  begin
    if (Key=VK_TAB) or
     (Key=VK_RETURN) or
     (Key=VK_ESCAPE) or
     (Key=VK_LEFT) or
     (Key=VK_RIGHT) then
    begin
      if StringGrid1.FixedRows=0 then StringGrid1.Col:=1;
      StringGrid1.FixedRows:=1;
    end;
  end;

  //Teclas para voltar as colunas fixas
  if StringGrid1.Row=0 then
  begin
    if (Key=VK_TAB) or
     (Key=VK_RETURN) or
     (Key=VK_ESCAPE) or
     (Key=VK_LEFT) or
     (Key=VK_RIGHT) then
    begin
      if StringGrid1.FixedCols=0 then StringGrid1.Row:=1;
      StringGrid1.FixedCols:=1;
    end;
  end;

  //Inserir coluna
  if (Key=VK_F1) then
  begin
    StringGridToArray(StringGrid1);
    InsertCol(StringGrid1);

    //Redimensionar colunas
    for i := 0 to StringGrid1.ColCount - 1 do
      AutoSizeCol(StringGrid1, i);
  end;

  //Inserir Linha
  if (Key=VK_F3) then
  begin
    StringGridToArray(StringGrid1);
    InsertRow(StringGrid1);
  end;
end;


//Botão Adicionar Coluna na StringGrid
procedure TfrmPrinc.mnuAdicionarColunaClick(Sender: TObject);
var
  i: Integer;
begin
  StringGridToArray(StringGrid1);
  InsertCol(StringGrid1);

  //Redimensionar colunas
  for i := 0 to StringGrid1.ColCount - 1 do
    AutoSizeCol(StringGrid1, i);
end;


//Botão Adicionar Linha na StringGrid
procedure TfrmPrinc.mnuAdicionarLinhaClick(Sender: TObject);
begin
  StringGridToArray(StringGrid1);
  InsertRow(StringGrid1);
end;


//Botão Deletar Coluna na StringGrid
procedure TfrmPrinc.mnuDeletarColunaClick(Sender: TObject);
var
  i: Integer;
begin
  StringGridToArray(StringGrid1);
  DeleteCol(StringGrid1, StringGrid1.Col);
  //Redimensionar colunas
  for i := 0 to StringGrid1.ColCount - 1 do
    AutoSizeCol(StringGrid1, i);
end;


//Botão Deletar Linha na StringGrid
procedure TfrmPrinc.mnuDeletarLinhaClick(Sender: TObject);
begin
  StringGridToArray(StringGrid1);
  DeleteRow(StringGrid1, StringGrid1.Row);
end;


procedure TfrmPrinc.mnuDividirClick(Sender: TObject);
begin
  frmDividir.Show;
end;

//Reconhecer Right Click na celula
procedure TfrmPrinc.StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  PMouse: TPoint;
  i, j, k, Col, Row, but, but2: integer;
  valor,temp: string;
begin
  //Right Click
  if Button = mbRight then
  begin
    //Testar qual coluna clicou
    PMouse := Mouse.CursorPos;
    PMouse := StringGrid1.ScreenToClient(PMouse);
    StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

    //Se for uma coluna não fixa
    if (((Col<>0) and (Row<>0)) or ((Col=0) and (Row<>0))) then
    begin
      //Setar focus
      StringGrid1.Col := Col;
      StringGrid1.Row := Row;

      but := Mensagem('Preencher coluna', mtCustom, [mbYes, mbNo],['Copiar','Serie'], 'Preencher');
      if (but = 6) then
      begin
        //ShowMessage('Copiar valor');
        StringGridToArray(StringGrid1);
        valor := StringGrid1.Cells[Col,Row];
        for i := 1 to StringGrid1.RowCount do
        begin
          StringGrid1.Cells[Col,i] := valor;
        end;
      end
      else if (but = 7) then
      begin
        //ShowMessage('Preencher serie');
        StringGridToArray(StringGrid1);
        valor := StringGrid1.Cells[Col,Row];
        j:=1;
        for i := Row+1 to StringGrid1.RowCount do
        begin
          StringGrid1.Cells[Col,i] := IntToStr(StrToInt(valor)+j);
          j:=j+1;
        end;
      end;
    end
    //Se for uma fixa pergunta se deseja mesclar ou copiar coluna
    else if Row=0 then
    begin
      but := Mensagem('Mesclar ou Copiar coluna ou marcar coluna como Update', mtCustom, [mbYes, mbNo, mbIgnore],['Mesclar','Copiar','Update'], 'Mesclar - Copiar - Marcar Update');
      if (but = 6) then
      begin
        //ShowMessage('Mesclar Coluna');
        but2 := Mensagem('Mesclar com a coluna', mtCustom, [mbYes, mbNo],['À Esquerda','À Direita'], 'Mesclar colunas');
        if (but2 = 6) then
        begin
          if Col=0 then
          begin
            ShowMessage('Não existem mais colunas à esquerda.');
          end
          else begin
            StringGridToArray(StringGrid1);
            for i := 0 to StringGrid1.RowCount-1 do
            begin
              StringGrid1.Cells[Col-1,i] := StringGrid1.Cells[Col-1,i] + StringGrid1.Cells[Col,i];
            end;
            DeleteCol(StringGrid1, Col);
          end;
          //Redimensionar colunas
          for i := 0 to StringGrid1.ColCount - 1 do
            AutoSizeCol(StringGrid1, i);
        end
        else if (but2 = 7) then
        begin
          if Col=StringGrid1.ColCount then
          begin
            ShowMessage('Não existem mais colunas à direita.');
          end
          else begin
            StringGridToArray(StringGrid1);
            for i := 0 to StringGrid1.RowCount-1 do
            begin
              StringGrid1.Cells[Col+1,i] := StringGrid1.Cells[Col,i] + StringGrid1.Cells[Col+1,i];
            end;
            DeleteCol(StringGrid1, Col);
          end;
        end;
        //Redimensionar colunas
        for i := 0 to StringGrid1.ColCount - 1 do
          AutoSizeCol(StringGrid1, i);
      end
      else if (but = 7) then
      begin
        //ShowMessage('Copiar Coluna');
        StringGridToArray(StringGrid1);
        StringGrid1.ColCount := StringGrid1.ColCount + 1;
        i:= StringGrid1.ColCount;
        while i>Col do
        begin
          for j := 0 to StringGrid1.RowCount do
          begin
            temp := StringGrid1.Cells[i,j];
            StringGrid1.Cells[i,j] := StringGrid1.Cells[i-1,j];
          end;
          i:= i-1;
        end;
      end
      else if (but = 5) then
      begin
        //ShowMessage('Marcar coluna como Update');

        //Testa se selecionou um tipo de importação
        if SelectImport.Text = 'Tipo de Importação' then begin
          ShowMessage('Selecione um tipo de Importação.');
          Exit;
        end;

        //Verificar se ja existe nos marcados como Update
        j := -1; //j recebera a posicao encontrada
        for i := 0 to colUpdateCount-1 do begin
          temp := StringGrid1.Cells[Col,0];
          temp := colUpdate[i];
          if colUpdate[i] = StringGrid1.Cells[Col,0] then begin
            j := i; //Achou
            Break;
          end;
        end;
        //Se achou deve remover dos Updates
        if j <> -1 then begin
          for i := j+1 to colUpdateCount-1 do begin
            colUpdate[i-1] := colUpdate[i];
          end;
          colUpdateCount := colUpdateCount - 1;
          SetLength(colUpdate,colUpdateCount);
        end
        //Se não achou adiciona aos Updates
        else begin
          //Testar se é uma coluna válida para adicionar nos updates
          k := 0; //K é um flag, se ficar 0 depois do FOR da EXIT
          frmColunas.LabelTipoImp.Caption := SelectImport.Text;
          frmColunas.MostrarColunas;
          for i := 0 to frmColunas.ListColunas.Count-1 do begin
            if StringGrid1.Cells[Col,0] = '' then
              Continue;
            if LowerCase(frmColunas.ListColunas.Items[i]) = LowerCase(StringGrid1.Cells[Col,0]) then begin
              k := 1;
              Break;
            end;
          end;
          if (k = 0) and (SelectImport.Text <> 'CUSTOM') then begin
            ShowMessage('Coluna não pode ser usada como Update.');
            Exit;
          end;

          colUpdateCount := colUpdateCount + 1;
          SetLength(colUpdate,colUpdateCount);
          colUpdate[colUpdateCount-1] := StringGrid1.Cells[Col,0];
        end;
        //Preencher a label
        lblColUpdate.Caption := 'Colunas Update:';
        for i := 0 to colUpdateCount-1 do begin
          lblColUpdate.Caption := lblColUpdate.Caption + ' '+colUpdate[i];
        end;
        //Aciona visibilidade da label se tem pelo menos 1 valor nos updates
        if colUpdateCount > 0 then lblColUpdate.Visible := True
        else lblColUpdate.Visible := False;
        //Atualizar coluna
        StringGrid1DrawCell(StringGrid1, Col, Row, StringGrid1.CellRect(Col,Row), [gdFixed,gdSelected,gdFocused,gdRowSelected,gdHotTrack,gdPressed]);
      end;

      //Redimensionar colunas
      for i := 0 to StringGrid1.ColCount - 1 do
        AutoSizeCol(StringGrid1, i);
    end;
  end
  else //Left Click
  if Button = mbLeft then
  begin
    //Testar qual coluna clicou
    PMouse := Mouse.CursorPos;
    PMouse := StringGrid1.ScreenToClient(PMouse);
    StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

    //Voltar as celulas fixas após clicar fora
    if Row<>0 then
    begin
      StringGrid1.FixedRows:=1;
      //Setar focus
      StringGrid1.Row := Row;
    end;
    if Col<>0 then
    begin
      StringGrid1.FixedCols:=1;
      //Setar focus
      StringGrid1.Col := Col;
    end;
  end;
end;


procedure TfrmPrinc.SelectImportChange(Sender: TObject);
begin
  if SelectImport.Text = 'CUSTOM' then
  begin
    lblTableName.Visible := True;
    edtTableName.Visible := True;
  end
  else
  begin
    lblTableName.Visible := False;
    edtTableName.Visible := False;
  end;
end;

procedure TfrmPrinc.StringGrid1DblClick(Sender: TObject);
var
  PMouse: TPoint;
  Col, Row: integer;
begin

  //Desabilitar celulas fixas ao dar dois cliques
  PMouse := Mouse.CursorPos;
  PMouse := StringGrid1.ScreenToClient(PMouse);

  StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

  if Row=0 then
  begin
    StringGrid1.FixedRows:=0;
    //Setar focus
    StringGrid1.Col := Col;
    StringGrid1.Row := Row;
  end;
  if Col=0 then
  begin
    StringGrid1.FixedCols:=0;
    //Setar focus
    StringGrid1.Col := Col;
    StringGrid1.Row := Row;
  end;

end;


procedure TfrmPrinc.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  i, CellLeftMargin, CellTopMargin: Integer;
begin
  with (Sender as TStringGrid) do
  begin
    //Cores do tema escolhido (claro: faixas azul e amarelo)
    Canvas.Font.Color := CorTexto;
    Canvas.Brush.Color := CorCabecalho;
    // Don't change color for first Column, first row
    if (ARow <> 0) and (ACol <> 0) then begin
      // Draw the Band
      Canvas.Brush.Color := CorFaixa(ARow);
    end;
    for i := 0 to colUpdateCount-1 do begin
      if StringGrid1.Cells[Acol,Arow] = colUpdate[i] then
        Canvas.Font.Color := CorUpdate;
    end;
    Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, cells[acol, arow]);
    Canvas.FrameRect(Rect);
  end;
end;


//Menu Tema: temas disponíveis (claro do Windows e os embutidos no executável)
procedure TfrmPrinc.FormCreate(Sender: TObject);
begin
  MontaMenuTema(mnuTema, mnuTemaItemClick);
end;

//Troca o tema na hora e salva a escolha para as próximas vezes
procedure TfrmPrinc.mnuTemaItemClick(Sender: TObject);
begin
  AplicaTema((Sender as TMenuItem).Hint);
  MontaMenuTema(mnuTema, mnuTemaItemClick);
  StringGrid1.Invalidate;
end;

//Menu Backup: cópia da base de destino na hora
procedure TfrmPrinc.mnuBackupAgoraClick(Sender: TObject);
begin
  BackupManual;
end;

//Menu Backup: abre a pasta dos backups da empresa
procedure TfrmPrinc.mnuBackupPastaClick(Sender: TObject);
begin
  AbrePastaBackup;
end;

//Botão abrir cadastro da empresa
procedure TfrmPrinc.mnuDadosEmprClick(Sender: TObject);
begin
  if DBPath.Text = 'Caminho da base de dados - FDB' then
  begin
    ShowMessage('Selecione a base de dados primeiro!');
  end
  else begin
    //Criar tela de loading
    frmEmpr.Show;
  end;
end;

{$R *.dfm}

initialization
SetLength(gridTemp,1);
SetLength(gridTemp[0],1);
SetLength(colUpdate,1);
colUpdateCount := 0;
end.
