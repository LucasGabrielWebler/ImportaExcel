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
    mnuLimpaGrupos: TMenuItem;
    mnuLimpaSubGrupos: TMenuItem;
    mnuLimpaMarcas: TMenuItem;
    mnuLimpaProdutos: TMenuItem;
    mnuLimpaTituP: TMenuItem;
    mnuLimpaTituR: TMenuItem;
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
    procedure mnuLimpaGruposClick(Sender: TObject);
    procedure mnuLimpaSubGruposClick(Sender: TObject);
    procedure mnuLimpaMarcasClick(Sender: TObject);
    procedure mnuLimpaProdutosClick(Sender: TObject);
    procedure mnuLimpaTituPClick(Sender: TObject);
    procedure mnuLimpaTituRClick(Sender: TObject);
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

  private
    { Private declarations }
  public
    { Public declarations }
    function StringGridToArray(Grid: TStringGrid): Integer;
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
  uImportaClieForn, uImportaProduto, uImportaGrupo, uImportaSubGrupo,
  uImportaMarca, uImportaTituP, uImportaTituR, uImportaCUSTOM;

//Função para definir a quantidade de empresas
function TfrmPrinc.quantidadeEmpresas(colEmpr: Integer): Integer;
var
  i, max: Integer;
begin
  max := 1;
  for i := StrToInt(StartLine.Text) to StringGrid1.RowCount-1 do begin
    if StrToInt(StringGrid1.Cells[colEmpr,i]) > max  then begin
      max := StrToInt(StringGrid1.Cells[colEmpr,i]);
    end;
  end;
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
  gen_id: Integer;
  queryTemp: TSQLQuery;

begin
  try
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;

      //Desativar Trigger das cidades
      queryTemp.CommandText := 'ALTER TRIGGER clieforn_biu0 INACTIVE;';
      queryTemp.ExecSQL;
      //Executar INSERT
      queryTemp.CommandText := 'INSERT INTO clieforn ('+ colClieForn +') VALUES ' + '(' + dadosClieForn + ');';
      queryTemp.ExecSQL;
      //Reativar Trigger das cidades
      queryTemp.CommandText := 'ALTER TRIGGER clieforn_biu0 ACTIVE;';
      queryTemp.ExecSQL;

      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('SELECT c.codi FROM clieforn c WHERE c.codi = gen_id(gen_clieforn_id,0);');
      queryTemp.Open;

    except
      on e: exception do
      begin
        ShowMessage('Erro SQL: '+e.message+sLineBreak+queryTemp.CommandText);
      end;
    end;
  finally
    if queryTemp.IsEmpty then
    begin
      Result := -1;
    end
    else begin
      Result := queryTemp.FieldByName('CODI').AsInteger;
    end;
    queryTemp.Close;
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
  then Result := '(SELECT CID_CODI FROM CIDADE WHERE CID_DESC = '''+Cidade+''' AND CID_UF = '''+UF+''')'
  else
  begin
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
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
        frmPrinc.conDestino.Open;
        queryTemp := TSQLQuery.Create(nil);
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


  //Status se esta OK ou se tem erro, setado como OK
  status := 1;

  //Flag se o botão continuar foi clicado
  butContinue := 0;

  try
    try
      //Criar tela de loading
      frmImportando.Show;
      frmImportando.Label2.Font.Color := clBlack;

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
        conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
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
            //Verificar se existe coluna margem, recalcular preços
            i:=BuscaColuna(StringGrid1,'margem');
            if (i<>-1) then
            begin
              frmImportando.atualizaStatus('Ajustando Preços.');
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_preco_prazo = pc.cust_custo_real+(pc.cust_custo_real * pc.cust_margem1 /100) WHERE pc.cust_custo_real>0;';
              SQL.ExecSQL;
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_preco_vista = pc.cust_preco_prazo;';
              SQL.ExecSQL;
            end
            //Se não existir coluna margem, recalcular margem
            else begin
              //Atualizar MARGEM1
              frmImportando.atualizaStatus('Ajustando MARGENS.');
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem1= ABS(pc.cust_preco_prazo - pc.cust_custo_real)/pc.cust_custo_real WHERE pc.cust_custo_real>0;';
              SQL.ExecSQL;
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem1 = pc.cust_margem1 * 100;';
              SQL.ExecSQL;
            end;

            //Atualizar MARGEM2
            SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem2 = (CAST(pc.cust_preco_vista as numeric (18,2))/CAST(pc.cust_preco_prazo as numeric (18,2)) -1)*100 WHERE CAST(pc.cust_preco_prazo as numeric (18,2))>0;';
            SQL.ExecSQL;
            //Criar registro na PROD_AJUS
            frmImportando.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
            SQL.CommandText := 'INSERT INTO prod_ajus (codi,data) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE);';
            SQL.ExecSQL;
          end
          //Se for UPDATE
          else begin
            if (colUpdateCount > 0) and (BuscaColuna(StringGrid1,'qtd') <> -1) then begin
              //Criar registro na PROD_AJUS somente se for feito update na QTD da Itens
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
              SQL.CommandText := 'INSERT INTO prod_ajus (codi,data) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE);';
              SQL.ExecSQL;
            end;

            //Verificar se existe coluna margem, recalcular preços
            i:=BuscaColuna(StringGrid1,'margem');
            if (i<>-1) then
            begin
              frmImportando.atualizaStatus('Ajustando Preços.');
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_preco_prazo = pc.cust_custo_real+(pc.cust_custo_real * pc.cust_margem1 /100) WHERE pc.cust_custo_real>0;';
              SQL.ExecSQL;
              SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_preco_vista = pc.cust_preco_prazo;';
              SQL.ExecSQL;;
            end
            //Se não existir coluna margem, ve se tem alguma coluna de preço e recalcular margem
            else begin
              //Preço A PRAZO
              i:=BuscaColuna(StringGrid1,'preco_prazo');
              if (i<>-1) then
              begin
                //Atualizar MARGEM1
                frmImportando.atualizaStatus('Ajustando MARGEM 1.');
                SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem1= ABS(pc.cust_preco_prazo - pc.cust_custo_real)/pc.cust_custo_real WHERE pc.cust_custo_real>0;';
                SQL.ExecSQL;
                SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem1 = pc.cust_margem1 * 100;';
                SQL.ExecSQL;
              end;
              //Preço A VISTA
              i:=BuscaColuna(StringGrid1,'preco_vista');
              if (i<>-1) then
              begin
                //Atualizar MARGEM2
                frmImportando.atualizaStatus('Ajustando MARGEM 2.');
                SQL.CommandText := 'UPDATE prod_custos pc SET pc.cust_margem2= (pc.cust_preco_prazo - pc.cust_preco_vista)*100/pc.cust_preco_prazo WHERE pc.cust_preco_prazo>0;';
                SQL.ExecSQL;
              end;
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

            //Verificar se existe coluna margem, recalcular preços
            i:=BuscaColuna(StringGrid1,'margem');
            if (i<>-1) then
            begin
              frmImportando.atualizaStatus('Ajustando Preços.');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_preco_prazo = pc.cust_custo_real+(pc.cust_custo_real * pc.cust_margem1 /100) WHERE pc.cust_custo_real>0;');
              WriteLn(fileTXT, 'COMMIT WORK;');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_preco_vista = pc.cust_preco_prazo;');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end
            //Se não existir coluna margem, recalcular margem
            else begin
              //Atualizar MARGEM1
              frmImportando.atualizaStatus('Ajustando MARGENS.');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem1= ABS(pc.cust_preco_prazo - pc.cust_custo_real)/pc.cust_custo_real WHERE pc.cust_custo_real>0;');
              WriteLn(fileTXT, 'COMMIT WORK;');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem1 = pc.cust_margem1 * 100;');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end;

            //Atualizar MARGEM2
            WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem2 = (CAST(pc.cust_preco_vista as numeric (18,2))/CAST(pc.cust_preco_prazo as numeric (18,2)) -1)*100 WHERE CAST(pc.cust_preco_prazo as numeric (18,2))>0;');
            WriteLn(fileTXT, 'COMMIT WORK;');
            //Criar registro na PROD_AJUS
            frmImportando.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
            WriteLn(fileTXT, 'INSERT INTO prod_ajus (codi,data) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE);');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end

          //Se for UPDATE
          else begin
            if (colUpdateCount > 0) and (BuscaColuna(StringGrid1,'qtd') <> -1) then begin
              //Criar registro na PROD_AJUS somente se for feito update na QTD da Itens
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
              WriteLn(fileTXT, 'INSERT INTO prod_ajus (codi,data) VALUES (gen_id(gen_prod_ajus_id,1),CURRENT_DATE);');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end;

            //Verificar se existe coluna margem, recalcular preços
            i:=BuscaColuna(StringGrid1,'margem');
            if (i<>-1) then
            begin
              frmImportando.atualizaStatus('Ajustando Preços.');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_preco_prazo = pc.cust_custo_real+(pc.cust_custo_real * pc.cust_margem1 /100) WHERE pc.cust_custo_real>0;');
              WriteLn(fileTXT, 'COMMIT WORK;');
              WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_preco_vista = pc.cust_preco_prazo;');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end
            //Se não existir coluna margem, ve se tem alguma coluna de preço e recalcular margem
            else begin
              //Preço A PRAZO
              i:=BuscaColuna(StringGrid1,'preco_prazo');
              if (i<>-1) then
              begin
                //Atualizar MARGEM1
                frmImportando.atualizaStatus('Ajustando MARGEM 1.');
                WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem1= ABS(pc.cust_preco_prazo - pc.cust_custo_real)/pc.cust_custo_real WHERE pc.cust_custo_real>0;');
                WriteLn(fileTXT, 'COMMIT WORK;');
                WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem1 = pc.cust_margem1 * 100;');
                WriteLn(fileTXT, 'COMMIT WORK;');
              end;
              //Preço A VISTA
              i:=BuscaColuna(StringGrid1,'preco_vista');
              if (i<>-1) then
              begin
                //Atualizar MARGEM2
                frmImportando.atualizaStatus('Ajustando MARGEM 2.');
                WriteLn(fileTXT, 'UPDATE prod_custos pc SET pc.cust_margem2= (pc.cust_preco_prazo - pc.cust_preco_vista)*100/pc.cust_preco_prazo WHERE pc.cust_preco_prazo>0;');
                WriteLn(fileTXT, 'COMMIT WORK;');
              end;
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

        else if SelectImport.Text='Marcas' then
        begin
          //Arrumar Generator das MARCAS
          frmImportando.atualizaStatus('Alterar Generator das Marcas.');
          WriteLn(fileTXT, 'SELECT gen_id(GEN_MARCA_ID, ABS((SELECT MAX(CODI) FROM marca) - (SELECT gen_id(GEN_MARCA_ID,0) FROM RDB$DATABASE)) ) FROM RDB$DATABASE;');
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


//Limpar dados de clientes e fornecedores do banco
procedure TfrmPrinc.mnuLimpaClieFornClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM clieforn;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_CLIEFORN_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de clientes e fornecedores.');

  //Fechar conexoes
  SQL.Free;
  conDestino.Close;
end;


//Limpar dados de grupos do banco
procedure TfrmPrinc.mnuLimpaGruposClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
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
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
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
procedure TfrmPrinc.mnuLimpaProdutosClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
  SQL.SQLConnection := conDestino;

  SQL.CommandText := 'DELETE FROM prod;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM prod_ajus;';
  SQL.ExecSQL;
  SQL.CommandText := 'DELETE FROM itens;';
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
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
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
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
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
  //Abrir conexoes
  conDestino.Open;
  SQL := TSQLDataSet.Create(Nil);
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
    Canvas.Font.Color := clBlack;
    Canvas.Brush.Color := clWhite;
    // Don't change color for first Column, first row
    if (ARow <> 0) and (ACol <> 0) then begin
      // Draw the Band
      if ARow mod 2 = 0 then
        Canvas.Brush.Color := $00E1FFF9
      else
        Canvas.Brush.Color := $00FFEBDF;
    end;
    for i := 0 to colUpdateCount-1 do begin
      if StringGrid1.Cells[Acol,Arow] = colUpdate[i] then
        Canvas.Font.Color := clBlue;
    end;
    Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, cells[acol, arow]);
    Canvas.FrameRect(Rect);
  end;
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
