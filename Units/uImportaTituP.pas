unit uImportaTituP;

interface

uses
  System.SysUtils, System.Classes, System.StrUtils, System.UITypes, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaTituP = class
    private
      colTituP, dadosTituP: string;
      colBTitu, dadosBTitu: string;
      cadastroForn: string;
      codi, empr, forn: string;
      quitado: Boolean;
      function Celula(coluna: string): string;
      function Inteiro(coluna, padrao: string): string;
      function Numero(coluna: string): string;
      function DataSQL(coluna: string; futuro: Boolean): string;
      function CodigoClieForn(valor: string): string;
      function CodigoTitulo(valor: string): string;
      procedure Adiciona(coluna, valor: string; baixa: Boolean);

    public
      k: Integer;
      StringGrid1: TStringGrid;
      constructor ImportaRegistro(numReg: Integer; Grid: TStringGrid);
      class procedure Inicia;
      procedure CarregaColunas;
      procedure Gravar;
  end;

implementation

uses
  importa_excel, importando;

const
  TABELA = 'titup';
  TABELA_BAIXA = 'btitup';
  COLUNA_CF = 'forn';
  NOME_CF = 'fornecedor';
  OPER_PADRAO = '101';

var
  //Títulos gravados nesta importação (código|fornecedor|empresa), para não repetir o código no TXT/SQL
  titulosGravados: TStringList;
  //Usuário escolheu "Cadastrar todos" para os fornecedores informados pelo nome
  cadastrarTodos: Boolean;

constructor cImportaTituP.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

//Chave (código|fornecedor|empresa) de um INSERT de título já gravado no arquivo TXT/SQL
//O INSERT começa sempre com (codi,fornecedor,empr,...)
function ChaveInsert(linha: string): string;
var
  p, n, nivel: Integer;
  aspas: Boolean;
  partes: array[0..2] of string;
begin
  Result := '';
  p := Pos('VALUES (', UpperCase(linha));
  if p = 0 then Exit;
  linha := Copy(linha, p + 8, MaxInt);
  n := 0;
  nivel := 0;
  aspas := False;
  for p := 1 to Length(linha) do
  begin
    if linha[p] = '''' then aspas := not aspas
    else if not aspas then
    begin
      if linha[p] = '(' then Inc(nivel)
      else if linha[p] = ')' then Dec(nivel)
      else if (linha[p] = ',') and (nivel = 0) then
      begin
        Inc(n);
        if n > 2 then Break;
        Continue;
      end;
    end;
    partes[n] := partes[n] + linha[p];
  end;
  if n <= 2 then Exit;
  //Código sem as aspas do SQL
  partes[0] := stringreplace(Copy(partes[0], 2, Length(partes[0])-2), '''''', '''',[rfReplaceAll]);
  Result := partes[0] + '|' + partes[1] + '|' + partes[2];
end;

//Chamado no início de cada importação
class procedure cImportaTituP.Inicia;
var
  arquivo: TStringList;
  i: Integer;
  chave: string;
begin
  titulosGravados.Clear;
  cadastrarTodos := False;

  //TXT/SQL: títulos que já estão no arquivo (importação continuada a partir de uma linha)
  if ((UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
      (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')) and
     FileExists(frmPrinc.DBPath.Text) then
  begin
    arquivo := TStringList.Create;
    try
      arquivo.LoadFromFile(frmPrinc.DBPath.Text);
      for i := 0 to arquivo.Count-1 do
        if StartsText('INSERT INTO ' + TABELA + ' (', arquivo[i]) then
        begin
          chave := ChaveInsert(arquivo[i]);
          if chave <> '' then titulosGravados.Add(chave);
        end;
    finally
      arquivo.Free;
    end;
  end;
end;

//Valor da célula da coluna na linha atual ('' se a coluna não existir)
function cImportaTituP.Celula(coluna: string): string;
var
  i: Integer;
begin
  Result := '';
  i := BuscaColuna(StringGrid1, coluna);
  if i <> -1 then Result := Trim(StringGrid1.Cells[i,k]);
end;

//Código inteiro da coluna; vazio ou sem a coluna assume o padrão
function cImportaTituP.Inteiro(coluna, padrao: string): string;
begin
  Result := Celula(coluna);
  if Result = '' then Result := padrao
  else if not IsNumeric(Result) then
    RegistraErro(coluna, Result, 'código numérico inválido')
  else
    //Sem zeros à esquerda (01 = 1)
    Result := IntToStr(StrToInt(Result));
end;

//Valor numérico da coluna para o SQL (aceita R$ e 2.500,50 / 2500.50); vazio é erro
function cImportaTituP.Numero(coluna: string): string;
var
  valor: string;
begin
  valor := Celula(coluna);
  Result := NumeroSQL(stringreplace(valor, 'R$', '',[rfReplaceAll, rfIgnoreCase]));
  if valor = '' then
    RegistraErro(coluna, valor, 'valor não informado')
  else if Result = '' then
    RegistraErro(coluna, valor, 'valor numérico inválido');
end;

//Data da coluna para o SQL ('dd.mm.aaaa'); vazio retorna ''
//Aceita dd/mm/aaaa, d/m/aa, dd-mm-aaaa, dd.mm.aaaa, aaaa-mm-dd, ddmmaaaa e ddmmaa
//Ano com 2 dígitos: futuro = True sempre 20aa (vencimento), senão segue AnoQuatroDigitos
function cImportaTituP.DataSQL(coluna: string; futuro: Boolean): string;
var
  original, temp, dia, mes, ano: string;
  partes: TStringList;
  d: TDateTime;
  c: Char;
begin
  Result := '';
  original := Celula(coluna);
  if original = '' then Exit;

  //Desconsidera a hora (01/02/2026 00:00:00)
  temp := original;
  if Pos(' ', temp) > 0 then temp := Copy(temp, 1, Pos(' ', temp)-1);
  temp := stringreplace(temp, '-', '/',[rfReplaceAll]);
  temp := stringreplace(temp, '.', '/',[rfReplaceAll]);

  if Pos('/', temp) > 0 then
  begin
    partes := TStringList.Create;
    try
      partes.Delimiter := '/';
      partes.StrictDelimiter := True;
      partes.DelimitedText := temp;
      if partes.Count <> 3 then temp := ''
      else if Length(partes[0]) = 4 then
        temp := Format('%2s%2s%s', [partes[2], partes[1], partes[0]])
      else
        temp := Format('%2s%2s%s', [partes[0], partes[1], partes[2]]);
      temp := stringreplace(temp, ' ', '0',[rfReplaceAll]);
    finally
      partes.Free;
    end;
  end;

  for c in temp do
    if not CharInSet(c, ['0'..'9']) then temp := '';

  dia := Copy(temp,1,2);
  mes := Copy(temp,3,2);
  ano := Copy(temp,5,4);
  if Length(temp) = 6 then
  begin
    if futuro then ano := '20' + ano
    else ano := AnoQuatroDigitos(ano);
  end
  else if Length(temp) <> 8 then ano := '';

  if (ano = '') or (StrToInt(ano) < 1900) or (StrToInt(ano) > 2099) or not TryEncodeDate(StrToInt(ano), StrToInt(mes), StrToInt(dia), d) then
    RegistraErro(coluna, original, 'data inválida')
  else
    Result := '''' + dia + '.' + mes + '.' + ano + '''';
end;

//Código do fornecedor para o SQL: pelo código ou pelo nome
//Nome não encontrado: pergunta se deseja cadastrar ou corrigir a planilha
function cImportaTituP.CodigoClieForn(valor: string): string;
var
  nome: string;
  bancoFDB: Boolean;
  resposta: Integer;
begin
  Result := '';
  bancoFDB := UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB';
  if valor = '' then
  begin
    RegistraErro(COLUNA_CF, valor, NOME_CF + ' não informado');
    Exit;
  end;

  //Números: código do fornecedor
  if IsNumeric(valor) then
  begin
    valor := IntToStr(StrToInt(valor));
    if bancoFDB and (querySelect('SELECT codi FROM clieforn WHERE codi = ' + valor) = '') then
      RegistraErro(COLUNA_CF, valor, 'código de ' + NOME_CF + ' não encontrado')
    else
      Result := valor;
    Exit;
  end;

  //Número com decimais (1.0, 1,5) não é código nem nome
  if NumeroSQL(valor) <> '' then
  begin
    RegistraErro(COLUNA_CF, valor, 'código numérico inválido');
    Exit;
  end;

  //Texto: nome do fornecedor (maiúsculas com acento, como no cadastro)
  nome := TextoSQL(Corta(AnsiUpperCase(valor), 60, COLUNA_CF), MaxInt);
  if errosLinha.Count > 0 then Exit;
  if bancoFDB then
  begin
    Result := querySelect('SELECT FIRST 1 codi FROM clieforn WHERE nome = ''' + nome + '''');
    if Result <> '' then Exit;
  end;

  if not cadastrarTodos then
  begin
    if bancoFDB then
      resposta := Mensagem('Linha ' + IntToStr(k) + ': ' + NOME_CF + ' "' + AnsiUpperCase(valor) + '" não encontrado.' + sLineBreak + sLineBreak +
        'Cadastrar: cadastra o ' + NOME_CF + ' com este nome.' + sLineBreak +
        'Corrigir: para a importação nesta linha para informar um código ou nome já cadastrado.',
        mtConfirmation, [mbYes, mbNo, mbYesToAll], ['Cadastrar', 'Corrigir', 'Cadastrar todos'], NOME_CF + ' não encontrado')
    else
      resposta := Mensagem('Linha ' + IntToStr(k) + ': ' + NOME_CF + ' "' + AnsiUpperCase(valor) + '" informado pelo nome.' + sLineBreak +
        'No arquivo TXT/SQL não é possível conferir se ele já existe.' + sLineBreak + sLineBreak +
        'Cadastrar: o script cadastra o ' + NOME_CF + ' caso não exista.' + sLineBreak +
        'Corrigir: para a importação nesta linha para informar o código.',
        mtConfirmation, [mbYes, mbNo, mbYesToAll], ['Cadastrar', 'Corrigir', 'Cadastrar todos'], NOME_CF + ' pelo nome');
    if resposta = mrYesToAll then cadastrarTodos := True
    else if resposta <> mrYes then
    begin
      Result := '';
      RegistraErro(COLUNA_CF, valor, NOME_CF + ' não cadastrado (informe um código ou nome já cadastrado)');
      Exit;
    end;
  end;

  if bancoFDB then
  begin
    Result := IntToStr(frmPrinc.cadastraClieForn('nome', '''' + nome + ''''));
    if Result = '-1' then
    begin
      Result := '';
      RegistraErro(COLUNA_CF, valor, 'erro ao cadastrar o ' + NOME_CF);
    end;
  end
  else begin
    //TXT/SQL: cadastra somente se não existir e busca o código pelo nome na execução do script
    cadastroForn := 'INSERT INTO clieforn (nome) SELECT ''' + nome + ''' FROM rdb$database' +
      ' WHERE NOT EXISTS (SELECT 1 FROM clieforn WHERE nome = ''' + nome + ''');';
    Result := '(SELECT FIRST 1 codi FROM clieforn WHERE nome = ''' + nome + ''')';
  end;
end;

//Código do título: se já existir para o mesmo fornecedor e empresa, acrescenta /1, /2...
function cImportaTituP.CodigoTitulo(valor: string): string;
var
  n: Integer;

  function Existe(codigo: string): Boolean;
  begin
    Result := titulosGravados.IndexOf(codigo + '|' + forn + '|' + empr) <> -1;
    if (not Result) and (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB') then
      Result := querySelect('SELECT codi FROM ' + TABELA + ' WHERE codi = ''' + TextoSQL(codigo, MaxInt) +
        ''' AND ' + COLUNA_CF + ' = ' + forn + ' AND empr = ' + empr) <> '';
  end;

begin
  Result := valor;
  if Length(valor) > 15 then
  begin
    RegistraErro('codi', valor, 'texto maior que 15 caracteres');
    Exit;
  end;
  n := 0;
  while Existe(Result) do
  begin
    Inc(n);
    Result := valor + '/' + IntToStr(n);
  end;
  if Length(Result) > 15 then
    RegistraErro('codi', valor, 'código já existe para este ' + NOME_CF + ' e não há espaço para acrescentar /' + IntToStr(n));
end;

//Adiciona a coluna no INSERT do título e, se baixa = True, no INSERT da baixa
procedure cImportaTituP.Adiciona(coluna, valor: string; baixa: Boolean);
begin
  colTituP := colTituP + ',' + coluna;
  dadosTituP := dadosTituP + ',' + valor;
  if baixa then
  begin
    colBTitu := colBTitu + ',' + coluna;
    dadosBTitu := dadosBTitu + ',' + valor;
  end;
end;

procedure cImportaTituP.CarregaColunas;
var
  valo, sald, venc, dataBaixa, hist, oper: string;
  formato: TFormatSettings;
begin
  frmImportando.atualizaStatus('Títulos a Pagar '+IntToStr(k));

  colTituP := '';
  dadosTituP := '';
  colBTitu := '';
  dadosBTitu := '';
  cadastroForn := '';
  quitado := False;
  errosLinha.Clear;

  //Carregar informações para importar (valores inválidos param a importação nesta linha)
  //-------------------------------------------------------

  //EMPR (EMPRESA)
  empr := Inteiro('empr', '1');
  if (errosLinha.Count = 0) and (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB') and
     (querySelect('SELECT codi FROM dados_empre WHERE codi = ' + empr) = '') then
    RegistraErro('empr', empr, 'empresa não encontrada');

  //VALO (VALOR) e SALD (SALDO): sem uma das colunas, assume o valor da outra
  formato := TFormatSettings.Create;
  formato.DecimalSeparator := '.';
  valo := '';
  sald := '';
  if BuscaColuna(StringGrid1,'valo') <> -1 then valo := Numero('valo');
  if BuscaColuna(StringGrid1,'sald') <> -1 then sald := Numero('sald');
  if BuscaColuna(StringGrid1,'valo') = -1 then valo := sald;
  if BuscaColuna(StringGrid1,'sald') = -1 then sald := valo;
  if (valo <> '') and (sald <> '') then
  begin
    if StrToFloat(valo, formato) <= 0 then
      RegistraErro(IfThen(BuscaColuna(StringGrid1,'valo') <> -1, 'valo', 'sald'), valo, 'valor deve ser maior que zero')
    else if StrToFloat(sald, formato) < 0 then
      RegistraErro('sald', Celula('sald'), 'saldo negativo')
    else if StrToFloat(sald, formato) > StrToFloat(valo, formato) then
      RegistraErro('sald', Celula('sald'), 'saldo maior que o valor do título');
    //Saldo zero: título quitado, grava a baixa
    quitado := StrToFloat(sald, formato) = 0;
  end;

  //Datas: VENC obrigatório e pode ser futuro (ano com 2 dígitos = 20aa)
  venc := DataSQL('venc', True);
  if (venc = '') and (Celula('venc') = '') then
    RegistraErro('venc', '', 'vencimento não informado');
  dataBaixa := DataSQL('data_baixa', False);
  if dataBaixa = '' then dataBaixa := venc;

  //HIST (HISTORICO): maiúsculas com acento, como no ERP
  hist := Corta(AnsiUpperCase(Celula('hist')), 200, 'hist');

  //Campos do título
  Adiciona('empr', empr, True);
  Adiciona('loca_cobr', Inteiro('loca_cobr', '1'), False);
  Adiciona('cart', Inteiro('cart', '1'), False);
  oper := Corta(Celula('oper'), 10, 'oper');
  if oper = '' then oper := OPER_PADRAO;
  Adiciona('oper', '''' + TextoSQL(oper, MaxInt) + '''', True);
  Adiciona('c_func', Inteiro('c_func', '1'), True);
  if Celula('data') <> '' then Adiciona('data', DataSQL('data', False), False);
  Adiciona('venc', venc, False);
  Adiciona('valo', valo, True);
  Adiciona('sald', sald, False);
  if hist <> '' then
  begin
    colTituP := colTituP + ',hist';
    dadosTituP := dadosTituP + ',''' + TextoSQL(hist, MaxInt) + '''';
    //Histórico da baixa tem 150 caracteres
    colBTitu := colBTitu + ',hist';
    dadosBTitu := dadosBTitu + ',''' + TextoSQL(hist, 150) + '''';
  end;

  //Campos da baixa (somente título quitado)
  colBTitu := colBTitu + ',data,tota,id,juro,mult,desco,cont,empr_baix';
  dadosBTitu := dadosBTitu + ',' + dataBaixa + ',' + valo + ',gen_id(gen_' + TABELA_BAIXA + '_id,1),0,0,0,1,' + empr;

  //Fornecedor e código somente se o restante da linha estiver correto
  //(evita perguntar sobre o cadastro do fornecedor numa linha que vai parar)
  if Celula(COLUNA_CF) = '' then RegistraErro(COLUNA_CF, '', NOME_CF + ' não informado');
  if errosLinha.Count > 0 then Exit;

  //FORN (FORNECEDOR): código ou nome
  forn := CodigoClieForn(Celula(COLUNA_CF));
  if errosLinha.Count > 0 then Exit;

  //CODI (CÓDIGO): sem a coluna, usa o número da linha da planilha
  if BuscaColuna(StringGrid1,'codi') <> -1 then
    codi := stringreplace(Celula('codi'), '.', '',[rfReplaceAll])
  else
    codi := IntToStr(k);
  if codi = '' then
  begin
    RegistraErro('codi', '', 'código do título não informado');
    Exit;
  end;
  codi := CodigoTitulo(codi);

  colTituP := 'codi,' + COLUNA_CF + colTituP;
  dadosTituP := '''' + TextoSQL(codi, MaxInt) + ''',' + forn + dadosTituP;
  colBTitu := 'codi,' + COLUNA_CF + colBTitu;
  dadosBTitu := '''' + TextoSQL(codi, MaxInt) + ''',' + forn + dadosBTitu;
end;

procedure cImportaTituP.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Valores inválidos: para nesta linha para o usuário corrigir e continuar daqui
  if errosLinha.Count > 0 then
  begin
    ShowMessage('Linha ' + IntToStr(k) + ' com valores inválidos:' + sLineBreak + sLineBreak +
      errosLinha.Text + sLineBreak + 'Corrija na planilha e clique em Importar para continuar a partir desta linha.');
    status := 0;
    Exit;
  end;

  //----------------------------------------
  //Gravar no banco de Títulos a Pagar
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    SQL := TSQLDataSet.Create(Nil);
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela TITUP.');
        SQL.CommandText := 'INSERT INTO ' + TABELA + ' ('+ colTituP +') VALUES ' + '(' + dadosTituP + ');';
        SQL.ExecSQL;

        if quitado then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Inserindo dados na tabela BTITUP.');
          SQL.CommandText := 'INSERT INTO ' + TABELA_BAIXA + ' ('+ colBTitu +') VALUES ' + '(' + dadosBTitu +');';
          SQL.ExecSQL;
        end;

      except
        on e: exception do
        begin
          ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
          status := 0;
          Exit;
        end;
      end;

    finally
      SQL.Free;
      frmPrinc.conDestino.Close;
    end;
  end
  //Gravar comandos em TXT
  else if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
          (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then begin
    //Carregar arquivo TXT
    AssignFile(fileTXT, frmPrinc.DBPath.Text);
    if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
    else append(fileTXT);
    try
      try
        //Fornecedor informado pelo nome: cadastra se não existir
        if cadastroForn <> '' then begin
          WriteLn(fileTXT, '----------Comandos da CLIEFORN----------');
          WriteLn(fileTXT, 'ALTER TRIGGER clieforn_biu0 INACTIVE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
          WriteLn(fileTXT, cadastroForn);
          WriteLn(fileTXT, 'COMMIT WORK;');
          WriteLn(fileTXT, 'ALTER TRIGGER clieforn_biu0 ACTIVE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end;

        frmImportando.atualizaStatus('Comandos da TITUP.');
        WriteLn(fileTXT, '----------Comandos da TITUP----------');
        WriteLn(fileTXT, 'INSERT INTO ' + TABELA + ' ('+ colTituP +') VALUES ' + '(' + dadosTituP + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');

        if quitado then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Comandos BTITUP.');
          WriteLn(fileTXT, '----------Comandos da BTITUP----------');
          WriteLn(fileTXT, 'INSERT INTO ' + TABELA_BAIXA + ' ('+ colBTitu +') VALUES ' + '(' + dadosBTitu +');');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end;
      except
        on e: exception do
        begin
          ShowMessage('Erro TXT: '+e.message);
          status := 0;
          Exit;
        end;
      end;
    finally
      CloseFile(fileTXT);
    end;
  end;

  titulosGravados.Add(codi + '|' + forn + '|' + empr);
end;

initialization
  titulosGravados := TStringList.Create;

finalization
  titulosGravados.Free;

end.
