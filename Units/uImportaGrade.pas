unit uImportaGrade;

interface

uses
  System.SysUtils, System.Classes, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr, Data.DBXCommon,
  uUtil;

const
  //Observação dos ajustes de estoque (PROD_AJUS) criados pela importação de grades
  OBS_AJUSTE = 'IMPORTADO PELO IMPORTA EXCEL (GRADES)';

type
  cImportaGrade = class
    private
      //Ajuste da importação (GRAD_AJUS + PROD_AJUS), criado na primeira linha que movimenta estoque
      class var temAjuste: Boolean;
      class var ajusteGrade, ajusteProd: string;

      prodCod, emprCod, gradeCod, gradeDesc, titLinha, titColuna, descLinha, descColuna: string;
      qtdFiscal, qtdFisico, condMovimento: string;
      function GradeValor: string;
      function CondGrade: string;
      function SemEstoque: string;
      function ProdPrincipal: string;
      function Celula: string;
      function Dimensao(grade: string): string;
      function Numero(coluna: string): string;
      procedure MontaComandos(estrutura, estoque: TStringList);
      function Substitui(cmd, lancamento: string): string;

    public
      k: Integer;
      StringGrid1: TStringGrid;
      class procedure Inicia;
      class function SQLMatriz: string;
      constructor ImportaRegistro(numReg: Integer; Grid: TStringGrid);
      procedure CarregaColunas;
      procedure Gravar;
  end;

implementation

uses
  importa_excel, importando;

const
  //Coluna neutra das grades de uma dimensão (só linha, ex.: tamanho)
  COLUNA_UNICA = 'UNICO';

  //Generators das tabelas usadas: avança o generator que estiver atrás do maior código (nunca volta)
  GENERATORS: array[0..8] of array[0..2] of string = (
    ('gen_grade_id', 'grade', 'grad_id'),
    ('gen_grade_linha_id', 'grade_linha', 'grli_id'),
    ('gen_grade_coluna_id', 'grade_coluna', 'grco_id'),
    ('gen_grade_estoque_id', 'grade_estoque', 'gres_id'),
    ('gen_grade_lancamento_id', 'grade_lancamento', 'grla_id'),
    ('gen_grade_lancamento_itens_id', 'grade_lancamento_itens', 'grit_id'),
    ('gen_grade_ajus_id', 'grad_ajus', 'graj_id'),
    ('gen_prod_ajus_id', 'prod_ajus', 'codi'),
    ('gen_itens_id', 'itens', 'codi'));

  //Completa o GRADE_ESTOQUE (estoque zero) de todas as combinações linha x coluna x empresa
  //dos produtos com grade, como o ERP faz ao vincular a grade ao produto
  SQL_MATRIZ = 'INSERT INTO grade_estoque (gres_id, gres_qtd, gres_qtd_max, grad_id, prod_codi, prod_empr, grli_id, grco_id, empr) ' +
    'SELECT gen_id(gen_grade_estoque_id,1), 0, 0, p.grade_grad_id, p.codi, p.empr, l.grli_id, c.grco_id, e.codi ' +
    'FROM prod p ' +
    'JOIN grade_linha l ON l.grad_id = p.grade_grad_id ' +
    'JOIN grade_coluna c ON c.grad_id = p.grade_grad_id ' +
    'CROSS JOIN dados_empre e ' +
    'WHERE %s ' +
    'AND NOT EXISTS (SELECT 1 FROM grade_estoque ge WHERE ge.prod_codi = p.codi AND ge.prod_empr = p.empr ' +
    'AND ge.grli_id = l.grli_id AND ge.grco_id = c.grco_id AND ge.empr = e.codi);';

//Inicia a importação: o ajuste é criado só quando alguma linha movimentar estoque
class procedure cImportaGrade.Inicia;
var
  SQL: TSQLQuery;
  fileTXT: TextFile;
  i: Integer;
  cmd: string;
begin
  temAjuste := False;
  ajusteGrade := '';
  ajusteProd := '';

  //Generators atrás do maior código (base com cadastros feitos por script) causariam código repetido
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then
  begin
    SQL := TSQLQuery.Create(nil);
    try
      frmPrinc.conDestino.Open;
      SQL.SQLConnection := frmPrinc.conDestino;
      for i := 0 to High(GENERATORS) do
      begin
        SQL.SQL.Text := 'SELECT gen_id(' + GENERATORS[i][0] + ', MAXVALUE(0, (SELECT COALESCE(MAX(' + GENERATORS[i][2] + '),0) FROM ' +
                        GENERATORS[i][1] + ') - gen_id(' + GENERATORS[i][0] + ',0))) FROM RDB$DATABASE';
        SQL.Open;
        SQL.Close;
      end;
    finally
      SQL.Free;
      frmPrinc.conDestino.Close;
    end;
  end
  else if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
          (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL') then
  begin
    AssignFile(fileTXT, frmPrinc.DBPath.Text);
    if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
    else append(fileTXT);
    try
      WriteLn(fileTXT, '----------Generators das grades----------');
      for i := 0 to High(GENERATORS) do
      begin
        cmd := 'SELECT gen_id(' + GENERATORS[i][0] + ', MAXVALUE(0, (SELECT COALESCE(MAX(' + GENERATORS[i][2] + '),0) FROM ' +
               GENERATORS[i][1] + ') - gen_id(' + GENERATORS[i][0] + ',0))) FROM RDB$DATABASE;';
        WriteLn(fileTXT, cmd);
      end;
      WriteLn(fileTXT, 'COMMIT WORK;');
    finally
      CloseFile(fileTXT);
    end;
  end;
end;

//Pós-importação: completa o estoque da grade de todos os produtos com grade
//(ex.: tamanho novo cadastrado numa grade que outros produtos já usavam)
class function cImportaGrade.SQLMatriz: string;
begin
  Result := Format(SQL_MATRIZ, ['p.grade_grad_id IS NOT NULL']);
end;

constructor cImportaGrade.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

//Código da grade informada, por consulta (nulo se a grade não existir; vazio sem a coluna GRADE)
function cImportaGrade.GradeValor: string;
begin
  if gradeCod <> '' then
    Result := '(SELECT g.grad_id FROM grade g WHERE g.grad_id = ' + gradeCod + ')'
  else if gradeDesc <> '' then
    Result := '(SELECT FIRST 1 g.grad_id FROM grade g WHERE UPPER(g.grad_descricao) = ''' + TextoSQL(gradeDesc, 90) + ''' ORDER BY g.grad_id)'
  else
    Result := '';
end;

//Condição da grade do produto: a grade informada (código ou descrição) ou, sem a coluna GRADE, a que o produto já tem
function cImportaGrade.CondGrade: string;
begin
  if GradeValor <> '' then
    Result := 'p.grade_grad_id = ' + GradeValor
  else
    Result := 'p.grade_grad_id IS NOT NULL';
end;

//Produto sem estoque (fiscal e físico zerados em todas as empresas): só assim pode receber a grade
function cImportaGrade.SemEstoque: string;
begin
  Result := 'NOT EXISTS (SELECT 1 FROM prod_esto pe WHERE pe.cod_prod = ' + prodCod + ' AND (pe.qtd <> 0 OR pe.qtd_max <> 0))';
end;

//Cadastro do produto (empresa onde foi criado)
function cImportaGrade.ProdPrincipal: string;
begin
  Result := 'p.codi = ' + prodCod + ' AND p.empr = (SELECT MIN(p2.empr) FROM prod p2 WHERE p2.codi = ' + prodCod + ')';
end;

//Coluna coerente com a grade: sem COLUNA (UNICO) só em grade sem colunas; com COLUNA só em grade que não é de uma dimensão
function cImportaGrade.Dimensao(grade: string): string;
begin
  if descColuna = COLUNA_UNICA then
    Result := 'NOT EXISTS (SELECT 1 FROM grade_coluna c2 WHERE c2.grad_id = ' + grade + ' AND UPPER(c2.grco_descricao) <> ''' + COLUNA_UNICA + ''')'
  else
    Result := 'NOT EXISTS (SELECT 1 FROM grade_coluna c2 WHERE c2.grad_id = ' + grade + ' AND UPPER(c2.grco_descricao) = ''' + COLUNA_UNICA + ''')';
end;

//Combinação (linha x coluna) do produto na empresa do estoque
function cImportaGrade.Celula: string;
begin
  Result := 'FROM prod p ' +
    'JOIN grade_estoque ge ON ge.prod_codi = p.codi AND ge.prod_empr = p.empr AND ge.grad_id = p.grade_grad_id AND ge.empr = ' + emprCod + ' ' +
    'JOIN grade_linha l ON l.grli_id = ge.grli_id AND l.grad_id = p.grade_grad_id ' +
    'JOIN grade_coluna c ON c.grco_id = ge.grco_id AND c.grad_id = p.grade_grad_id ' +
    'WHERE ' + ProdPrincipal + ' AND ' + CondGrade + ' ' +
    'AND UPPER(l.grli_descricao) = ''' + TextoSQL(descLinha, 90) + ''' ' +
    'AND UPPER(c.grco_descricao) = ''' + TextoSQL(descColuna, 90) + '''';
end;

//Quantidade da planilha: vazio não altera o estoque; aceita 2.500,50 / 2500.50 e negativo (até 4 casas decimais)
function cImportaGrade.Numero(coluna: string): string;
var
  i: Integer;
  valor: Double;
  formato: TFormatSettings;
begin
  Result := '';
  i := BuscaColuna(StringGrid1, coluna);
  if (i = -1) or (Trim(StringGrid1.Cells[i,k]) = '') then Exit;
  formato := TFormatSettings.Create;
  formato.DecimalSeparator := '.';
  formato.ThousandSeparator := ',';
  if (NumeroSQL(LimpaNumero(StringGrid1.Cells[i,k])) = '') or
     not TryStrToFloat(NumeroSQL(LimpaNumero(StringGrid1.Cells[i,k])), valor, formato) then
    RegistraErro(coluna, StringGrid1.Cells[i,k], 'valor numérico inválido')
  //Estoque da grade tem 4 casas decimais e o da ITENS até 12 dígitos inteiros
  else if Abs(valor) >= 1E12 then
    RegistraErro(coluna, StringGrid1.Cells[i,k], 'quantidade muito grande')
  else
    Result := FormatFloat('0.####', valor, formato);
end;

procedure cImportaGrade.CarregaColunas;
var
  i, n: Integer;
  temp, gradeProd, gradeAlvo: string;
  bancoFDB: Boolean;
  partes: TStringList;
begin
  prodCod := '';
  emprCod := '1';
  gradeCod := '';
  gradeDesc := '';
  titLinha := '';
  titColuna := '';
  descLinha := '';
  descColuna := COLUNA_UNICA;
  qtdFiscal := '';
  qtdFisico := '';
  errosLinha.Clear;
  bancoFDB := UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB';

  frmImportando.atualizaStatus('Grade '+ IntToStr(k));

  //PROD: código do produto (obrigatório, já importado em Produtos)
  i := BuscaColuna(StringGrid1,'prod');
  temp := Trim(stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll]));
  if temp = '' then
    RegistraErro('prod', '', 'código do produto em branco')
  else if StrToIntDef(temp,-1) <= 0 then
    RegistraErro('prod', StringGrid1.Cells[i,k], 'código numérico inválido')
  else
    prodCod := IntToStr(StrToInt(temp));

  //EMPR: empresa onde o estoque vai aparecer (padrão 1)
  i := BuscaColuna(StringGrid1,'empr');
  if (i <> -1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
  begin
    temp := Trim(StringGrid1.Cells[i,k]);
    if StrToIntDef(temp,-1) <= 0 then
    begin
      RegistraErro('empr', StringGrid1.Cells[i,k], 'código numérico inválido');
      emprCod := '';
    end
    else
      emprCod := IntToStr(StrToInt(temp));
  end;

  //GRADE: código (somente números) de uma grade cadastrada ou descrição (cadastrada se não existir)
  i := BuscaColuna(StringGrid1,'grade');
  if (i <> -1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
  begin
    temp := Trim(StringGrid1.Cells[i,k]);
    if SoNumeros(temp) then
    begin
      if StrToIntDef(temp,-1) <= 0 then
        RegistraErro('grade', StringGrid1.Cells[i,k], 'código de grade inválido')
      else
        gradeCod := IntToStr(StrToInt(temp));
    end
    else
      gradeDesc := Corta(AnsiUpperCase(temp), 90, 'grade');
  end;

  //Títulos usados somente ao cadastrar a grade (ex.: TAMANHO e COR)
  i := BuscaColuna(StringGrid1,'tit_linha');
  if i <> -1 then titLinha := Corta(AnsiUpperCase(Trim(StringGrid1.Cells[i,k])), 90, 'tit_linha');
  i := BuscaColuna(StringGrid1,'tit_coluna');
  if i <> -1 then titColuna := Corta(AnsiUpperCase(Trim(StringGrid1.Cells[i,k])), 90, 'tit_coluna');

  //LINHA (obrigatória, ex.: tamanho) e COLUNA (ex.: cor; vazia = grade de uma dimensão, coluna UNICO)
  i := BuscaColuna(StringGrid1,'linha');
  temp := Trim(StringGrid1.Cells[i,k]);
  if temp = '' then
    RegistraErro('linha', '', 'linha da grade em branco')
  else
    descLinha := Corta(AnsiUpperCase(temp), 90, 'linha');
  i := BuscaColuna(StringGrid1,'coluna');
  if (i <> -1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
  begin
    descColuna := Corta(AnsiUpperCase(Trim(StringGrid1.Cells[i,k])), 90, 'coluna');
  end;

  //Estoque: QTD = fiscal e físico; EST = só fiscal; MAX = só físico (saldo final da combinação)
  temp := Numero('qtd');
  if temp <> '' then
  begin
    qtdFiscal := temp;
    qtdFisico := temp;
  end;
  temp := Numero('est');
  if temp <> '' then qtdFiscal := temp;
  temp := Numero('max');
  if temp <> '' then qtdFisico := temp;

  //Banco FDB: confere produto, empresa e grade antes de gravar
  if bancoFDB and (prodCod <> '') then
  begin
    partes := TStringList.Create;
    try
      partes.Delimiter := '|';
      partes.StrictDelimiter := True;
      if gradeCod <> '' then temp := '(SELECT grad_id FROM grade WHERE grad_id = ' + gradeCod + ')'
      else if gradeDesc <> '' then temp := '(SELECT FIRST 1 grad_id FROM grade WHERE UPPER(grad_descricao) = ''' + TextoSQL(gradeDesc, 90) + ''' ORDER BY grad_id)'
      else temp := 'NULL';
      partes.DelimitedText := querySelect('SELECT (SELECT COUNT(*) FROM prod WHERE codi = ' + prodCod + ') || ''|'' || ' +
        'COALESCE((SELECT FIRST 1 grade_grad_id FROM prod WHERE codi = ' + prodCod + ' ORDER BY empr), 0) || ''|'' || ' +
        '(SELECT COUNT(*) FROM prod_esto WHERE cod_prod = ' + prodCod + ' AND (qtd <> 0 OR qtd_max <> 0)) || ''|'' || ' +
        '(SELECT COUNT(*) FROM dados_empre WHERE codi = ' + IntToStr(StrToIntDef(emprCod, 0)) + ') || ''|'' || ' +
        'COALESCE(' + temp + ', 0) FROM RDB$DATABASE');
      if partes.Count = 5 then
      begin
        gradeProd := partes[1];
        if partes[0] = '0' then
          RegistraErro('prod', prodCod, 'produto não encontrado')
        else begin
          if (emprCod <> '') and (partes[3] = '0') then
            RegistraErro('empr', emprCod, 'empresa não encontrada');
          if (gradeCod <> '') and (partes[4] = '0') then
            RegistraErro('grade', gradeCod, 'grade não encontrada');
          if gradeProd = '0' then
          begin
            //Produto ainda sem grade: precisa da grade e não pode ter estoque (somaria com o da grade)
            if (gradeCod = '') and (gradeDesc = '') then
              RegistraErro('grade', '', 'produto sem grade vinculada: informe a GRADE');
            if partes[2] <> '0' then
              RegistraErro('prod', prodCod, 'produto possui estoque sem grade (zere o estoque do produto antes de vincular a grade)');
            gradeAlvo := partes[4];
          end
          else begin
            if ((gradeCod <> '') or (gradeDesc <> '')) and (partes[4] <> gradeProd) then
              RegistraErro('grade', StringGrid1.Cells[BuscaColuna(StringGrid1,'grade'),k], 'produto já vinculado a outra grade (código ' + gradeProd + ')');
            gradeAlvo := gradeProd;
          end;

          //Grade já cadastrada: confere a coluna (uma ou duas dimensões) e cadastros repetidos
          if (gradeAlvo <> '0') and (descLinha <> '') and (descColuna <> '') then
          begin
            partes.DelimitedText := querySelect('SELECT ' +
              '(SELECT COUNT(*) FROM grade_coluna WHERE grad_id = ' + gradeAlvo + ' AND UPPER(grco_descricao) = ''' + COLUNA_UNICA + ''') || ''|'' || ' +
              '(SELECT COUNT(*) FROM grade_coluna WHERE grad_id = ' + gradeAlvo + ' AND UPPER(grco_descricao) <> ''' + COLUNA_UNICA + ''') || ''|'' || ' +
              '(SELECT COUNT(*) FROM grade_linha WHERE grad_id = ' + gradeAlvo + ' AND UPPER(grli_descricao) = ''' + TextoSQL(descLinha, 90) + ''') || ''|'' || ' +
              '(SELECT COUNT(*) FROM grade_coluna WHERE grad_id = ' + gradeAlvo + ' AND UPPER(grco_descricao) = ''' + TextoSQL(descColuna, 90) + ''') FROM RDB$DATABASE');
            if partes.Count = 4 then
            begin
              if (descColuna = COLUNA_UNICA) and (partes[1] <> '0') then
                RegistraErro('coluna', '', 'coluna em branco: a grade possui colunas (ex.: cor)')
              else if (descColuna <> COLUNA_UNICA) and (partes[0] <> '0') then
                RegistraErro('coluna', descColuna, 'a grade não possui colunas (grade de uma dimensão)');
              n := StrToIntDef(partes[2], 0);
              if n > 1 then
                RegistraErro('linha', descLinha, 'linha cadastrada mais de uma vez na grade');
              n := StrToIntDef(partes[3], 0);
              if n > 1 then
                RegistraErro('coluna', descColuna, 'coluna cadastrada mais de uma vez na grade');
            end;
          end;
        end;
      end;
    finally
      partes.Free;
    end;
  end;
end;

//Comandos da linha (iguais no banco e no script): cada um só age se o produto puder receber/tiver a grade,
//assim no script as linhas que o FDB barraria não alteram nada
procedure cImportaGrade.MontaComandos(estrutura, estoque: TStringList);
var
  dF, dM, cel, emp: string;
  comandos: TStringList;
begin
  comandos := estrutura;
  condMovimento := '';
  //Mesmas condições que o banco FDB confere antes de gravar: empresa cadastrada e coluna coerente com a grade
  //(grade de uma dimensão tem só a coluna UNICO; grade com colunas, ex.: cor, não recebe a UNICO)
  emp := 'EXISTS (SELECT 1 FROM dados_empre e2 WHERE e2.codi = ' + emprCod + ')';

  //Grade nova pela descrição (produto sem grade e sem estoque)
  if gradeDesc <> '' then
  begin
    if titLinha = '' then dF := 'NULL' else dF := '''' + TextoSQL(titLinha, 90) + '''';
    if titColuna = '' then dM := 'NULL' else dM := '''' + TextoSQL(titColuna, 90) + '''';
    comandos.Add('INSERT INTO grade (grad_id, grad_descricao, grad_tit_linha, grad_tit_coluna, ativo) ' +
      'SELECT gen_id(gen_grade_id,1), ''' + TextoSQL(gradeDesc, 90) + ''', ' + dF + ', ' + dM + ', 1 FROM RDB$DATABASE ' +
      'WHERE NOT EXISTS (SELECT 1 FROM grade g WHERE UPPER(g.grad_descricao) = ''' + TextoSQL(gradeDesc, 90) + ''') ' +
      'AND EXISTS (SELECT 1 FROM prod p WHERE p.codi = ' + prodCod + ' AND p.grade_grad_id IS NULL) ' +
      'AND ' + SemEstoque + ' AND ' + emp + ';');
  end;

  //Vínculo da grade com o produto
  if (gradeCod <> '') or (gradeDesc <> '') then
    comandos.Add('UPDATE prod SET grade_grad_id = ' + GradeValor + ' ' +
      'WHERE codi = ' + prodCod + ' AND grade_grad_id IS NULL AND ' + GradeValor + ' IS NOT NULL AND ' + SemEstoque + ' ' +
      'AND ' + emp + ' AND ' + Dimensao(GradeValor) + ';');

  //Linha e coluna cadastradas na grade do produto quando ainda não existirem
  comandos.Add('INSERT INTO grade_linha (grli_id, grad_id, grli_descricao) ' +
    'SELECT gen_id(gen_grade_linha_id,1), p.grade_grad_id, ''' + TextoSQL(descLinha, 90) + ''' FROM prod p ' +
    'WHERE ' + ProdPrincipal + ' AND ' + CondGrade + ' AND ' + emp + ' AND ' + Dimensao('p.grade_grad_id') + ' ' +
    'AND NOT EXISTS (SELECT 1 FROM grade_linha l WHERE l.grad_id = p.grade_grad_id AND UPPER(l.grli_descricao) = ''' + TextoSQL(descLinha, 90) + ''');');
  comandos.Add('INSERT INTO grade_coluna (grco_id, grad_id, grco_descricao) ' +
    'SELECT gen_id(gen_grade_coluna_id,1), p.grade_grad_id, ''' + TextoSQL(descColuna, 90) + ''' FROM prod p ' +
    'WHERE ' + ProdPrincipal + ' AND ' + CondGrade + ' AND ' + emp + ' AND ' + Dimensao('p.grade_grad_id') + ' ' +
    'AND NOT EXISTS (SELECT 1 FROM grade_coluna c WHERE c.grad_id = p.grade_grad_id AND UPPER(c.grco_descricao) = ''' + TextoSQL(descColuna, 90) + ''');');

  //Estoque da grade do produto: todas as combinações em todas as empresas (zero)
  comandos.Add(Format(SQL_MATRIZ, [ProdPrincipal + ' AND ' + CondGrade + ' AND ' + emp + ' AND ' + Dimensao('p.grade_grad_id')]));

  //Movimentação do estoque: lança só a diferença até o saldo da planilha
  //Fiscal: tipo 5 (entrada) / 2 (saída); físico: tipo 4 (entrada) / 1 (saída), igual ao ajuste de estoque do ERP
  if (qtdFiscal = '') and (qtdFisico = '') then Exit;
  comandos := estoque;
  cel := Celula;
  if qtdFiscal <> '' then dF := '(' + qtdFiscal + ' - ge.gres_qtd)' else dF := '0';
  if qtdFisico <> '' then dM := '(' + qtdFisico + ' - ge.gres_qtd_max)' else dM := '0';
  //Há diferença entre o saldo da planilha e o estoque atual da combinação
  condMovimento := 'SELECT COUNT(*) ' + cel + ' AND (' + dF + ' <> 0 OR ' + dM + ' <> 0)';

  comandos.Add('INSERT INTO grade_lancamento (grla_id, grade_grad_id, grla_origem_tabela, grla_origem_id, prod_codi, prod_empr, grla_data, grla_empr, grla_origem_tabela_cabecalho) ' +
    'SELECT {LANCAMENTO_NOVO}, p.grade_grad_id, ''GRADE_AJUS'', {AJUSTE_GRADE}, p.codi, p.empr, CURRENT_DATE, ' + emprCod + ', ''GRADE_ESTOQUE'' ' +
    cel + ' AND (' + dF + ' <> 0 OR ' + dM + ' <> 0);');
  //ITENS do ajuste (PROD_AJUS): atualizam o estoque total do produto (PROD_ESTO) pela trigger
  if qtdFiscal <> '' then
    comandos.Add('INSERT INTO itens (codi, prodcod, nume, tipo, epv, qtd, empr) ' +
      'SELECT gen_id(gen_itens_id,1), p.codi, {AJUSTE_PROD}, CASE WHEN ' + dF + ' > 0 THEN 5 ELSE 2 END, ''A'', ABS(' + dF + '), ' + emprCod + ' ' +
      cel + ' AND ' + dF + ' <> 0;');
  if qtdFisico <> '' then
    comandos.Add('INSERT INTO itens (codi, prodcod, nume, tipo, epv, qtd, empr) ' +
      'SELECT gen_id(gen_itens_id,1), p.codi, {AJUSTE_PROD}, CASE WHEN ' + dM + ' > 0 THEN 4 ELSE 1 END, ''A'', ABS(' + dM + '), ' + emprCod + ' ' +
      cel + ' AND ' + dM + ' <> 0;');
  //Itens do lançamento da grade: a trigger recalcula o GRADE_ESTOQUE (o fiscal não altera o físico)
  if qtdFiscal <> '' then
    comandos.Add('INSERT INTO grade_lancamento_itens (grit_id, grla_id, grli_id, grco_id, grit_qtd, grit_tipo, grit_qtd_saldo) ' +
      'SELECT gen_id(gen_grade_lancamento_itens_id,1), {LANCAMENTO}, ge.grli_id, ge.grco_id, ABS(' + dF + '), CASE WHEN ' + dF + ' > 0 THEN 5 ELSE 2 END, 0 ' +
      cel + ' AND ' + dF + ' <> 0;');
  if qtdFisico <> '' then
    comandos.Add('INSERT INTO grade_lancamento_itens (grit_id, grla_id, grli_id, grco_id, grit_qtd, grit_tipo, grit_qtd_saldo) ' +
      'SELECT gen_id(gen_grade_lancamento_itens_id,1), {LANCAMENTO}, ge.grli_id, ge.grco_id, ABS(' + dM + '), CASE WHEN ' + dM + ' > 0 THEN 4 ELSE 1 END, 0 ' +
      cel + ' AND ' + dM + ' <> 0;');
end;

//Códigos do ajuste e do lançamento: no banco os códigos reservados; no script o generator/consulta
function cImportaGrade.Substitui(cmd, lancamento: string): string;
begin
  Result := cmd;
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then
  begin
    Result := stringreplace(Result, '{LANCAMENTO_NOVO}', lancamento, [rfReplaceAll]);
    Result := stringreplace(Result, '{LANCAMENTO}', lancamento, [rfReplaceAll]);
    Result := stringreplace(Result, '{AJUSTE_GRADE}', ajusteGrade, [rfReplaceAll]);
    Result := stringreplace(Result, '{AJUSTE_PROD}', ajusteProd, [rfReplaceAll]);
  end
  else begin
    Result := stringreplace(Result, '{LANCAMENTO_NOVO}', 'gen_id(gen_grade_lancamento_id,1)', [rfReplaceAll]);
    //Lançamento desta combinação no ajuste da importação (o último criado)
    Result := stringreplace(Result, '{LANCAMENTO}', '(SELECT MAX(gl.grla_id) FROM grade_lancamento gl WHERE gl.grla_origem_tabela = ''GRADE_AJUS'' ' +
      'AND gl.grla_origem_id = gen_id(gen_grade_ajus_id,0) AND gl.prod_codi = ' + prodCod + ' AND gl.grla_empr = ' + emprCod + ')', [rfReplaceAll]);
    Result := stringreplace(Result, '{AJUSTE_GRADE}', 'gen_id(gen_grade_ajus_id,0)', [rfReplaceAll]);
    Result := stringreplace(Result, '{AJUSTE_PROD}', 'gen_id(gen_prod_ajus_id,0)', [rfReplaceAll]);
  end;
end;

procedure cImportaGrade.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
  gen: TSQLQuery;
  tr: TDBXTransaction;
  estrutura, estoque: TStringList;
  j: Integer;
  cmd, lancamento, novoAjusteGrade, novoAjusteProd: string;
  criaAjuste: Boolean;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //Valores inválidos na linha: para a importação para corrigir e continuar desta linha
  if errosLinha.Count > 0 then
  begin
    ShowMessage('Linha ' + IntToStr(k) + ' com valores inválidos:' + sLineBreak + sLineBreak +
                errosLinha.Text + sLineBreak +
                'Corrija na planilha e clique em Importar para continuar a partir desta linha.');
    status := 0;
    Exit;
  end;

  estrutura := TStringList.Create;
  estoque := TStringList.Create;
  try
    MontaComandos(estrutura, estoque);
    criaAjuste := (estoque.Count > 0) and not temAjuste;

    //----------------------------------------
    //Gravar no banco: a linha inteira numa transação (grava tudo ou nada)
    if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
      SQL := TSQLDataSet.Create(Nil);
      gen := TSQLQuery.Create(Nil);
      cmd := '';
      try
        frmPrinc.conDestino.Open;
        SQL.SQLConnection := frmPrinc.conDestino;
        gen.SQLConnection := frmPrinc.conDestino;
        tr := frmPrinc.conDestino.BeginTransaction;
        try
          frmImportando.atualizaStatus('Inserindo dados da grade.');
          //Grade, vínculo com o produto, linha, coluna e estoque da grade (zero) do produto
          for j := 0 to estrutura.Count-1 do
          begin
            cmd := estrutura[j];
            SQL.CommandText := cmd;
            SQL.ExecSQL;
          end;
          //Estoque: só movimenta (e só cria o ajuste) se o saldo da planilha for diferente do atual
          if estoque.Count > 0 then
          begin
            cmd := condMovimento;
            gen.SQL.Text := cmd;
            gen.Open;
            if gen.Fields[0].AsInteger = 0 then
            begin
              estoque.Clear;
              criaAjuste := False;
            end;
            gen.Close;
          end;
          novoAjusteGrade := ajusteGrade;
          novoAjusteProd := ajusteProd;
          lancamento := 'NULL';
          if estoque.Count > 0 then
          begin
            //Ajuste da importação: cabeçalho da grade (GRAD_AJUS) e do estoque (PROD_AJUS), um para a importação inteira
            if criaAjuste then
            begin
              gen.SQL.Text := 'SELECT gen_id(gen_grade_ajus_id,1) FROM RDB$DATABASE';
              gen.Open;
              novoAjusteGrade := gen.Fields[0].AsString;
              gen.Close;
              gen.SQL.Text := 'SELECT gen_id(gen_prod_ajus_id,1) FROM RDB$DATABASE';
              gen.Open;
              novoAjusteProd := gen.Fields[0].AsString;
              gen.Close;
              cmd := 'INSERT INTO grad_ajus (graj_id, graj_data) VALUES (' + novoAjusteGrade + ', CURRENT_DATE);';
              SQL.CommandText := cmd;
              SQL.ExecSQL;
              cmd := 'INSERT INTO prod_ajus (codi, data, obs) VALUES (' + novoAjusteProd + ', CURRENT_DATE, ''' + OBS_AJUSTE + ''');';
              SQL.CommandText := cmd;
              SQL.ExecSQL;
            end;
            //Código do lançamento reservado antes (outro lançamento feito ao mesmo tempo no ERP não interfere)
            gen.SQL.Text := 'SELECT gen_id(gen_grade_lancamento_id,1) FROM RDB$DATABASE';
            gen.Open;
            lancamento := gen.Fields[0].AsString;
            gen.Close;
          end;
          ajusteGrade := novoAjusteGrade;
          ajusteProd := novoAjusteProd;
          for j := 0 to estoque.Count-1 do
          begin
            cmd := Substitui(estoque[j], lancamento);
            SQL.CommandText := cmd;
            SQL.ExecSQL;
          end;
          frmPrinc.conDestino.CommitFreeAndNil(tr);
          if criaAjuste then temAjuste := True;
        except
          on e: Exception do
          begin
            frmPrinc.conDestino.RollbackFreeAndNil(tr);
            if not temAjuste then
            begin
              ajusteGrade := '';
              ajusteProd := '';
            end;
            ShowMessage('Erro SQL: '+e.message+sLineBreak+cmd);
            status := 0;
            Exit;
          end;
        end;
      finally
        gen.Free;
        SQL.Free;
        frmPrinc.conDestino.Close;
      end;
    end
    //Gravar comandos em TXT
    else if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
            (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
    then begin
      AssignFile(fileTXT, frmPrinc.DBPath.Text);
      try
        //Carregar arquivo TXT
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);
      except
        on e: exception do
        begin
          ShowMessage('Erro TXT: '+e.message);
          status := 0;
          Exit;
        end;
      end;
      try
        try
          frmImportando.atualizaStatus('Comandos da grade.');
          if criaAjuste then
          begin
            WriteLn(fileTXT, '----------Ajuste de estoque das grades----------');
            WriteLn(fileTXT, 'INSERT INTO grad_ajus (graj_id, graj_data) VALUES (gen_id(gen_grade_ajus_id,1), CURRENT_DATE);');
            WriteLn(fileTXT, 'INSERT INTO prod_ajus (codi, data, obs) VALUES (gen_id(gen_prod_ajus_id,1), CURRENT_DATE, ''' + OBS_AJUSTE + ''');');
            WriteLn(fileTXT, 'COMMIT WORK;');
            temAjuste := True;
          end;
          WriteLn(fileTXT, '----------Grade linha ' + IntToStr(k) + '----------');
          for j := 0 to estrutura.Count-1 do
            WriteLn(fileTXT, estrutura[j]);
          for j := 0 to estoque.Count-1 do
            WriteLn(fileTXT, Substitui(estoque[j], ''));
          WriteLn(fileTXT, 'COMMIT WORK;');
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
  finally
    estoque.Free;
    estrutura.Free;
  end;
end;

end.
