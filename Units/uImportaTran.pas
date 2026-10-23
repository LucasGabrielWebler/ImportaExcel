unit uImportaTran;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaTran = class
    private
      colTran, dadosTran, condUpdateTran, dadosUpdateTran: String;
      procedure Campo(coluna, campo, valor: string);

    public
      k: Integer;
      StringGrid1: TStringGrid;
      constructor ImportaRegistro(numReg: Integer; Grid: TStringGrid);
      procedure CarregaColunas;
      procedure Gravar;
  end;

implementation

uses
  importa_excel, importando;

constructor cImportaTran.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

//Inclui o campo no INSERT e, no Update, na condição (coluna chave) ou nos dados alterados
//valor já vem pronto para o SQL (entre aspas, número ou subconsulta)
procedure cImportaTran.Campo(coluna, campo, valor: string);
begin
  colTran := colTran + ',' + campo;
  dadosTran := dadosTran + ',' + valor;
  if VerificaUpdate(coluna) = 1 then
  begin
    if condUpdateTran <> '' then condUpdateTran := condUpdateTran + ' AND ';
    condUpdateTran := condUpdateTran + campo + '=' + valor;
  end
  else begin
    if dadosUpdateTran <> '' then dadosUpdateTran := dadosUpdateTran + ', ';
    dadosUpdateTran := dadosUpdateTran + campo + '=' + valor;
  end;
end;

//Campos do cadastro principal da transportadora (tabela TRAN), iguais à tela do ERP
//A aba Adicionais (CT-e/MDF-e) não é importada: fica com os padrões da tabela
procedure cImportaTran.CarregaColunas;
var
  i: Integer;
  nome, temp, uf, cidade, tipoDoc, colunaDoc: string;
  bancoFDB: Boolean;
begin
  colTran := '';
  dadosTran := '';
  condUpdateTran := '';
  dadosUpdateTran := '';
  errosLinha.Clear;
  bancoFDB := UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB';
  tipoDoc := '';
  colunaDoc := '';

  frmImportando.atualizaStatus('Transportadora '+ IntToStr(k));

  //CODI: sem a coluna usa o generator; no Update é usado somente como chave, nunca é alterado
  i := BuscaColuna(StringGrid1,'codi');
  if i <> -1 then
  begin
    temp := Trim(stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll]));
    if temp = '' then
      RegistraErro('codi', '', 'código em branco')
    else if StrToIntDef(temp,-1) <= 0 then
      RegistraErro('codi', StringGrid1.Cells[i,k], 'código numérico inválido')
    else begin
      temp := IntToStr(StrToInt(temp));
      //Inclusão no banco FDB: o código não pode já estar cadastrado
      if bancoFDB and (colUpdateCount <= 0) and (querySelect('SELECT codi FROM tran WHERE codi = ' + temp) <> '') then
        RegistraErro('codi', StringGrid1.Cells[i,k], 'código já cadastrado');
    end;
    colTran := 'codi';
    dadosTran := temp;
    if VerificaUpdate('codi') = 1 then
      condUpdateTran := 'codi=' + temp;
  end
  else begin
    colTran := 'codi';
    dadosTran := 'gen_id(gen_tran_id,1)';
  end;

  //Cidade: pelo nome da cidade + UF (como nos clientes) ou pelo código (CODI_CIDA)
  //O ERP grava somente o código da cidade
  i := BuscaColuna(StringGrid1,'cida');
  if (i <> -1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
  begin
    cidade := Corta(UpperCase(RemoveAcento(Trim(StringGrid1.Cells[i,k]))), 60, 'cida');
    uf := '';
    if BuscaColuna(StringGrid1,'uf') <> -1 then
      uf := UpperCase(RemoveAcento(Trim(StringGrid1.Cells[BuscaColuna(StringGrid1,'uf'),k])));
    if uf = '' then
      RegistraErro('cida', StringGrid1.Cells[i,k], 'cidade sem UF informada')
    else if Length(uf) <> 2 then
      RegistraErro('uf', uf, 'UF inválida (precisa ter 2 letras)')
    else begin
      temp := frmPrinc.buscaCidade(cidade, uf);
      if temp = '' then
        RegistraErro('cida', StringGrid1.Cells[i,k], 'cidade não encontrada para a UF informada')
      else begin
        colTran := colTran + ',codi_cida';
        dadosTran := dadosTran + ',' + temp;
        if (VerificaUpdate('cida') = 1) or (VerificaUpdate('uf') = 1) then
        begin
          if condUpdateTran <> '' then condUpdateTran := condUpdateTran + ' AND ';
          condUpdateTran := condUpdateTran + 'codi_cida=' + temp;
        end
        else begin
          if dadosUpdateTran <> '' then dadosUpdateTran := dadosUpdateTran + ', ';
          dadosUpdateTran := dadosUpdateTran + 'codi_cida=' + temp;
        end;
      end;
    end;
  end;

  for i := 0 to StringGrid1.ColCount-1 do
  begin
    nome := LowerCase(Trim(StringGrid1.Cells[i,0]));
    temp := Trim(StringGrid1.Cells[i,k]);

    //NOME (obrigatório; maiúsculas com acento, padrão do ERP)
    if nome = 'nome' then
    begin
      if temp = '' then
        RegistraErro('nome', '', 'nome não informado')
      else
        Campo('nome', 'nome', '''' + TextoSQL(Corta(AnsiUpperCase(temp), 50, 'nome'), 50) + '''');
    end
    //Textos: vazio não altera o campo
    else if (nome = 'ende') and (temp <> '') then
      Campo('ende', 'ende', '''' + TextoSQL(Corta(AnsiUpperCase(temp), 50, 'ende'), 50) + '''')
    else if (nome = 'insc') and (temp <> '') then
      Campo('insc', 'insc', '''' + TextoSQL(Corta(AnsiUpperCase(temp), 20, 'insc'), 20) + '''')
    else if (nome = 'fone') and (temp <> '') then
      Campo('fone', 'fone', '''' + TextoSQL(Corta(temp, 21, 'fone'), 21) + '''')
    //E-mail gravado como digitado (o ERP não converte para maiúsculas)
    else if (nome = 'email') and (temp <> '') then
      Campo('email', 'email', '''' + TextoSQL(Corta(temp, 255, 'email'), 255) + '''')
    else if ((nome = 'placa') or (nome = 'placa2') or (nome = 'placa3')) and (temp <> '') then
      Campo(nome, nome, '''' + TextoSQL(Corta(AnsiUpperCase(temp), 8, nome), 8) + '''')
    //ATIVO: S/1/A = ativo, N/0/I/2 = inativo (igual aos clientes); vazio fica ativo
    else if (nome = 'ativo') and (temp <> '') then
    begin
      temp := UpperCase(temp);
      if (temp = 'S') or (temp = '1') or (temp = 'A') then Campo('ativo', 'ativo', '''S''')
      else if (temp = 'N') or (temp = '0') or (temp = 'I') or (temp = '2') then Campo('ativo', 'ativo', '''N''')
      else RegistraErro('ativo', StringGrid1.Cells[i,k], 'valor inválido (use S ou N)');
    end
    //CPF / CNPJ: formatados como no ERP e definem o tipo (Física / Jurídica)
    else if ((nome = 'cpf') or (nome = 'cnpj') or (nome = 'cpf_cnpj')) and (temp <> '') then
    begin
      temp := stringreplace(temp, '-', '',[rfReplaceAll]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll]);
      if (Length(temp) = 11) and SoNumeros(temp) and (nome <> 'cnpj') then
      begin
        Campo(nome, 'cpf', '''' + Copy(temp,1,3) + '.' + Copy(temp,4,3) + '.' + Copy(temp,7,3) + '-' + Copy(temp,10,2) + '''');
        tipoDoc := 'F';
        colunaDoc := nome;
      end
      else if (Length(temp) = 14) and SoNumeros(temp) and (nome <> 'cpf') then
      begin
        Campo(nome, 'cnpj', '''' + Copy(temp,1,2) + '.' + Copy(temp,3,3) + '.' + Copy(temp,6,3) + '/' + Copy(temp,9,4) + '-' + Copy(temp,13,2) + '''');
        tipoDoc := 'J';
        colunaDoc := nome;
      end
      else if nome = 'cpf' then
        RegistraErro('cpf', StringGrid1.Cells[i,k], 'CPF inválido (precisa ter 11 dígitos)')
      else if nome = 'cnpj' then
        RegistraErro('cnpj', StringGrid1.Cells[i,k], 'CNPJ inválido (precisa ter 14 dígitos)')
      else
        RegistraErro('cpf_cnpj', StringGrid1.Cells[i,k], 'CPF/CNPJ inválido (precisa ter 11 ou 14 dígitos)');
    end
    //OBS ou OBSERVACAO (nome do campo no banco): texto livre sem limite de tamanho, gravado como digitado (igual aos clientes)
    else if ((nome = 'obs') or (nome = 'observacao')) and (temp <> '') then
      Campo(nome, 'observacao', '''' + TextoSQL(StringGrid1.Cells[i,k], MaxInt) + '''')
    //CODI_CIDA: código da cidade direto da planilha
    else if (nome = 'codi_cida') and (temp <> '') then
    begin
      if StrToIntDef(temp,-1) <= 0 then
        RegistraErro('codi_cida', temp, 'código de cidade inválido')
      else if bancoFDB and (querySelect('SELECT cid_codi FROM cidade WHERE cid_codi = ' + IntToStr(StrToInt(temp))) = '') then
        RegistraErro('codi_cida', temp, 'código de cidade não encontrado')
      else
        Campo('codi_cida', 'codi_cida', IntToStr(StrToInt(temp)));
    end;
  end;

  //Tipo (Física / Jurídica) pelo último CPF/CNPJ válido da linha, gravado uma única vez
  //Inclusão sem CPF/CNPJ: tipo Jurídica, igual ao padrão da tela do ERP
  if tipoDoc <> '' then
    Campo(colunaDoc, 'tipo', '''' + tipoDoc + '''')
  else if colUpdateCount <= 0 then
  begin
    colTran := colTran + ',tipo';
    dadosTran := dadosTran + ',''J''';
  end;
end;

procedure cImportaTran.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
  j: Integer;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //Update: todas as colunas chave precisam estar preenchidas (uma chave vazia ampliaria o filtro)
  for j := 0 to colUpdateCount-1 do
    if (BuscaColuna(StringGrid1, colUpdate[j]) <> -1) and (Trim(StringGrid1.Cells[BuscaColuna(StringGrid1, colUpdate[j]), k]) = '') then
      RegistraErro(colUpdate[j], '', 'coluna chave (Update) em branco');

  //Update sem valor na(s) coluna(s) chave: não gera SQL inválido
  if (colUpdateCount > 0) and (condUpdateTran = '') then
    RegistraErro('Update', '', 'coluna marcada como Update (chave) está em branco');

  //Valores inválidos na linha: para a importação para corrigir e continuar desta linha
  if errosLinha.Count > 0 then
  begin
    ShowMessage('Linha ' + IntToStr(k) + ' com valores inválidos:' + sLineBreak + sLineBreak +
                errosLinha.Text + sLineBreak +
                'Corrija na planilha e clique em Importar para continuar a partir desta linha.');
    status := 0;
    Exit;
  end;

  //Update sem nenhum dado para alterar
  if (colUpdateCount > 0) and (dadosUpdateTran = '') then
    Exit;

  //----------------------------------------
  //Gravar no banco de dados TRAN
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    SQL := TSQLDataSet.Create(Nil);
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL.SQLConnection := frmPrinc.conDestino;

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          frmImportando.atualizaStatus('Inserindo dados na tabela TRAN.');
          SQL.CommandText := 'INSERT INTO tran ('+ colTran +') VALUES ' + '(' + dadosTran + ');';
          SQL.ExecSQL;
        end
        //Se for UPDATE
        else begin
          frmImportando.atualizaStatus('Atualizando dados na tabela TRAN.');
          SQL.CommandText := 'UPDATE tran SET '+ dadosUpdateTran +' WHERE ' + condUpdateTran + ';';
          SQL.ExecSQL;
        end;

      except
        on e: exception do
        begin
          ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
          status := 0;
          //SQL e conexão são liberados no finally
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
        frmImportando.atualizaStatus('Comandos da TRAN.');
        WriteLn(fileTXT, '----------Comandos da TRAN----------');

        //Se for INSERT
        if colUpdateCount <= 0 then
          WriteLn(fileTXT, 'INSERT INTO tran ('+ colTran +') VALUES ' + '(' + dadosTran + ');')
        //Se for UPDATE
        else
          WriteLn(fileTXT, 'UPDATE tran SET '+ dadosUpdateTran +' WHERE ' + condUpdateTran + ';');
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
end;

end.
