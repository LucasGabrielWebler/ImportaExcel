unit uImportaClieForn;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaClieForn = class
    private
      colClieForn, dadosClieForn, condUpdateClieForn, dadosUpdateClieForn: String;

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

constructor cImportaClieForn.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaClieForn.CarregaColunas;
var
  i, count: Integer;
  temp, temp2: string;
  fileTXT: TextFile;
begin
  colClieForn := '';
  dadosClieForn := '';
  condUpdateClieForn := '';
  dadosUpdateClieForn := '';
  errosLinha.Clear;
  //Textos em maiúsculas mantendo os acentos (padrão do ERP); CIDA sem acento, igual à tabela CIDADE

  frmImportando.atualizaStatus('Clie/Forn '+ IntToStr(k));

  //Codigo é obrigatório, se não tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    if (Trim(StringGrid1.Cells[i,k])='') then
    begin
      RegistraErro('codi', '', 'código em branco');
    end
    else begin
      StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
      //Código precisa ser numérico (texto daria erro de conversão no banco)
      if not IsNumeric(StringGrid1.Cells[i,k]) then
        RegistraErro('codi', StringGrid1.Cells[i,k], 'código numérico inválido');
      colClieForn := colClieForn + 'codi';
      dadosClieForn := dadosClieForn + '''' + StringGrid1.Cells[i,k] + '''';
      //Testa se é Update
      //No Update o CODI é usado somente como chave, nunca é alterado
      if VerificaUpdate('codi') = 1 then
        condUpdateClieForn := condUpdateClieForn + 'codi=' + '''' + StringGrid1.Cells[i,k] + '''';
    end;
  end
  else begin
    colClieForn := colClieForn + 'codi';
    dadosClieForn := dadosClieForn + 'gen_id(gen_clieforn_id,1)';
  end;

  //Importar cidade se tiver, precisa ter o UF antes de importar a cidade
  //UF
  temp := '';
  i:=BuscaColuna(StringGrid1,'uf');
  if (i<>-1) then
  begin
    if (StringGrid1.Cells[i,k]<>'') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,2,'uf'),2); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',uf';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('uf') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'uf=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'uf=' + '''' + temp + '''';
      end;
    end;
  end;
  //CIDA (CIDADE)
  i:=BuscaColuna(StringGrid1,'cida');
  if ((i<>-1) and (temp<>'')) then
  begin
    temp2 := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp2 := Corta(temp2,60,'cida');
    //Busca com o nome original (ex.: SAO JORGE D'OESTE); o apóstrofo é tratado na busca
    temp := frmPrinc.buscaCidade(temp2, temp);
    //Apóstrofo do nome da cidade para o SQL
    temp2 := TextoSQL(temp2,60);
    if (temp <> '') then
    begin
      colClieForn := colClieForn + ',cida,codi_cida';
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      //Código da cidade (FDB) ou subconsulta da cidade (TXT/SQL), sem aspas
      dadosClieForn := dadosClieForn + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('cida') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'cida=' + '''' + temp2 + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'cida=' + '''' + temp2 + ''', codi_cida=' + temp;
      end;
    end
    else if Trim(StringGrid1.Cells[i,k]) <> '' then
      RegistraErro('cida', StringGrid1.Cells[i,k], 'cidade não encontrada para a UF informada');
  end
  //Cidade informada sem UF
  else if (i<>-1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
    RegistraErro('cida', StringGrid1.Cells[i,k], 'cidade sem UF informada');


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //EMPRESA
    if ((LowerCase(StringGrid1.Cells[i,0])='empresa') and (StringGrid1.Cells[i,0]<>'')) then
    begin
      //Número inteiro, se vazio ou inválido não grava o campo
      temp := Trim(StringGrid1.Cells[i,k]);
      if not IsNumeric(temp) then temp := '';
      if temp <> '' then
      begin
        colClieForn := colClieForn + ',empresa';
        dadosClieForn := dadosClieForn + ',' + temp;
        //Testa se é Update
        if VerificaUpdate('empresa') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'empresa=' + temp;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'empresa=' + temp;
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('empresa', StringGrid1.Cells[i,k], 'valor numérico inválido');
    end
    //GRUPO (vazio é ignorado, não cadastra grupo em branco)
    else if ((LowerCase(StringGrid1.Cells[i,0])='grupo') and (StringGrid1.Cells[i,0]<>'') and (Trim(StringGrid1.Cells[i,k])<>'')) then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,50,'grupo');
      //Se for letras, buscar código.
      if not (IsNumeric(temp)) then
      begin
        //Arquivo TXT/SQL: cadastra o grupo somente se ainda não existir e o código
        //é buscado pela descrição na execução do script
        if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
           (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL') then
        begin
          AssignFile(fileTXT, frmPrinc.DBPath.Text);
          if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
          else append(fileTXT);
          WriteLn(fileTXT, 'INSERT INTO grupo_cliente (CODI,DESCR,COMISSAO) SELECT gen_id(gen_grupo_cliente_id,1),'''+temp+''',1 FROM RDB$DATABASE WHERE NOT EXISTS (SELECT 1 FROM grupo_cliente gc WHERE gc.descr='''+temp+''');');
          WriteLn(fileTXT, 'COMMIT WORK;');
          CloseFile(fileTXT);
          temp2 := '(SELECT FIRST 1 gc.codi FROM grupo_cliente gc WHERE gc.descr='''+temp+''')';
        end
        //Banco FDB: busca o código pela descrição, se não encontrar cadastra o grupo e busca novamente
        else begin
          temp2 := querySelect('SELECT g.codi FROM grupo_cliente g WHERE g.descr = '''+temp+'''');
          if temp2='' then begin
            queryInsert('INSERT INTO grupo_cliente (CODI,DESCR,COMISSAO) VALUES (gen_id(gen_grupo_cliente_id,1),'''+temp+''',1);');
            temp2 := querySelect('SELECT g.codi FROM grupo_cliente g WHERE g.descr = '''+temp+'''');
          end;
          temp2 := '''' + temp2 + '''';
        end;

        colClieForn := colClieForn + ',codi_grupo_clie';
        dadosClieForn := dadosClieForn + ',' + temp2;
        //Testa se é Update
        //Descrição do grupo não pode ser chave de Update (somente o código)
        if VerificaUpdate('grupo') = 1 then begin
          RegistraErro('grupo', temp, 'a descrição do grupo não pode ser chave de Update (use o código do grupo)');
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'codi_grupo_clie=' + temp2;
        end;
      end
      else begin
        //Se for números, considera como código
        //Arquivo TXT/SQL: não há banco para validar, usa o código informado
        if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
           (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL') then
          temp2 := temp
        //Banco FDB: busca se existe o código cadastrado
        else
          temp2 := querySelect('SELECT g.codi FROM grupo_cliente g WHERE g.codi = '''+temp+'''');
        //Se não encontrar o código, para a importação para corrigir
        if temp2='' then begin
          RegistraErro('grupo', temp, 'código de grupo não encontrado');
        end
        //Se encontrar usa o código
        else begin
          colClieForn := colClieForn + ',codi_grupo_clie';
          dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        end;
        //Testa se é Update
        if VerificaUpdate('grupo') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'codi_grupo_clie=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'codi_grupo_clie=' + '''' + temp + '''';
        end;
      end;
    end
    //NOME
    else if ((LowerCase(StringGrid1.Cells[i,0])='nome') and (StringGrid1.Cells[i,0]<>'')) then
    begin
      colClieForn := colClieForn + ',nome';
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      //Nome (razão social) é obrigatório; o nome fantasia pode ficar vazio
      if Trim(temp) = '' then RegistraErro('nome', '', 'nome (razão social) não informado');
      temp2 := TextoSQL(Corta(temp,60,'nome'),60); //Apóstrofo e tamanho do campo
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      //Testa se é Update
      if VerificaUpdate('nome') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'nome=' + '''' + temp2 + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'nome=' + '''' + temp2 + '''';
      end;

      //Testar se não existe coluna NOME_FANT, se não existir joga o valor da NOME
      count:=BuscaColuna(StringGrid1,'nome_fant');
      if (count=-1) then
      begin
        colClieForn := colClieForn + ',nome_fant';
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      end;
    end
    //NOME_FANT (NOME FANTASIA)
    else if (LowerCase(StringGrid1.Cells[i,0])='nome_fant') then
    begin
      colClieForn := colClieForn + ',nome_fant';
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp2 := TextoSQL(Corta(temp,60,'nome_fant'),60); //Apóstrofo e tamanho do campo
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      //Testa se é Update
      if VerificaUpdate('nome_fant') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'nome_fant=' + '''' + temp2 + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'nome_fant=' + '''' + temp2 + '''';
      end;

      //Testar se não existe coluna NOME, se não existir joga o valor da NOME_FANT
      count:=BuscaColuna(StringGrid1,'nome');
      if (count=-1) then
      begin
        colClieForn := colClieForn + ',nome';
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Inclusão sem a coluna NOME: o nome fantasia vira o nome (razão social), então é obrigatório
        if (colUpdateCount <= 0) and (Trim(temp) = '') then
          RegistraErro('nome_fant', '', 'nome não informado (sem a coluna NOME, o NOME_FANT é gravado como nome)');
      end;
    end
    //DATA_NASC
    else if (LowerCase(StringGrid1.Cells[i,0])='data_nasc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      //Data inválida não grava o campo
      if not DataValida(temp) then
      begin
        if temp <> '' then RegistraErro('data_nasc', StringGrid1.Cells[i,k], 'data inválida');
        temp := '';
      end;
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colClieForn := colClieForn + ',data_nasc';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('data_nasc') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'data_nasc=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'data_nasc=' + '''' + temp + '''';
        end;
      end
      else if temp.Length = 6 then begin
        //Ano com 2 dígitos: maior que o ano atual é do século passado (em 2026: 85 = 1985, 10 = 2010)
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ AnoQuatroDigitos(Copy(temp,5,2));
        colClieForn := colClieForn + ',data_nasc';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('data_nasc') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'data_nasc=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'data_nasc=' + '''' + temp + '''';
        end;
      end;
    end
    //CPF
    else if (LowerCase(StringGrid1.Cells[i,0])='cpf') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if (temp.Length = 11) and SoNumeros(temp) then
      begin
        temp := (Copy(temp,1,3))+ '.' + (Copy(temp,4,3)) + '.' + (Copy(temp,7,3)) + '-' + (Copy(temp,10,2));
        colClieForn := colClieForn + ',cpf';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Setar tipo (Fisca ou Juridica)
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',''' + 'F' + '''';
        //Testa se é Update
        if VerificaUpdate('cpf') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cpf=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cpf=' + '''' + temp + '''';
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + '''' + 'F' + '''';
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('cpf', StringGrid1.Cells[i,k], 'CPF inválido (precisa ter 11 dígitos)');
    end
    //CNPJ
    else if (LowerCase(StringGrid1.Cells[i,0])='cnpj') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if (temp.Length = 14) and SoNumeros(temp) then
      begin
        temp := (Copy(temp,1,2))+ '.' + (Copy(temp,3,3)) + '.' + (Copy(temp,6,3)) + '/' + (Copy(temp,9,4)) + '-' + (Copy(temp,13,2));
        colClieForn := colClieForn + ',cnpj';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Setar tipo (Fisca ou Juridica)
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',''' + 'J' + '''';
        //Testa se é Update
        if VerificaUpdate('cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cnpj=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cnpj=' + '''' + temp + '''';
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + '''' + 'J' + '''';
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('cnpj', StringGrid1.Cells[i,k], 'CNPJ inválido (precisa ter 14 dígitos)');
    end
    //CPF OU CNPJ NO MESMO CAMPO
    else if (LowerCase(StringGrid1.Cells[i,0])='cpf_cnpj') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if (temp.Length = 11) and SoNumeros(temp) then
      begin
        temp2 := (Copy(temp,1,3))+ '.' + (Copy(temp,4,3)) + '.' + (Copy(temp,7,3)) + '-' + (Copy(temp,10,2));
        colClieForn := colClieForn + ',cpf';
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Setar tipo (Fisca ou Juridica)
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',''' + 'F' + '''';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cpf=' + '''' + temp2 + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cpf=' + '''' + temp2 + '''';
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + '''' + 'F' + '''';
        end;
      end
      else if (temp.Length = 14) and SoNumeros(temp) then
      begin
        temp2 := (Copy(temp,1,2))+ '.' + (Copy(temp,3,3)) + '.' + (Copy(temp,6,3)) + '/' + (Copy(temp,9,4)) + '-' + (Copy(temp,13,2));
        colClieForn := colClieForn + ',cnpj';
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Setar tipo (Fisca ou Juridica)
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',''' + 'J' + '''';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cnpj=' + '''' + temp2 + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cnpj=' + '''' + temp2 + '''';
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + '''' + 'J' + '''';
        end;
      end
      else
      begin
        //Inclusão: vazio grava '0' (padrão do ERP), senão grava no CPF o valor que veio, com tipo F
        temp2 := Copy(Trim(StringGrid1.Cells[i,k]),1,16);
        if temp2 = '' then temp2 := '0';
        temp2 := stringreplace(temp2, '''', '''''',[rfReplaceAll, rfIgnoreCase]);
        colClieForn := colClieForn + ',cpf';
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Update: CPF NULL
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cpf IS NULL';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cpf=' + 'null';
        end;

        //Inclusão: CNPJ '0' (padrão do ERP) / Update: CNPJ null
        colClieForn := colClieForn + ',cnpj';
        dadosClieForn := dadosClieForn + ',''' + '0' + '''';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cnpj IS NULL';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cnpj=' + 'null';
        end;

        //Setar tipo (Fisca ou Juridica)
        //Inclusão: tipo F / Update: tipo null
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',''' + 'F' + '''';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'tipo IS NULL';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'tipo=' + 'null';
        end;
      end;
    end
    //RG
    else if (LowerCase(StringGrid1.Cells[i,0])='rg') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := TextoSQL(Corta(temp,16,'rg'),16); //Apóstrofo e tamanho do campo
      if temp.Length > 1 then
      begin
        colClieForn := colClieForn + ',rg';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('rg') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'rg=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'rg=' + '''' + temp + '''';
        end;
      end
      else begin
        colClieForn := colClieForn + ',rg';
        dadosClieForn := dadosClieForn + ',''' + '' + '''';
      end;
    end
    //INSC (INSCRICAO ESTADUAL-IE)
    else if (LowerCase(StringGrid1.Cells[i,0])='insc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := TextoSQL(Corta(temp,20,'insc'),20); //Apóstrofo e tamanho do campo
      if temp.Length > 1 then
      begin
        colClieForn := colClieForn + ',insc';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('insc') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'insc=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'insc=' + '''' + temp + '''';
        end;
      end
      else begin
        colClieForn := colClieForn + ',insc';
        dadosClieForn := dadosClieForn + ',''' + '' + '''';
      end;
    end
    //INSCR_PRODUTOR (INSCRICAO DE PRODUTOR RURAL)
    else if (LowerCase(StringGrid1.Cells[i,0])='inscr_produtor') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := TextoSQL(Corta(temp,60,'inscr_produtor'),60); //Apóstrofo e tamanho do campo
      if temp.Length > 1 then
      begin
        colClieForn := colClieForn + ',inscr_produtor';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('inscr_produtor') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'inscr_produtor=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'inscr_produtor=' + '''' + temp + '''';
        end;
      end
      else begin
        colClieForn := colClieForn + ',inscr_produtor';
        dadosClieForn := dadosClieForn + ',''' + '' + '''';
      end;
    end
    //ENDE (ENDERECO)
    else if (LowerCase(StringGrid1.Cells[i,0])='ende') then
    begin
      colClieForn := colClieForn + ',ende';
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'ende');
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('ende') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'ende=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'ende=' + '''' + temp + '''';
      end;
    end
    //BAIR (BAIRRO)
    else if (LowerCase(StringGrid1.Cells[i,0])='bair') then
    begin
      colClieForn := colClieForn + ',bair';
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,30,'bair');
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('bair') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'bair=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'bair=' + '''' + temp + '''';
      end;
    end
    //COMP (COMPLEMENTO)
    else if (LowerCase(StringGrid1.Cells[i,0])='comp') then
    begin
      colClieForn := colClieForn + ',comp';
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,30,'comp');
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('comp') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'comp=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'comp=' + '''' + temp + '''';
      end;
    end
    //CEP
    else if (LowerCase(StringGrid1.Cells[i,0])='cep') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := TextoSQL(Corta(temp,10,'cep'),10); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',cep';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('cep') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'cep=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'cep=' + '''' + temp + '''';
      end;
    end
    //PROX (PROXIMIDADE)
    else if (LowerCase(StringGrid1.Cells[i,0])='prox') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,60,'prox'),60); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',prox';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('prox') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'prox=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'prox=' + '''' + temp + '''';
      end;
    end
    //FONE
    else if (LowerCase(StringGrid1.Cells[i,0])='fone') then
    begin
      temp := StringGrid1.Cells[i,k];
      temp := TextoSQL(Corta(temp,21,'fone'),21); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',fone';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('fone') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'fone=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'fone=' + '''' + temp + '''';
      end;
    end
    //FONE2
    else if (LowerCase(StringGrid1.Cells[i,0])='fone2') then
    begin
      temp := StringGrid1.Cells[i,k];
      temp := TextoSQL(Corta(temp,21,'fone2'),21); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',fone2';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('fone2') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'fone2=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'fone2=' + '''' + temp + '''';
      end;
    end
    //FONE_FIRM
    else if (LowerCase(StringGrid1.Cells[i,0])='fone_firm') then
    begin
      temp := StringGrid1.Cells[i,k];
      temp := TextoSQL(Corta(temp,21,'fone_firm'),21); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',fone_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('fone_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'fone_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'fone_firm=' + '''' + temp + '''';
      end;
    end
    //FAX
    else if (LowerCase(StringGrid1.Cells[i,0])='fax') then
    begin
      temp := StringGrid1.Cells[i,k];
      temp := TextoSQL(Corta(temp,21,'fax'),21); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',fax';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('fax') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'fax=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'fax=' + '''' + temp + '''';
      end;
    end
    //FIRM (Firma ou Empresa que trabalha)
    else if (LowerCase(StringGrid1.Cells[i,0])='firm') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'firm');
      colClieForn := colClieForn + ',firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'firm=' + '''' + temp + '''';
      end;
    end
    //TRABALHA_DESDE (Trabalha na empresa desde quando)
    else if (LowerCase(StringGrid1.Cells[i,0])='trabalha_desde') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      //Data inválida não grava o campo
      if not DataValida(temp) then
      begin
        if temp <> '' then RegistraErro('trabalha_desde', StringGrid1.Cells[i,k], 'data inválida');
        temp := '';
      end;
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colClieForn := colClieForn + ',desd';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('trabalha_desde') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'desd=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'desd=' + '''' + temp + '''';
        end;
      end
      else if temp.Length = 6 then begin
        //Ano com 2 dígitos: maior que o ano atual é do século passado (em 2026: 85 = 1985, 10 = 2010)
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ AnoQuatroDigitos(Copy(temp,5,2));
        colClieForn := colClieForn + ',desd';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('trabalha_desde') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'desd=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'desd=' + '''' + temp + '''';
        end;
      end;
    end
    //ENDE_FIRM (Endereço da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='ende_firm') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,60,'ende_firm'),60); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',ende_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('ende_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'ende_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'ende_firm=' + '''' + temp + '''';
      end;
    end
    //CARG (Cargo da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='carg') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,30,'carg'),30); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',carg';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('carg') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'carg=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'carg=' + '''' + temp + '''';
      end;
    end
    //SALA (Sala da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='sala') then
    begin
      //Número (aceita 2.500,50 / 2500.50), se vazio ou inválido não grava o campo
      temp := NumeroSQL(StringGrid1.Cells[i,k]);
      if temp <> '' then
      begin
        colClieForn := colClieForn + ',sala';
        dadosClieForn := dadosClieForn + ',' + temp;
        //Testa se é Update
        if VerificaUpdate('sala') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'sala=' + temp;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'sala=' + temp;
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('sala', StringGrid1.Cells[i,k], 'valor numérico inválido');
    end
    //BAIR_FIRM (Bairro da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='bair_firm') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,30,'bair_firm'),30); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',bair_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('bair_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'bair_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'bair_firm=' + '''' + temp + '''';
      end;
    end
    //CIDA_FIRM (Cidade da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='cida_firm') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,30,'cida_firm'),30); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',cida_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('cida_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'cida_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'cida_firm=' + '''' + temp + '''';
      end;
    end
    //UF_FIRM (UF da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='uf_firm') then
    begin
      temp := AnsiUpperCase(Trim(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,2,'uf_firm'),2); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',uf_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('uf_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'uf_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'uf_firm=' + '''' + temp + '''';
      end;
    end
    //CEP_FIRM (CEP da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='cep_firm') then
    begin
      //Campo inteiro: remove a pontuação do CEP, se vazio ou inválido não grava o campo
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if not IsNumeric(temp) then temp := '';
      if temp <> '' then
      begin
        colClieForn := colClieForn + ',cep_firm';
        dadosClieForn := dadosClieForn + ',' + temp;
        //Testa se é Update
        if VerificaUpdate('cep_firm') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cep_firm=' + temp;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cep_firm=' + temp;
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('cep_firm', StringGrid1.Cells[i,k], 'valor numérico inválido');
    end
    //ESTA_CIVI (Estado Civil)
    else if (LowerCase(StringGrid1.Cells[i,0])='esta_civi') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := TextoSQL(Corta(temp,20,'esta_civi'),20); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',esta_civi';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('esta_civi') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'esta_civi=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'esta_civi=' + '''' + temp + '''';
      end;
    end
    //NOME_PAI
    else if (LowerCase(StringGrid1.Cells[i,0])='nome_pai') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'nome_pai');
      colClieForn := colClieForn + ',nome_pai';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('nome_pai') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'nome_pai=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'nome_pai=' + '''' + temp + '''';
      end;
    end
    //NOME_MAE
    else if (LowerCase(StringGrid1.Cells[i,0])='nome_mae') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'nome_mae');
      colClieForn := colClieForn + ',nome_mae';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('nome_mae') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'nome_mae=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'nome_mae=' + '''' + temp + '''';
      end;
    end
    //CONJ (Nome do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'conj');
      colClieForn := colClieForn + ',conj';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('conj') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'conj=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'conj=' + '''' + temp + '''';
      end;
    end
    //CONJ_FIRM (Trabalho do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj_firm') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'conj_firm');
      colClieForn := colClieForn + ',conj_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('conj_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'conj_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'conj_firm=' + '''' + temp + '''';
      end;
    end
    //CONJ_FIRM (Trabalho do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj_firm') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := Corta(temp,60,'conj_firm');
      colClieForn := colClieForn + ',conj_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('conj_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'conj_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'conj_firm=' + '''' + temp + '''';
      end;
    end
    //CONJ_SALA (Sala de Trabalho do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj_sala') then
    begin
      //Número (aceita 2.500,50 / 2500.50), se vazio ou inválido não grava o campo
      temp := NumeroSQL(StringGrid1.Cells[i,k]);
      if temp <> '' then
      begin
        colClieForn := colClieForn + ',conj_sala';
        dadosClieForn := dadosClieForn + ',' + temp;
        //Testa se é Update
        if VerificaUpdate('conj_sala') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'conj_sala=' + temp;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'conj_sala=' + temp;
        end;
      end
      else if Trim(StringGrid1.Cells[i,k]) <> '' then
        RegistraErro('conj_sala', StringGrid1.Cells[i,k], 'valor numérico inválido');
    end
    //CONJ_CARG (Cargo do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj_carg') then
    begin
      temp := AnsiUpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := TextoSQL(Corta(temp,30,'conj_carg'),30); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',conj_carg';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('conj_carg') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'conj_carg=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'conj_carg=' + '''' + temp + '''';
      end;
    end
    //DATA_CONJ (Data da união com conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='data_conj') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      //Data inválida não grava o campo
      if not DataValida(temp) then
      begin
        if temp <> '' then RegistraErro('data_conj', StringGrid1.Cells[i,k], 'data inválida');
        temp := '';
      end;
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colClieForn := colClieForn + ',data_conj';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('data_conj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'data_conj=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'data_conj=' + '''' + temp + '''';
        end;
      end
      else if temp.Length = 6 then begin
        //Ano com 2 dígitos: maior que o ano atual é do século passado (em 2026: 85 = 1985, 10 = 2010)
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ AnoQuatroDigitos(Copy(temp,5,2));
        colClieForn := colClieForn + ',data_conj';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('data_conj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'data_conj=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'data_conj=' + '''' + temp + '''';
        end;
      end;
    end
    //OBS
    else if (LowerCase(StringGrid1.Cells[i,0])='obs') then
    begin
      colClieForn := colClieForn + ',obs';
      temp := StringGrid1.Cells[i,k];
      temp := TextoSQL(temp,MaxInt); //Apóstrofo (campo texto sem limite de tamanho)
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('obs') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'obs=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'obs=' + '''' + temp + '''';
      end;
    end
    //REFE_COME (Referencia Comercial)
    else if (LowerCase(StringGrid1.Cells[i,0])='refe_come') then
    begin
      temp := StringGrid1.Cells[i,k];
      if Length(temp) > 330 then RegistraErro('refe_come', temp, 'texto maior que 330 caracteres');
      //Referência comercial não pode ser chave de Update (a importação também é bloqueada no início)
      if VerificaUpdate('refe_come') = 1 then
        RegistraErro('refe_come', '', 'a referência comercial não pode ser chave de Update');
      if (temp.Length>220) then
      begin
        colClieForn := colClieForn + ',refe_come3';
        temp2 := TextoSQL(Copy(temp,221,110),110);
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come3=' + '''' + temp2 + '''';
      end;
      if (temp.Length>110) then
      begin
        colClieForn := colClieForn + ',refe_come2';
        temp2 := TextoSQL(Copy(temp,111,110),110);
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come2=' + '''' + temp2 + '''';
      end;
      colClieForn := colClieForn + ',refe_come1';
      temp2 := TextoSQL(Copy(temp,1,110),110);
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come1=' + '''' + temp2 + '''';
        //Update com texto menor: limpar as partes que sobraram do texto anterior
        if (temp.Length<=110) then
          dadosUpdateClieForn := dadosUpdateClieForn + ', refe_come2=null';
        if (temp.Length<=220) then
          dadosUpdateClieForn := dadosUpdateClieForn + ', refe_come3=null';
      end;
    end
    //MAIL (EMAIL)
    else if (LowerCase(StringGrid1.Cells[i,0])='mail') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := TextoSQL(Corta(temp,250,'mail'),250); //Apóstrofo e tamanho do campo
      colClieForn := colClieForn + ',mail';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('mail') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'mail=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'mail=' + '''' + temp + '''';
      end;
    end
    //SEXO ('M' ou 'F')
    else if (LowerCase(StringGrid1.Cells[i,0])='sexo') then
    begin
      temp := UpperCase(Trim(StringGrid1.Cells[i,k]));
      if ((temp='MASCULINO') or (temp='MASC') or (temp='M')) then
      begin
        temp := 'M';
        colClieForn := colClieForn + ',sexo';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('sexo') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'sexo=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'sexo=' + '''' + temp + '''';
        end;
      end
      else if ((temp='FEMININO') or (temp='FEM') or (temp='F')) then
      begin
        temp := 'F';
        colClieForn := colClieForn + ',sexo';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('sexo') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'sexo=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'sexo=' + '''' + temp + '''';
        end;
      end
      else begin
        //Vazio não grava o campo (fica nulo)
        if temp <> '' then
          RegistraErro('sexo', StringGrid1.Cells[i,k], 'valor não reconhecido (use M ou F)');
      end;
    end
    //TIPOCAD (A=AMBOS, C=CLIENTE, F=FORNECEDOR)
    else if (LowerCase(StringGrid1.Cells[i,0])='tipocad') then
    begin
      temp := UpperCase(Trim(StringGrid1.Cells[i,k]));
      if ((temp='S') or (temp='1') or (temp='C')) then begin
        temp:='C';
      end
      else if ((temp='N') or (temp='2') or (temp='F')) then begin
        temp:='F';
      end
      else begin
        if (temp <> '') and (temp <> 'A') then
          RegistraErro('tipocad', StringGrid1.Cells[i,k], 'valor não reconhecido (use C, F ou A)');
        temp:='A';
      end;
      colClieForn := colClieForn + ',tipocad';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('tipocad') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'tipocad=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'tipocad=' + '''' + temp + '''';
      end;
    end
    //ATIVO
    else if (LowerCase(StringGrid1.Cells[i,0])='ativo') then
    begin
      temp := UpperCase(Trim(StringGrid1.Cells[i,k]));
      if ((temp='S') or (temp='1') or (temp='A')) then
      begin
        temp := 'S';
        colClieForn := colClieForn + ',ativo';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('ativo') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'ativo=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'ativo=' + '''' + temp + '''';
        end;
      end
      else if ((temp='N') or (temp='0') or (temp='I') or (temp='2')) then
      begin
        temp := 'N';
        colClieForn := colClieForn + ',ativo';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('ativo') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'ativo=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'ativo=' + '''' + temp + '''';
        end;
      end
      else if temp <> '' then
        RegistraErro('ativo', StringGrid1.Cells[i,k], 'valor não reconhecido (use S ou N)');
    end
    ;

  //Fim do for das colunas
  end;

  //CPF e CNPJ válidos na mesma linha (cada um define o TIPO F/J)
  temp := colClieForn + ',';
  count := 0;
  while Pos(',tipo,', temp) > 0 do
  begin
    count := count + 1;
    Delete(temp, Pos(',tipo,', temp), 5);
  end;
  if count > 1 then
    RegistraErro('cpf/cnpj', '', 'informe somente CPF ou somente CNPJ, não os dois');
end;

procedure cImportaClieForn.Gravar;
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
  if (colUpdateCount > 0) and (condUpdateClieForn = '') then
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
  if (colUpdateCount > 0) and (dadosUpdateClieForn = '') then
    Exit;

  //----------------------------------------
  //Gravar no banco de dados ClieForn
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        SQL := TSQLDataSet.Create(Nil);
        frmPrinc.conDestino.Open;
        SQL.SQLConnection := frmPrinc.conDestino;

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          frmImportando.atualizaStatus('Inserindo dados na tabela CLIEFORN.');

          //Desativar Trigger das cidades
          SQL.CommandText := 'ALTER TRIGGER clieforn_biu0 INACTIVE;';
          SQL.ExecSQL;
          try
            //Executar INSERT
            SQL.CommandText := 'INSERT INTO clieforn ('+ colClieForn +') VALUES ' + '(' + dadosClieForn + ');';
            SQL.ExecSQL;
          finally
            //Reativar Trigger das cidades, mesmo se o INSERT der erro
            //(ExecuteDirect para manter o INSERT no SQL.CommandText da mensagem de erro)
            frmPrinc.conDestino.ExecuteDirect('ALTER TRIGGER clieforn_biu0 ACTIVE;');
          end;
        end
        //Se for UPDATE
        else begin
          frmImportando.atualizaStatus('Atualizando dados na tabela CLIEFORN.');

          if dadosUpdateClieForn = '' then Exit;

          //Executar UPDATE
          SQL.CommandText := 'UPDATE clieforn SET '+ dadosUpdateClieForn +' WHERE ' + condUpdateClieForn + ';';
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
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Comandos da CLIEFORN.');
        WriteLn(fileTXT, '----------Comandos da CLIEFORN----------');

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          //Desativar Trigger das cidades
          WriteLn(fileTXT, 'ALTER TRIGGER clieforn_biu0 INACTIVE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
          //Executar INSERT
          WriteLn(fileTXT, 'INSERT INTO clieforn ('+ colClieForn +') VALUES ' + '(' + dadosClieForn + ');');
          WriteLn(fileTXT, 'COMMIT WORK;');
          //Reativar Trigger das cidades
          WriteLn(fileTXT, 'ALTER TRIGGER clieforn_biu0 ACTIVE;');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end
        //Se for UPDATE
        else begin
          //Executar UPDATE
          WriteLn(fileTXT, 'UPDATE clieforn SET '+ dadosUpdateClieForn +' WHERE ' + condUpdateClieForn + ';');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end;
      except
        on e: exception do
        begin
          ShowMessage('Erro TXT: '+e.message);
          status := 0;
          //Arquivo é fechado no finally
          Exit;
        end;
      end;
    finally
      CloseFile(fileTXT);
    end;
  end;
end;

end.
