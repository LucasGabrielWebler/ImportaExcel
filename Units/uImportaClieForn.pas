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

  frmImportando.atualizaStatus('Clie/Forn '+ IntToStr(k));

  //Codigo é obrigatório, se não tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    if (StringGrid1.Cells[i,k]='') then
    begin
      ShowMessage('Código em branco na linha '+IntToStr(k));
    end
    else begin
      StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
      colClieForn := colClieForn + 'codi';
      dadosClieForn := dadosClieForn + '''' + StringGrid1.Cells[i,k] + '''';
      //Testa se é Update
      if VerificaUpdate('codi') = 1 then
        condUpdateClieForn := condUpdateClieForn + 'codi=' + '''' + StringGrid1.Cells[i,k] + ''''
      else dadosUpdateClieForn := dadosUpdateClieForn + 'codi=' + '''' + StringGrid1.Cells[i,k] + '''';
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
    temp2 := stringreplace(temp2, '''', QuotedStr(''''),[rfReplaceAll, rfIgnoreCase]);
    temp := frmPrinc.buscaCidade(temp2, temp);
    if (temp <> '') then
    begin
      colClieForn := colClieForn + ',cida,codi_cida';
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('cida') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'cida=' + '''' + temp2 + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'cida=' + '''' + temp2 + ''', codi_cida=' + temp;
      end;
    end;
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //EMPRESA
    if ((LowerCase(StringGrid1.Cells[i,0])='empresa') and (StringGrid1.Cells[i,0]<>'')) then
    begin
      colClieForn := colClieForn + ',empresa';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('empresa') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'empresa=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'empresa=' + '''' + temp + '''';
      end;
    end
    //GRUPO
    else if ((LowerCase(StringGrid1.Cells[i,0])='grupo') and (StringGrid1.Cells[i,0]<>'')) then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,50));
      //Se for letras, buscar código.
      if not (IsNumeric(temp)) then
      begin
        temp2 := querySelect('SELECT g.codi FROM grupo_cliente g WHERE g.descr = '''+temp+'''');
        //Se não encontrar a string, cadastrar sub grupo
        if temp2='' then begin
          if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
            //Carregar arquivo TXT
            AssignFile(fileTXT, frmPrinc.DBPath.Text);
            if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
            else append(fileTXT);
            WriteLn(fileTXT, 'INSERT INTO grupo_cliente (CODI,DESCR,COMISSAO) VALUES (CASE WHEN (SELECT gc.codi FROM grupo_cliente gc WHERE gc.descr='''+temp+''') IS NULL THEN gen_id(gen_grupo_cliente_id,1 ) ELSE (SELECT gc.codi FROM grupo_cliente gc WHERE gc.descr='''+temp+''') END,'''+temp+''',1);');
            WriteLn(fileTXT, 'COMMIT WORK;');
            CloseFile(fileTXT);
          end
          else begin
            queryInsert('INSERT INTO grupo_cliente (CODI,DESCR,COMISSAO) VALUES (gen_id(gen_grupo_cliente_id,1),'''+temp+''',1);');
          end;
          colClieForn := colClieForn + ',codi_grupo_clie';
          dadosClieForn := dadosClieForn + ',' + 'gen_id(gen_grupo_cliente_id,0)';
        end
        else begin
          colClieForn := colClieForn + ',codi_grupo_clie';
          dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
          //Testa se é Update
          if VerificaUpdate('grupo') = 1 then begin
            ShowMessage('Só é possível utilizar o código do grupo como Update, não da pra usar a descrição.');
            Exit;
          end
          else begin
            if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
            dadosUpdateClieForn := dadosUpdateClieForn + 'grupo=' + '''' + temp + '''';
          end;
        end;

      end
      else begin
        //Se for números, considera como código
        //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
        temp2 := querySelect('SELECT g.codi FROM grupo_cliente g WHERE g.codi = '''+temp+'''');
        //Se não encontrar o codigo, colocamos o generator
        if temp2='' then begin
          colClieForn := colClieForn + ',codi_grupo_clie';
          dadosClieForn := dadosClieForn + ',' + 'gen_id(gen_grupo_cliente_id,0)';
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', QuotedStr(''''),[rfReplaceAll, rfIgnoreCase]);
      temp2 := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', QuotedStr(''''),[rfReplaceAll, rfIgnoreCase]);
      temp2 := (Copy(temp,1,60));
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
      end;
    end
    //DATA_NASC
    else if (LowerCase(StringGrid1.Cells[i,0])='data_nasc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
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
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com nascimento do cliente
        //Se os caracteres da data de nascimento do cliente forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
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
      if temp.Length = 11 then
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
      end;
    end
    //CNPJ
    else if (LowerCase(StringGrid1.Cells[i,0])='cnpj') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length = 14 then
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
      end;
    end
    //CPF OU CNPJ NO MESMO CAMPO
    else if (LowerCase(StringGrid1.Cells[i,0])='cpf_cnpj') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length = 11 then
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
      else if temp.Length = 14 then
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
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + '''' + 'F' + '''';
        end;
      end
      else
      begin
        //CPF NULL
        colClieForn := colClieForn + ',cpf';
        dadosClieForn := dadosClieForn + ',' + 'null';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cpf IS NULL';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cpf=' + 'null';
        end;

        //CNPJ null
        //Trazer em branco
        colClieForn := colClieForn + ',cnpj';
        dadosClieForn := dadosClieForn + ',' + 'null';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'cnpj if null';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'cnpj=' + 'null';
        end;

        //TIPO null
        //Trazer em branco
        //Setar tipo (Fisca ou Juridica)
        colClieForn := colClieForn + ',tipo';
        dadosClieForn := dadosClieForn + ',' + 'null';
        //Testa se é Update
        if VerificaUpdate('cpf_cnpj') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'tipo IS NULL';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + ',tipo=' + 'null';
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
      temp := (Copy(temp,1,16));
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
        dadosClieForn := dadosClieForn + ',''' + 'ISENTO' + '''';
      end;
    end
    //INSC (INSCRICAO ESTADUAL-IE)
    else if (LowerCase(StringGrid1.Cells[i,0])='insc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,20));
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
        dadosClieForn := dadosClieForn + ',''' + 'ISENTO' + '''';
      end;
    end
    //INSCR_PRODUTOR (INSCRICAO DE PRODUTOR RURAL)
    else if (LowerCase(StringGrid1.Cells[i,0])='inscr_produtor') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,20));
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
        dadosClieForn := dadosClieForn + ',''' + 'ISENTO' + '''';
      end;
    end
    //ENDE (ENDERECO)
    else if (LowerCase(StringGrid1.Cells[i,0])='ende') then
    begin
      colClieForn := colClieForn + ',ende';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
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
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colClieForn := colClieForn + ',desd';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('desd') = 1 then begin
          if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
          condUpdateClieForn := condUpdateClieForn + 'desd=' + '''' + temp + '''';
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'desd=' + '''' + temp + '''';
        end;
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com a data desc
        //Se os caracteres da data de nascimento do cliente forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        colClieForn := colClieForn + ',desd';
        dadosClieForn := dadosClieForn + ',''' + temp + '''';
        //Testa se é Update
        if VerificaUpdate('desd') = 1 then begin
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
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := Trim(StringGrid1.Cells[i,k]);
      colClieForn := colClieForn + ',sala';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('sala') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'sala=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'sala=' + '''' + temp + '''';
      end;
    end
    //BAIR_FIRM (Bairro da empresa)
    else if (LowerCase(StringGrid1.Cells[i,0])='bair_firm') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := Trim(StringGrid1.Cells[i,k]);
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
      temp := Trim(StringGrid1.Cells[i,k]);
      colClieForn := colClieForn + ',cep_firm';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('cep_firm') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'cep_firm=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'cep_firm=' + '''' + temp + '''';
      end;
    end
    //ESTA_CIVI (Estado Civil)
    else if (LowerCase(StringGrid1.Cells[i,0])='esta_civi') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
      colClieForn := colClieForn + ',conj_sala';
      dadosClieForn := dadosClieForn + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('conj_sala') = 1 then begin
        if condUpdateClieForn <> '' then condUpdateClieForn := condUpdateClieForn + ' AND ';
        condUpdateClieForn := condUpdateClieForn + 'conj_sala=' + '''' + temp + '''';
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'conj_sala=' + '''' + temp + '''';
      end;
    end
    //CONJ_CARG (Cargo do Conjuge)
    else if (LowerCase(StringGrid1.Cells[i,0])='conj_carg') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,60));
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
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com a data desc
        //Se os caracteres da data de nascimento do cliente forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
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
      temp := (Copy(temp,1,80));
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
      if (temp.Length>220) then
      begin
        colClieForn := colClieForn + ',refe_come3';
        temp2 := (Copy(temp,221,110));
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Testa se é Update
        if VerificaUpdate('refe_come') = 1 then begin
          ShowMessage('Referencia comercial não pode ser uma condição de Update. Sry mas não da.');
          Exit;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come3=' + '''' + temp2 + '''';
        end;
      end;
      if (temp.Length>110) then
      begin
        colClieForn := colClieForn + ',refe_come2';
        temp2 := (Copy(temp,111,110));
        dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
        //Testa se é Update
        if VerificaUpdate('refe_come') = 1 then begin
          ShowMessage('Referencia comercial não pode ser uma condição de Update. Sry mas não da.');
          Exit;
        end
        else begin
          if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
          dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come2=' + '''' + temp2 + '''';
        end;
      end;
      colClieForn := colClieForn + ',refe_come1';
      temp2 := (Copy(temp,1,110));
      dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
      //Testa se é Update
      if VerificaUpdate('refe_come') = 1 then begin
        ShowMessage('Referencia comercial não pode ser uma condição de Update. Sry mas não da.');
        Exit;
      end
      else begin
        if dadosUpdateClieForn <> '' then dadosUpdateClieForn := dadosUpdateClieForn + ', ';
        dadosUpdateClieForn := dadosUpdateClieForn + 'refe_come1=' + '''' + temp2 + '''';
      end;
    end
    //MAIL (EMAIL)
    else if (LowerCase(StringGrid1.Cells[i,0])='mail') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
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
        colClieForn := colClieForn + ',sexo';
        dadosClieForn := dadosClieForn + ',''' + 'F' + '''';
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
      end;
    end
    ;

  //Fim do for das colunas
  end;
end;

procedure cImportaClieForn.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //----------------------------------------
  //Gravar no banco de dados ClieForn
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          frmImportando.atualizaStatus('Inserindo dados na tabela CLIEFORN.');

          //Desativar Trigger das cidades
          SQL.CommandText := 'ALTER TRIGGER clieforn_biu0 INACTIVE;';
          SQL.ExecSQL;
          //Executar INSERT
          SQL.CommandText := 'INSERT INTO clieforn ('+ colClieForn +') VALUES ' + '(' + dadosClieForn + ');';
          SQL.ExecSQL;
          //Reativar Trigger das cidades
          SQL.CommandText := 'ALTER TRIGGER clieforn_biu0 ACTIVE;';
          SQL.ExecSQL;
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
          SQL.Free;
          frmPrinc.conDestino.Close;
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
          CloseFile(fileTXT);
          Exit;
        end;
      end;
    finally
      CloseFile(fileTXT);
    end;
  end;
end;

end.
