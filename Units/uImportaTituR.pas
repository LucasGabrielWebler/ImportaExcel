unit uImportaTituR;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaTituR = class
    private
      colTituR, dadosTituR, condUpdateTituR, dadosUpdateTituR: string;
      colBTitu, dadosBTitu, condUpdateBTitu, dadosUpdateBTitu: string;
      saldo: Double;

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

constructor cImportaTituR.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaTituR.CarregaColunas;
var
  i, l, j, count: Integer;
  temp, temp2: string;
begin
  frmImportando.atualizaStatus('Títulos a Receber '+IntToStr(k));

  colTituR := '';
  dadosTituR := '';
  colBTitu := '';
  dadosBTitu := '';

  //Carregar informações para importar
  //-------------------------------------------------------

  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    StringGrid1.Cells[i,k] := (Copy(StringGrid1.Cells[i,k],1,15));
    l := Length(StringGrid1.Cells[i,k]);
    //Testar se ja existir o código do título e inserir uma barra.
    count := 0;
    while (frmPrinc.temCodTituloR(StringGrid1.Cells[i,k]) = True) do
    begin
      for j := 1 to count do
      begin
        StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '/'+IntToStr(j), '',[rfReplaceAll, rfIgnoreCase]);
        StringGrid1.Cells[i,k] := (Copy(StringGrid1.Cells[i,k],1,l));
      end;

      count := count+1;
      StringGrid1.Cells[i,k] := StringGrid1.Cells[i,k] + '/' + IntToStr(count);
    end;

    colTituR := colTituR + 'codi';
    dadosTituR := dadosTituR + '''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + 'codi';
    dadosBTitu := dadosBTitu + '''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + 'codi';
    dadosTituR := dadosTituR + IntToStr(k);
    colBTitu := colBTitu + 'codi';
    dadosBTitu := dadosBTitu + IntToStr(k);
  end;

  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    colTituR := colTituR + ',empr';
    dadosTituR := dadosTituR + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',empr';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + ',empr';
    dadosTituR := dadosTituR + ',1';
    colBTitu := colBTitu + ',empr';
    dadosBTitu := dadosBTitu + ',1';
  end;

  //CLIE (CLIENTE)
  i:=BuscaColuna(StringGrid1,'clie');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,60));
    //Se for letras, buscar código.
    if not (IsNumeric(temp)) then
    begin
      temp2 := IntToStr(frmPrinc.getCodiClieForn(''''+temp+''''));
      //Se não encontrar a string, cadastrar cliente
      if temp2='0' then begin
        temp2 := IntToStr(frmPrinc.cadastraClieForn('nome',''''+temp+''''));
      end;
      colTituR := colTituR + ',clie';
      dadosTituR := dadosTituR + ',''' + temp2 + '''';
      colBTitu := colBTitu + ',clie';
      dadosBTitu := dadosBTitu + ',''' + temp2 + '''';
    end
    else begin
      //Se for números, considera como código
      temp2 := IntToStr(frmPrinc.getCodiClieForn('(SELECT c.nome FROM clieforn c WHERE c.codi = '+temp+')'));
      //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
      if temp2='0' then begin
        colTituR := colTituR + ',clie';
        dadosTituR := dadosTituR + ',' + 'gen_id(gen_clieforn_id,0)';
        colBTitu := colBTitu + ',clie';
        dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_clieforn_id,0)';
      end
      else begin
        //Se achar o código, usamos o código
        colTituR := colTituR + ',clie';
        dadosTituR := dadosTituR + ',''' + temp + '''';
        colBTitu := colBTitu + ',clie';
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end;
    end;
  end
  else begin
    //Se não tiver fornecedor, colocar o generator.
    colTituR := colTituR + ',clie';
    dadosTituR := dadosTituR + ',' + 'gen_id(gen_clieforn_id,0)';
    colBTitu := colBTitu + ',clie';
    dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_clieforn_id,0)';
  end;

  //LOCA_COBR (LOCAL DE COBRANÇA)
  i:=BuscaColuna(StringGrid1,'loca_cobr');
  if (i<>-1) then
  begin
    colTituR := colTituR + ',loca_cobr';
    dadosTituR := dadosTituR + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + ',loca_cobr';
    dadosTituR := dadosTituR + ',''' + '1' + '''';
  end;

  //CART (TIPO DE CARTEIRA)
  i:=BuscaColuna(StringGrid1,'cart');
  if (i<>-1) then
  begin
    colTituR := colTituR + ',cart';
    dadosTituR := dadosTituR + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + ',cart';
    dadosTituR := dadosTituR + ',''' + '1' + '''';
  end;

  //OPER (OPERAÇÃO DO PLANO DE CONTAS)
  i:=BuscaColuna(StringGrid1,'oper');
  if (i<>-1) then
  begin
    colTituR := colTituR + ',oper';
    dadosTituR := dadosTituR + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',oper';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + ',oper';
    dadosTituR := dadosTituR + ',''' + '002' + '''';
    colBTitu := colBTitu + ',oper';
    dadosBTitu := dadosBTitu + ',''' + '002' + '''';
  end;

  //C_FUNC (FUNCIONÁRIO)
  i:=BuscaColuna(StringGrid1,'c_func');
  if (i<>-1) then
  begin
    colTituR := colTituR + ',c_func';
    dadosTituR := dadosTituR + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',c_func';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituR := colTituR + ',c_func';
    dadosTituR := dadosTituR + ',''' + '1' + '''';
    colBTitu := colBTitu + ',c_func';
    dadosBTitu := dadosBTitu + ',''' + '1' + '''';
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //DATA (Data de criação do título)
    if (LowerCase(StringGrid1.Cells[i,0])='data') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colTituR := colTituR + ',data';
        dadosTituR := dadosTituR + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        colTituR := colTituR + ',data';
        dadosTituR := dadosTituR + ',''' + temp + '''';
      end;
    end
    //VENC (Data de vencimento do título)
    else if (LowerCase(StringGrid1.Cells[i,0])='venc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colTituR := colTituR + ',venc';
        dadosTituR := dadosTituR + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        colTituR := colTituR + ',venc';
        dadosTituR := dadosTituR + ',''' + temp + '''';
      end;
    end
    //VALO (Valor do título)
    else if (LowerCase(StringGrid1.Cells[i,0])='valo') then
    begin
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0.0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      colTituR := colTituR + ',valo';
      dadosTituR := dadosTituR + ',' + temp;
      colBTitu := colBTitu + ',valo';
      dadosBTitu := dadosBTitu + ',' + temp;
      colBTitu := colBTitu + ',tota';
      dadosBTitu := dadosBTitu + ',' + temp;

      //Testar se não existe coluna saldo, se não existir joga o valor da VALO
      count:=BuscaColuna(StringGrid1,'sald');
      if (count=-1) then
      begin
        colTituR := colTituR + ',sald';
        if StringGrid1.Cells[i,k]='' then
        begin
          temp2 := '0.0';
        end
        else begin
          temp2 := StringGrid1.Cells[i,k];
        end;
        temp2 := corrigeFloat(temp2);
        dadosTituR := dadosTituR + ',' + temp2;
        temp2 := stringreplace(temp2, '.', ',',[rfReplaceAll, rfIgnoreCase]);
        saldo := StrToFloat(temp2);
      end;
    end
    //SALDO (Saldo do título)
    else if (LowerCase(StringGrid1.Cells[i,0])='sald') then
    begin
      colTituR := colTituR + ',sald';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0.0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosTituR := dadosTituR + ',' + temp;
      temp := stringreplace(temp, '.', ',',[rfReplaceAll, rfIgnoreCase]);
      saldo := StrToFloat(temp);

      //Testar se não existe coluna valo, se não existir joga o valor da SALD
      count:=BuscaColuna(StringGrid1,'valo');
      if (count=-1) then
      begin
        if StringGrid1.Cells[i,k]='' then
        begin
          temp2 := '0.0';
        end
        else begin
          temp2 := StringGrid1.Cells[i,k];
        end;
        temp2 := corrigeFloat(temp2);
        colTituR := colTituR + ',valo';
        dadosTituR := dadosTituR + ',' + temp2;
        colBTitu := colBTitu + ',valo';
        dadosBTitu := dadosBTitu + ',' + temp2;
        colBTitu := colBTitu + ',tota';
        dadosBTitu := dadosBTitu + ',' + temp2;
      end;
    end
    //HIST (HISTORICO)
    else if (LowerCase(StringGrid1.Cells[i,0])='hist') then
    begin
      colTituR := colTituR + ',hist';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,150));
      dadosTituR := dadosTituR + ',''' + temp + '''';
    end

    //Colunas extras para BTITUR
    //DATA_BAIXA (DATA DA BAIXA)
    else if (LowerCase(StringGrid1.Cells[i,0])='data_baixa') then
    begin
      colBTitu := colBTitu + ',data';
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que é um século antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end;
    end
    ;

  //Fim do For das colunas
  end;

  //ID (ID DA BTITUR)
  colBTitu := colBTitu + ',id';
  dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_btitur_id,1)';
  //JURO (JUROS)
  colBTitu := colBTitu + ',juro';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //MULT (MULTA)
  colBTitu := colBTitu + ',mult';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //DESCO (DESCONTO)
  colBTitu := colBTitu + ',desco';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //CONT (CONTA)
  colBTitu := colBTitu + ',cont';
  dadosBTitu := dadosBTitu + ',''' + '1' + '''';
  //EMPR_BAIX (EMPRESA ONDE O TÍTULO FOI BAIXADO)
  colBTitu := colBTitu + ',empr_baix';
  dadosBTitu := dadosBTitu + ',''' + '1' + '''';
end;

procedure cImportaTituR.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //----------------------------------------
  //Gravar no banco de Títulos a Receber
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela TITUR.');
        SQL.CommandText := 'INSERT INTO titur ('+ colTituR +') VALUES ' + '(' + dadosTituR + ');';
        SQL.ExecSQL;

        if saldo <= 0.0 then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Inserindo dados na tabela BTITUR.');
          SQL.CommandText := 'INSERT INTO btitur ('+ colBTitu +') VALUES ' + '(' + dadosBTitu +');';
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
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Comandos da TITUR.');
        WriteLn(fileTXT, '----------Comandos da TITUR----------');

        WriteLn(fileTXT, 'INSERT INTO titur ('+ colTituR +') VALUES ' + '(' + dadosTituR + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');

        if saldo <= 0.0 then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Comandos BTITUR.');
          WriteLn(fileTXT, '----------Comandos da BTITUR----------');
          WriteLn(fileTXT, 'INSERT INTO btitur ('+ colBTitu +') VALUES ' + '(' + dadosBTitu +');');
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
