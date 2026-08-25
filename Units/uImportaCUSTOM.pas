unit uImportaCUSTOM;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaCUSTOM = class
    private
      col, dados, condUpdate, dadosUpdate: String;

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

constructor cImportaCUSTOM.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaCUSTOM.CarregaColunas;
var
  i: Integer;
begin
  col := '';
  dados := '';
  condUpdate := '';
  dadosUpdate := '';

  frmImportando.atualizaStatus('CUSTOM '+ IntToStr(k));

  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //Ignorar colunas sem cabeçalho
    if StringGrid1.Cells[i,0] = '' then
    begin
      Continue;
    end;

    //Ignorar coluna de número da linha
    if StringGrid1.Cells[i,0] = '0' then
    begin
      Continue;
    end;

    if col = '' then
    begin
      col := StringGrid1.Cells[i,0];
      //Se registro em branco, colocar null, sen coloca Quoted
      if StringGrid1.Cells[i,k] = '' then
      begin
        dados := 'null';
      end
      else
      begin
        dados := QuotedStr(StringGrid1.Cells[i,k]);
      end;
    end
    else
    begin
      col := col + ',' + StringGrid1.Cells[i,0];
      //Se registro em branco, colocar null, sen coloca Quoted
      if StringGrid1.Cells[i,k] = '' then
      begin
        dados := dados + ',' + 'null';
      end
      else
      begin
        dados := dados + ',' + QuotedStr(StringGrid1.Cells[i,k]);
      end;
    end;

    //Testa se é Update
    if VerificaUpdate(StringGrid1.Cells[i,0]) = 1 then begin
      if condUpdate <> '' then condUpdate := condUpdate + ' AND ';
      condUpdate := condUpdate + StringGrid1.Cells[i,0] + '=' + QuotedStr(StringGrid1.Cells[i,k]);
    end
    else begin
      if dadosUpdate <> '' then dadosUpdate := dadosUpdate + ', ';
      dadosUpdate := dadosUpdate + StringGrid1.Cells[i,0] + '=' + QuotedStr(StringGrid1.Cells[i,k]);
    end;
  end;
end;

procedure cImportaCUSTOM.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //----------------------------------------
  //Gravar no banco de dados / arquivo
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          frmImportando.atualizaStatus('Inserindo dados na tabela '+frmPrinc.edtTableName.Text+'.');

          //Executar INSERT
          SQL.CommandText := 'INSERT INTO '+frmPrinc.edtTableName.Text+' ('+ col +') VALUES ' + '(' + dados + ');';
          SQL.ExecSQL;
        end
        //Se for UPDATE
        else begin
          frmImportando.atualizaStatus('Atualizando dados na tabela '+frmPrinc.edtTableName.Text+'.');

          if dadosUpdate = '' then Exit;

          //Executar UPDATE
          SQL.CommandText := 'UPDATE '+frmPrinc.edtTableName.Text+' SET '+ dadosUpdate +' WHERE ' + condUpdate + ';';
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

        frmImportando.atualizaStatus('Comandos da '+frmPrinc.edtTableName.Text+'.');
        WriteLn(fileTXT, '----------Comandos da '+frmPrinc.edtTableName.Text+'----------');

        //Se for INSERT
        if colUpdateCount <= 0 then begin
          //Executar INSERT
          WriteLn(fileTXT, 'INSERT INTO '+frmPrinc.edtTableName.Text+' ('+ col +') VALUES ' + '(' + dados + ');');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end
        //Se for UPDATE
        else begin
          //Executar UPDATE
          WriteLn(fileTXT, 'UPDATE '+frmPrinc.edtTableName.Text+' SET '+ dadosUpdate +' WHERE ' + condUpdate + ';');
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
