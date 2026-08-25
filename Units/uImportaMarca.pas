unit uImportaMarca;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaMarca = class
    private
      colMarca, dadosMarca, condUpdateMarca, dadosUpdateMarca: String;

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

constructor cImportaMarca.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaMarca.CarregaColunas;
var
  i: Integer;
  temp: string;
begin
  frmImportando.atualizaStatus('Marca '+IntToStr(k));

  colMarca := '';
  dadosMarca := '';

  //Carregar informações para importar
  //-------------------------------------------------------

  //Codigo é obrigatório, se não tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    colMarca := colMarca + 'codi';
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    dadosMarca := dadosMarca + '''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colMarca := colMarca + 'codi';
    dadosMarca := dadosMarca + 'gen_id(gen_marca_id,1)';
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //DESCR (DESCRICAO)
    if (LowerCase(StringGrid1.Cells[i,0])='descr') then
    begin
      colMarca := colMarca + ',descr';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
      dadosMarca := dadosMarca + ',''' + temp + '''';
    end
    ;

  //Fim do For das colunas
  end;
end;

procedure cImportaMarca.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //----------------------------------------
  //Gravar no banco de MARCA
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela MARCA.');
        SQL.CommandText := 'INSERT INTO marca ('+ colMarca +') VALUES ' + '(' + dadosMarca + ');';
        SQL.ExecSQL;

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

        frmImportando.atualizaStatus('Comandos da MARCA.');
        WriteLn(fileTXT, '----------Comandos da MARCA----------');

        WriteLn(fileTXT, 'INSERT INTO marca ('+ colMarca +') VALUES ' + '(' + dadosMarca + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');
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
