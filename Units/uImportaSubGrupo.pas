unit uImportaSubGrupo;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaSubGrupo = class
    private
      colSubGrupo, dadosSubGrupo, condUpdateSubGrupo, dadosUpdateSubGrupo: String;

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

constructor cImportaSubGrupo.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaSubGrupo.CarregaColunas;
var
  i: Integer;
  temp: string;
begin
  frmImportando.atualizaStatus('SubGrupo '+IntToStr(k));

  colSubGrupo := '';
  dadosSubGrupo := '';

  //Carregar informações para importar
  //-------------------------------------------------------

  //Codigo é obrigatório, se não tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    colSubGrupo := colSubGrupo + 'codi';
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    dadosSubGrupo := dadosSubGrupo + '''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colSubGrupo := colSubGrupo + 'codi';
    dadosSubGrupo := dadosSubGrupo + 'gen_id(gen_sub_grup_prod_id,1)';
  end;

  //Empresa é obrigatório, se não tiver preenche com 1
  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    colSubGrupo := colSubGrupo + ',empr';
    dadosSubGrupo := dadosSubGrupo + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colSubGrupo := colSubGrupo + ',empr';
    dadosSubGrupo := dadosSubGrupo + ',''' + '1' + '''';
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //DESCR (DESCRICAO)
    if (LowerCase(StringGrid1.Cells[i,0])='descr') then
    begin
      colSubGrupo := colSubGrupo + ',descr';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
      dadosSubGrupo := dadosSubGrupo + ',''' + temp + '''';
    end
    ;

  //Fim do For das colunas
  end;
end;

procedure cImportaSubGrupo.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //----------------------------------------
  //Gravar no banco de SUB_GRUP_PROD
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela SUB_GRUP_PROD.');
        SQL.CommandText := 'INSERT INTO sub_grup_prod ('+ colSubGrupo +') VALUES ' + '(' + dadosSubGrupo + ');';
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

        frmImportando.atualizaStatus('Comandos da SUB_GRUP_PROD.');
        WriteLn(fileTXT, '----------Comandos da SUB_GRUP_PROD----------');

        WriteLn(fileTXT, 'INSERT INTO sub_grup_prod ('+ colSubGrupo +') VALUES ' + '(' + dadosSubGrupo + ');');
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
