unit uImportaSubGrupo;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaSubGrupo = class
    private
      colSubGrupo, dadosSubGrupo: String;

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

const
  TABELA = 'sub_grup_prod';
  TAM_DESCR = 30;

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
  bancoFDB: Boolean;
begin
  frmImportando.atualizaStatus('SubGrupo '+IntToStr(k));

  colSubGrupo := '';
  dadosSubGrupo := '';
  errosLinha.Clear;
  bancoFDB := UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB';

  //Carregar informações para importar
  //-------------------------------------------------------

  //CODI (CODIGO): sem a coluna usa o generator; com a coluna precisa ser um número inteiro
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    temp := Trim(stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll]));
    if temp = '' then
      RegistraErro('codi', '', 'código em branco')
    else if StrToIntDef(temp,-1) <= 0 then
      RegistraErro('codi', StringGrid1.Cells[i,k], 'código numérico inválido')
    else begin
      temp := IntToStr(StrToInt(temp));
      //Banco FDB: o código não pode já estar cadastrado
      if bancoFDB and (querySelect('SELECT codi FROM ' + TABELA + ' WHERE codi = ' + temp) <> '') then
        RegistraErro('codi', StringGrid1.Cells[i,k], 'código já cadastrado');
    end;
    colSubGrupo := colSubGrupo + 'codi';
    dadosSubGrupo := dadosSubGrupo + '''' + temp + '''';
  end
  else begin
    colSubGrupo := colSubGrupo + 'codi';
    dadosSubGrupo := dadosSubGrupo + 'gen_id(gen_sub_grup_prod_id,1)';
  end;

  //EMPR (EMPRESA): sem a coluna ou vazia usa a empresa 1
  i:=BuscaColuna(StringGrid1,'empr');
  temp := '1';
  if (i<>-1) and (Trim(StringGrid1.Cells[i,k]) <> '') then
  begin
    temp := Trim(StringGrid1.Cells[i,k]);
    if StrToIntDef(temp,-1) <= 0 then
      RegistraErro('empr', StringGrid1.Cells[i,k], 'valor numérico inválido')
    else begin
      temp := IntToStr(StrToInt(temp));
      if bancoFDB and (querySelect('SELECT codi FROM dados_empre WHERE codi = ' + temp) = '') then
        RegistraErro('empr', StringGrid1.Cells[i,k], 'empresa não encontrada');
    end;
  end;
  colSubGrupo := colSubGrupo + ',empr';
  dadosSubGrupo := dadosSubGrupo + ',''' + temp + '''';

  //DESCR (DESCRICAO): obrigatória, em maiúsculas com acento (padrão do ERP)
  //Apóstrofo vira espaço, igual ao cadastro pela importação de produtos
  i:=BuscaColuna(StringGrid1,'descr');
  temp := '';
  if i<>-1 then
    temp := Trim(stringreplace(AnsiUpperCase(StringGrid1.Cells[i,k]), '''', ' ',[rfReplaceAll]));
  if temp = '' then
    RegistraErro('descr', '', 'descrição em branco')
  else
    temp := Corta(temp, TAM_DESCR, 'descr');
  colSubGrupo := colSubGrupo + ',descr';
  dadosSubGrupo := dadosSubGrupo + ',''' + temp + '''';
end;

procedure cImportaSubGrupo.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Valores inválidos na linha: para a importação para corrigir e continuar desta linha
  if errosLinha.Count > 0 then
  begin
    ShowMessage('Linha ' + IntToStr(k) + ' com valores inválidos:' + sLineBreak + sLineBreak +
                errosLinha.Text + sLineBreak +
                'Corrija na planilha e clique em Importar para continuar a partir desta linha.');
    status := 0;
    Exit;
  end;

  //----------------------------------------
  //Gravar no banco de SUB_GRUP_PROD
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    SQL := TSQLDataSet.Create(Nil);
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
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
        frmImportando.atualizaStatus('Comandos da SUB_GRUP_PROD.');
        WriteLn(fileTXT, '----------Comandos da SUB_GRUP_PROD----------');

        WriteLn(fileTXT, 'INSERT INTO sub_grup_prod ('+ colSubGrupo +') VALUES ' + '(' + dadosSubGrupo + ');');
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
