unit uImportaMarca;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaMarca = class
    private
      colMarca, dadosMarca: String;

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
  TABELA = 'marca';
  TAM_DESCR = 50;

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
  bancoFDB: Boolean;
begin
  frmImportando.atualizaStatus('Marca '+IntToStr(k));

  colMarca := '';
  dadosMarca := '';
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
    colMarca := colMarca + 'codi';
    dadosMarca := dadosMarca + '''' + temp + '''';
  end
  else begin
    colMarca := colMarca + 'codi';
    dadosMarca := dadosMarca + 'gen_id(gen_marca_id,1)';
  end;

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
  colMarca := colMarca + ',descr';
  dadosMarca := dadosMarca + ',''' + temp + '''';
end;

procedure cImportaMarca.Gravar;
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
  //Gravar no banco de MARCA
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    SQL := TSQLDataSet.Create(Nil);
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
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
        frmImportando.atualizaStatus('Comandos da MARCA.');
        WriteLn(fileTXT, '----------Comandos da MARCA----------');

        WriteLn(fileTXT, 'INSERT INTO marca ('+ colMarca +') VALUES ' + '(' + dadosMarca + ');');
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
