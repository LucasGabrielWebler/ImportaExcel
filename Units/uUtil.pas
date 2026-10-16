unit uUtil;

interface

uses
  System.SysUtils, System.Generics.Collections, Vcl.Grids, System.Classes, Vcl.Forms,
  Vcl.Controls, Vcl.Dialogs, Vcl.StdCtrls, Vcl.OleAuto, System.Variants, Data.SqlExpr;

  //Utilidades gerais
  procedure String2File(str: string; path: string = '');
  function corrigeFloat(nume:string): string;
  function RemoveAcento(Str: string): string;
  function Mensagem(CONST Msg: string; DlgTypt: TmsgDlgType; button: TMsgDlgButtons;
    Caption: ARRAY OF string; dlgcaption: string): Integer;
  function queryInsert(sql: string): Integer;
  function querySelect(sql: String): String;
  function IsNumeric(S : String) : Boolean;
  function PrimeiraMaiuscula(texto: string) : string;
  function TextoSQL(texto: string; tamanho: Integer): string;
  function NumeroSQL(texto: string): string;
  function DataValida(data: string): Boolean;
  function AnoQuatroDigitos(ano: string): string;

  //Validação dos valores da linha (a importação para na linha com erro)
  procedure RegistraErro(coluna, valor, motivo: string);
  function Corta(texto: string; tamanho: Integer; coluna: string): string;

  //Funções para GRID
  procedure DeleteRow(Grid: TStringGrid; ARow: Integer);
  procedure DeleteCol(Grid: TStringGrid; ACol: Integer);
  procedure InsertCol(Grid: TStringGrid);
  procedure InsertRow(Grid: TStringGrid);
  procedure RemoveWhiteRows(Grid: TStringGrid);
  procedure RemoveSpaces(Grid: TStringGrid);
  procedure AutoSizeCol(Grid: TStringGrid; Column: integer);
  function BuscaColuna(Grid: TStringGrid; colName: String) : Integer;
  function checkCol(grid: TStringGrid) : Boolean;
  procedure List_To_Grid(Grid: TStringGrid; tabela: TStringList);
  function SaveAsCSVFile(Grid: TstringGrid; FileName: string):Boolean;
  procedure CSV_To_StringGrid(StringGrid1: TStringGrid; AFileName: TFileName);
  function XlsHeaderLoad(AGrid: TStringGrid; AXLSFile: string): Boolean;
  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
  function VerificaUpdate(coluna: String): Integer;

var
  //Erros de valores da linha que está sendo importada
  errosLinha: TStringList;

implementation

uses
  importa_excel;

//----------------------
//Utilidade gerais
//----------------------
//Salvar variavel string em TXT
procedure String2File(str: string; path: string = '');
var
  fileTXT: TextFile;
begin
  path := path +'LOG_'+stringreplace(DateToStr(now), '/', '-',[rfReplaceAll, rfIgnoreCase])+'.TXT';
  try
    AssignFile(fileTXT, path);
    ReWrite(fileTXT);
    Write(fileTXT,str);
  finally
    CloseFile(fileTXT);
  end;
end;

//Função para tratar os numeros com ponto flutuante importados como texto
function corrigeFloat(nume:string): string;
var
  c : Char;
  flag : Integer;
begin
  {
  Flag para encontrar pontos e virgulas
  1- 1º Ponto
  2- 1ª Virgula
  3- Ponto é casa decimal
  4- Virgula é casa decimal
  }
  flag := 0;

  //Encontrar pontos e virgulas
  for c in nume do begin
    if flag = 0 then begin
      if c = '.'  then flag := 1;
      if c = ',' then flag := 2;
    end
    else begin
      if c = '.'  then flag := 3;
      if c = ',' then flag := 4;
    end;
  end;

  {
  Verificar a flag:
  1-Decimal separado por ponto, deixa assim
  2-Decimal separado por virgula, troca virgula por ponto
  3-Milhares separ. por virgulas, Decimal por ponto, tira virgulas
  4-Milhares separ. por pontos, Decimal por virgula, tira os pontos depois troca virgula por ponto
  }
  if flag = 2 then begin
    nume := stringreplace(nume, ',', '.',[rfReplaceAll, rfIgnoreCase]);
  end
  else if flag = 4 then begin
    nume := stringreplace(nume, '.', '',[rfReplaceAll, rfIgnoreCase]);
    nume := stringreplace(nume, ',', '.',[rfReplaceAll, rfIgnoreCase]);
  end
  else if flag = 3 then begin
    nume := stringreplace(nume, ',', '',[rfReplaceAll, rfIgnoreCase]);
  end;
  Result := nume;
end;

//FUNÇÃO DO JEFINHO PARA REMOVER ACENTOS
function RemoveAcento(Str: string): string;
const
  ComAcento = 'àâêôûãõáéíóúçüÀÂÊÔÛÃÕÁÉÍÓÚÇÜ';
  SemAcento = 'aaeouaoaeioucuAAEOUAOAEIOUCU';
var
  x: Integer;
begin;
  for x := 1 to Length(Str) do
    if Pos(Str[x], ComAcento) <> 0 then
      Str[x] := SemAcento[Pos(Str[x], ComAcento)];
  Result := Str;
end;

//Função para criar caixa de diálogos
function Mensagem(CONST Msg: string; DlgTypt: TmsgDlgType; button: TMsgDlgButtons;
  Caption: ARRAY OF string; dlgcaption: string): Integer;
var
  aMsgdlg: TForm;
  i: Integer;
  Dlgbutton: Tbutton;
  Captionindex: Integer;
begin
  aMsgdlg := createMessageDialog(Msg, DlgTypt, button);
  aMsgdlg.Caption := dlgcaption;
  aMsgdlg.BiDiMode := bdLeftToRight;
  Captionindex := 0;
  for i := 0 to aMsgdlg.componentcount - 1 Do
  begin
    if (aMsgdlg.components[i] is Tbutton) then
    Begin
      Dlgbutton := Tbutton(aMsgdlg.components[i]);
      if Captionindex <= High(Caption) then
        Dlgbutton.Caption := Caption[Captionindex];
      inc(Captionindex);
    end;
  end;
  Result := aMsgdlg.Showmodal;
end;

//FUNÇÃO PARA INSERTS SQL
function queryInsert(sql: string): Integer;
var
  gen_id: Integer;
  queryTemp: TSQLQuery;
  fileTXT: TextFile;
begin
  try
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.CommandText := sql;
      queryTemp.ExecSQL;

    except
      on e: exception do
      begin
        ShowMessage('Erro queryInsert SQL: '+e.message+sLineBreak+queryTemp.CommandText+'\nContinuando sem inserir.');
      end;
    end;
  finally
    if queryTemp.IsEmpty then
    begin
      Result := -1;
    end
    else begin
      Result := queryTemp.FieldByName('CODI').AsInteger;
    end;
    queryTemp.Close;
    frmPrinc.conDestino.Close;
  end;
end;

//FUNÇÃO PARA CONSULTAS SQL QUE RETORNAM 1 RESULTADO
function querySelect(sql: String): String;
var
  queryTemp: TSQLQuery;

begin
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then Result := ''

  else
    try
      try
        frmPrinc.conDestino.Open;
        queryTemp := TSQLQuery.Create(nil);
        queryTemp.SQLConnection := frmPrinc.conDestino;
        queryTemp.SQL.Clear;
        queryTemp.CommandText := sql;
        queryTemp.ExecSQL;
        queryTemp.Open;

        Result := queryTemp.Fields[0].AsString;

      except
        on e: exception do
        begin
          ShowMessage('Erro SQL: '+e.message+sLineBreak+queryTemp.CommandText);
        end;
      end;
    finally
      queryTemp.Free;
      frmPrinc.conDestino.Close;
    end;
end;

//Função que tenta transformar a string em numero e retorna TRUE se conseguir
function IsNumeric(S : String) : Boolean;
begin
  Result := True;
  Try
     StrToInt(S);
  Except
    Result := False;
  end;
end;

function PrimeiraMaiuscula(texto: string) : string;
begin
  Result := UpperCase(Copy(texto,1,1)) + LowerCase(Copy(texto,2,Length(texto)-1));
end;


//Função para preparar texto para o SQL: corta no tamanho do campo e duplica os apóstrofos
function TextoSQL(texto: string; tamanho: Integer): string;
begin
  Result := stringreplace(Copy(texto,1,tamanho), '''', '''''',[rfReplaceAll]);
end;


//Função para converter número da planilha para o SQL (2.500,50 / 2500.50 / 2500)
//Retorna vazio se não for um número válido
function NumeroSQL(texto: string): string;
var
  valor: Double;
  formato: TFormatSettings;
begin
  Result := '';
  texto := Trim(texto);
  if texto = '' then Exit;
  texto := corrigeFloat(texto);
  formato := TFormatSettings.Create;
  formato.DecimalSeparator := '.';
  formato.ThousandSeparator := ',';
  if TryStrToFloat(texto, valor, formato) then
    Result := texto;
end;


//Função para validar data da planilha já sem separadores (ddmmaaaa ou ddmmaa)
function DataValida(data: string): Boolean;
var
  d: TDateTime;
  c: Char;
begin
  Result := False;
  if Length(data) >= 8 then data := Copy(data,1,8)
  else if Length(data) <> 6 then Exit;
  for c in data do
    if not CharInSet(c, ['0'..'9']) then Exit;
  if Length(data) = 8 then
    Result := TryEncodeDate(StrToInt(Copy(data,5,4)), StrToInt(Copy(data,3,2)), StrToInt(Copy(data,1,2)), d)
  else
    Result := TryEncodeDate(StrToInt(AnoQuatroDigitos(Copy(data,5,2))), StrToInt(Copy(data,3,2)), StrToInt(Copy(data,1,2)), d);
end;


//Função para completar o ano de 2 dígitos (aa) com o século
//Ano maior que o ano atual é do século passado (ex.: em 2026, 85 = 1985 e 10 = 2010)
function AnoQuatroDigitos(ano: string): string;
begin
  if StrToInt(ano) > StrToInt(FormatDateTime('yy', Date)) then
    Result := '19' + ano
  else
    Result := '20' + ano;
end;


//Registra um valor inválido da linha que está sendo importada
procedure RegistraErro(coluna, valor, motivo: string);
begin
  //Mostra no máximo 60 caracteres do valor na mensagem
  if Length(valor) > 60 then valor := Copy(valor,1,60) + '...';
  errosLinha.Add(UpperCase(coluna) + ': ' + motivo + ' [' + valor + ']');
end;


//Corta o texto no tamanho do campo, registrando erro se for maior
function Corta(texto: string; tamanho: Integer; coluna: string): string;
begin
  Result := Copy(texto,1,tamanho);
  if Length(texto) > tamanho then
    RegistraErro(coluna, texto, 'texto maior que ' + IntToStr(tamanho) + ' caracteres');
end;



//----------------------
//Funções para GRID
//----------------------
//Deletar Linha da StringGrid
procedure DeleteRow(Grid: TStringGrid; ARow: Integer);
var
  i: Integer;
begin
  for i := ARow to Grid.RowCount - 2 do
    Grid.Rows[i].Assign(Grid.Rows[i + 1]);
  Grid.RowCount := Grid.RowCount - 1;
end;

//Deletar coluna na StringGrid
procedure DeleteCol(Grid: TStringGrid; ACol: Integer);
var
  i: Integer;
begin
  for i := ACol to Grid.ColCount - 2 do
    Grid.Cols[i].Assign(Grid.Cols[i + 1]);
  Grid.ColCount := Grid.ColCount - 1;
end;

//Inserir coluna na StringGrid
procedure InsertCol(Grid: TStringGrid);
var
  i,j: Integer;
  temp: string;
begin
  Grid.ColCount := Grid.ColCount + 1;
  i:= Grid.ColCount;
  while i>Grid.Col do
  begin
    for j := 0 to Grid.RowCount do
    begin
      temp := Grid.Cells[i,j];
      Grid.Cells[i,j] := Grid.Cells[i-1,j];
    end;
    i:= i-1;
  end;
  for j := 0 to Grid.RowCount do
      Grid.Cells[i,j] := '';
end;

//Inserir linha na StringGrid
procedure InsertRow(Grid: TStringGrid);
var
  i,j: Integer;
  temp: string;
begin
  Grid.RowCount := Grid.RowCount + 1;
  i:= Grid.RowCount;
  while i>Grid.Row do
  begin
    for j := 0 to Grid.ColCount do
    begin
      temp := Grid.Cells[j,i];
      Grid.Cells[j,i] := Grid.Cells[j,i-1];
    end;
    i:= i-1;
  end;
  for j := 0 to Grid.ColCount do
      Grid.Cells[j,i] := '';
end;

//Função para Remover linhas em branco
procedure RemoveWhiteRows(Grid: TStringGrid);
var
  i, j: integer;
  remove: Boolean;
begin
  //Percorre linhas
  for i := 0 to (Grid.RowCount - 1) do begin
    remove := True;
    //Percorre colunas
    for j := 1 to (Grid.ColCount - 1) do begin
      if Grid.Cells[j,i] <> '' then begin
        remove := False;
        Break;
      end;
    end;

    if remove = True then begin
      DeleteRow(Grid, i);
    end;
  end;
end;

//Função para Remover espaços no inicio e fim da string
procedure RemoveSpaces(Grid: TStringGrid);
var
  i, j: integer;
begin
  //Percorre linhas
  for i := 0 to (Grid.RowCount - 1) do begin
    //Percorre colunas
    for j := 1 to (Grid.ColCount - 1) do begin
      Grid.Cells[j,i] := TrimLeft(Grid.Cells[j,i]);
      Grid.Cells[j,i] := TrimRight(Grid.Cells[j,i]);
      Grid.Cells[j,i] := stringreplace(Grid.Cells[j,i], ';', '',[rfReplaceAll, rfIgnoreCase]);
    end;
  end;
end;

//Função para redimensionar coluna
procedure AutoSizeCol(Grid: TStringGrid; Column: integer);
var
  i, W, WMax: integer;
begin
  WMax := 0;
  for i := 0 to (Grid.RowCount - 1) do begin
    W := Grid.Canvas.TextWidth(Grid.Cells[Column, i]);
    if W > WMax then
      WMax := W;
  end;
  Grid.ColWidths[Column] := WMax + 10;

  //Se for uma coluna nova
  if Grid.ColWidths[Column] = 10 then
  begin
    Grid.ColWidths[Column] := 40;
  end;
end;

//Função para buscar a coluna no StringGrid e retornar o indice
function BuscaColuna(Grid: TStringGrid; colName: String) : Integer;
var
  i: integer;
begin
  colName := UpperCase(colName);

  for i := 0 to Grid.ColCount-1 do
  begin
    if UpperCase(Grid.Cells[i,0]) = colName then
      Break;
  end;
  if i = Grid.ColCount then
  begin
    Result:=-1;
  end
  else begin
    Result:=i;
  end;
end;

//Função para testar colunas com mesmo nome
function checkCol(grid: TStringGrid) : Boolean;
var
  i,j: Integer;
  temp: string;
begin
  Result := True;
  for i := 1 to grid.ColCount-1 do
  begin
    temp := grid.Cells[i,0];
    if temp='' then Continue;

    for j := 1 to grid.ColCount-1 do
    begin
      if i=j then Continue;
      if grid.Cells[j,0]=temp then
      begin
        ShowMessage('Colunas com mesmo nome ('+temp+'): '+IntToStr(i)+' e '+IntToStr(j));
        Result := False;
        Exit;
      end;
    end;
  end;
end;

//Função para carregar StringList na StringGrid
procedure List_To_Grid(Grid: TStringGrid; tabela: TStringList);
var
  oRowStrings:TStringList;
  i:integer;
begin
  oRowStrings := TStringList.Create;
  try
    Grid.RowCount := tabela.Count;
    for i := 0 to tabela.Count - 1 do
    begin
      oRowStrings.Clear;
      oRowStrings.Delimiter := ';';
      oRowStrings.StrictDelimiter := True;
      oRowStrings.DelimitedText := tabela[i];
      oRowStrings.Insert(0,IntToStr(i));
      if oRowStrings.Count > Grid.ColCount then
        Grid.ColCount := oRowStrings.Count;
      Grid.Rows[i].Assign(oRowStrings);
    end;
  finally
    oRowStrings.Free;
  end;
end;

//Função para salvar StringGrid em um arquivo CSV
function SaveAsCSVFile(Grid: TstringGrid; FileName: string):Boolean;
var
  i, j : Integer;
  CSV : TStrings;
  stream : string;
begin
  //Criar StringList
  CSV := TStringList.Create;

  Screen.Cursor := crHourGlass;
  CSV.Delimiter := ';';

  Try
    for i := 0 to Grid.RowCount - 1 do
    begin
      if Grid.ColCount >= 2 then
      begin
        stream := Grid.Cells[1,i];
      end
      else begin
        Exit;
      end;

      for j := 2 to Grid.ColCount -1 do
      begin
        stream := stream + ';' + Grid.Cells[j,i];
      end;
      CSV.Add(stream);
    end;

    try
      //Salvar no CSV
      CSV.SaveToFile(FileName);
      Result := True;
    except
      Result := False;
    end;
  Finally
    Screen.Cursor := crDefault;
    CSV.Free;
  End;
end;

//Função para carregar CSV na StringGrid
procedure CSV_To_StringGrid(StringGrid1: TStringGrid; AFileName: TFileName);
var
  oFileStrings:TStringList;
  oRowStrings:TStringList;
  i:integer;
begin
  oFileStrings := TStringList.Create;
  oRowStrings := TStringList.Create;
  try
    oFileStrings.LoadFromFile(AFileName);
    StringGrid1.RowCount := oFileStrings.Count;
    for i := 0 to oFileStrings.Count - 1 do
    begin
      oRowStrings.Clear;
      oRowStrings.Delimiter := ';';
      oRowStrings.StrictDelimiter := True;
      oRowStrings.DelimitedText := oFileStrings[i];
      oRowStrings.Insert(0,IntToStr(i));
      if oRowStrings.Count > StringGrid1.ColCount then
        StringGrid1.ColCount := oRowStrings.Count;
      StringGrid1.Rows[i].Assign(oRowStrings);
    end;
  finally
    oFileStrings.Free;
    oRowStrings.Free;
  end;
end;

//Função para carregar apenas o cabeçalho de uma planilha na StringGrid
function XlsHeaderLoad(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);

    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it

    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;


    //Verificar quantidade de colunas
    if AGrid.ColCount < y then AGrid.ColCount := y;
    //Testar se a primeira coluna não é um '0'
    if RangeMatrix[1, 1] <> '0' then
    begin
      //AGrid.ColCount := AGrid.ColCount +1;
      AGrid.Cells[0,0] := '0';
    end;

    //Iterar colunas
    for r := 1 to y do
    begin
      AGrid.Cells[r,0] := RangeMatrix[1, r];
    end;

    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;


  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      // XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

//Função para carregar planilha na StringGrid
function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  i, j, x, y, k, r: Integer;
  but: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    try

      // Hide Excel
      XLApp.Visible := False;

      // Open the Workbook
      XLApp.Workbooks.Open(AXLSFile);

      but := 0;
      AGrid.RowCount := 0;
      AGrid.ColCount := 1;
      k := 0;
      j := 1;
      //Percorrer os WorkSheets
      for i := 1 to XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Count do
      begin
        Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[i];

        //Retirar Filtros, para que não fique linhas escondidas
        if (Sheet.AutoFilterMode = True) then
        begin
          Sheet.AutoFilterMode := False;
        end;

        //Receber valores da ultima linha e coluna
        x := Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Row;
        y := Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Column;

        //Setar tamanho do StringGrid
        AGrid.RowCount := AGrid.RowCount + x;
        if y > AGrid.ColCount  then
        begin
          AGrid.ColCount := y + 1;
        end;

        // Assign the Variant associated with the WorkSheet to the Delphi Variant
        RangeMatrix := Sheet.Range['A1', Sheet.Cells.Item[X, Y]].Value;

        //Iterar linhas
        while k < AGrid.RowCount-1 do
        begin
          if but = 1 then Break;

          AGrid.Cells[0,k] := IntToStr(k);
          //Iterar colunas
          for r := 1 to AGrid.ColCount-1 do
          begin
            try
              AGrid.Cells[r,k] := RangeMatrix[j,r];
            except
              on E:Exception do
              begin
                if but = 7 then Continue;

                but := Mensagem('Erro no arquivo Excel linha: '+IntToStr(j)+' coluna: '+IntToStr(r)+' ('+RangeMatrix[1,r]+')'+#13+E.Message+#13+'Continuar irá deixar célula em branco. Ignorar irá deixar célula em branco para todos os erros sem perguntar.', mtCustom,[mbYes, mbNo, mbOK],['Continuar', 'Ignorar','Parar'],'Erro no arquivo Excel');
                if (but = 6) then begin
                  AGrid.Cells[r,k] := '';
                end
                else if (but = 7) then begin
                  Continue;
                end
                else if (but = 1) then begin
                  raise Exception.Create('Erro no arquivo Excel. Verificar se existem células contendo:'+#13+'#DIV/0!'+#13+'#N/A'+#13+'#NAME?'+#13+'#NULL!'+#13+'#NUM!'+#13+'#REF!'+#13+'#VALUE!');
                end;
              end;
            end;
          end;
          k := k + 1;
          j := j + 1;
        end;

        //Fazer proximos WorksSheet começar da 2 linha
        j := 2;
        AGrid.RowCount := AGrid.RowCount - 1;
      end;

    except
      on E:Exception do
      begin
        ShowMessage(E.Message);
      end;
    end;

  finally
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

//Verificar se coluna é Update ou não
function VerificaUpdate(coluna: String): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to colUpdateCount-1 do begin
    if LowerCase(coluna) = LowerCase(colUpdate[i]) then begin
      Result := 1;
      Break;
    end;
  end;
end;

initialization
  errosLinha := TStringList.Create;

finalization
  errosLinha.Free;

end.
