unit uBackup;

{ Backup da base de dados (cópia do arquivo .FDB) feito pelo importador
  As cópias ficam em C:\BACKUP_IMPORTACAO\<razão social da empresa>\NN_DADOS_APOS_<TIPO>.FDB,
  numeradas na ordem em que foram feitas (uma cópia nunca sobrescreve outra).
  Antes de copiar confere se outro programa (ex.: ADMERP) está conectado na base, para a cópia
  não pegar uma gravação pela metade. }

interface

uses
  Winapi.Windows, Winapi.ShellAPI, System.SysUtils, System.Classes, System.UITypes, Vcl.Dialogs, Vcl.Forms,
  Vcl.Controls;

var
  //Pasta onde ficam os backups, uma pasta por empresa (os testes apontam para outra pasta)
  PastaBackup: string = 'C:\BACKUP_IMPORTACAO';

function NomeTipoBackup(tipo, tabela: string; atualizacao: Boolean): string;
function BackupAntesDaImportacao(tipo: string): Boolean;
procedure BackupDepoisDaImportacao(tipo: string);
procedure BackupManual;
procedure AbrePastaBackup;

implementation

uses
  importa_excel, uUtil;

const
  TITULO = 'Backup da base';

//Tipo de importação no nome da cópia (ex.: CLIENTES_FORNECEDORES, TITULOS_A_PAGAR, PRODUTOS_ATUALIZACAO)
function NomeTipoBackup(tipo, tabela: string; atualizacao: Boolean): string;
var
  i: Integer;
begin
  if tipo = 'Clie/Forn' then tipo := 'CLIENTES_FORNECEDORES'
  else if tipo = 'CUSTOM' then tipo := 'CUSTOM_' + tabela;
  Result := UpperCase(RemoveAcento(Trim(tipo)));
  for i := 1 to Length(Result) do
    if not CharInSet(Result[i], ['A'..'Z', '0'..'9']) then Result[i] := '_';
  while Pos('__', Result) > 0 do Result := StringReplace(Result, '__', '_', [rfReplaceAll]);
  while (Result <> '') and (Result[1] = '_') do Delete(Result, 1, 1);
  while (Result <> '') and (Result[Length(Result)] = '_') do Delete(Result, Length(Result), 1);
  if Result = '' then Result := 'IMPORTACAO';
  if atualizacao then Result := Result + '_ATUALIZACAO';
end;

//Arquivo da base no disco: o destino precisa ser um .FDB deste computador
//(aceita o servidor local no caminho, ex.: localhost/3050:C:\ADMERP\DADOS.FDB)
function ArquivoLocal(out arquivo, motivo: string): Boolean;
var
  caminho, host: string;
  p: Integer;
begin
  Result := False;
  motivo := '';
  caminho := Trim(frmPrinc.DBPath.Text);
  arquivo := caminho;
  if UpperCase(ExtractFileExt(caminho)) <> '.FDB' then
  begin
    motivo := 'Selecione uma base de dados (.FDB) como destino para fazer o backup.';
    Exit;
  end;
  p := Pos(':', caminho);
  if p > 2 then
  begin
    host := LowerCase(Copy(caminho, 1, p - 1));
    if Pos('/', host) > 0 then host := Copy(host, 1, Pos('/', host) - 1);
    if (host <> 'localhost') and (host <> '127.0.0.1') then
    begin
      motivo := 'A base está em outro computador (' + Copy(caminho, 1, p - 1) + ').' + sLineBreak +
                'Faça a cópia do arquivo no próprio servidor.';
      Exit;
    end;
    arquivo := Copy(caminho, p + 1, MaxInt);
  end;
  if not FileExists(arquivo) then
  begin
    motivo := 'Arquivo da base não encontrado:' + sLineBreak + arquivo;
    Exit;
  end;
  Result := True;
end;

//Nome da pasta da empresa: razão social da empresa 1, sem os caracteres que o Windows não aceita
function PastaEmpresa: string;
var
  nome: string;
  i: Integer;
begin
  nome := Trim(querySelect('SELECT FIRST 1 COALESCE(NULLIF(TRIM(razao), ''''), NULLIF(TRIM(nome_fantasia), '''')) ' +
                           'FROM dados_empre ORDER BY codi'));
  for i := 1 to Length(nome) do
    if (nome[i] < ' ') or CharInSet(nome[i], ['\', '/', ':', '*', '?', '"', '<', '>', '|']) then nome[i] := '-';
  while Pos('  ', nome) > 0 do nome := StringReplace(nome, '  ', ' ', [rfReplaceAll]);
  //Nome de pasta não pode terminar em ponto ou espaço
  while (nome <> '') and CharInSet(nome[Length(nome)], ['.', ' ']) do Delete(nome, Length(nome), 1);
  nome := Trim(nome);
  if nome = '' then nome := 'EMPRESA SEM NOME';
  Result := IncludeTrailingPathDelimiter(PastaBackup) + nome;
end;

//A pasta da empresa já tem alguma cópia da base
function TemCopias(pasta: string): Boolean;
var
  sr: TSearchRec;
begin
  Result := FindFirst(IncludeTrailingPathDelimiter(pasta) + '*.FDB', faAnyFile, sr) = 0;
  if Result then FindClose(sr);
end;

//Próximo número da sequência: maior número das cópias da pasta + 1
function ProximoNumero(pasta: string): Integer;
var
  sr: TSearchRec;
  n: Integer;
begin
  Result := 1;
  if FindFirst(IncludeTrailingPathDelimiter(pasta) + '*.FDB', faAnyFile, sr) = 0 then
  try
    repeat
      n := StrToIntDef(Copy(sr.Name, 1, Pos('_', sr.Name) - 1), -1);
      if n + 1 > Result then Result := n + 1;
    until FindNext(sr) <> 0;
  finally
    FindClose(sr);
  end;
end;

function NomeCopia(pasta: string; numero: Integer; descricao: string): string;
begin
  Result := IncludeTrailingPathDelimiter(pasta) + Format('%.2d_DADOS_%s.FDB', [numero, descricao]);
end;

//Programas conectados na base, além do importador (ex.: ADMERP.exe)
function ConexoesAbertas: string;
var
  lista: TStringList;
  i: Integer;
begin
  lista := TStringList.Create;
  try
    lista.Delimiter := '|';
    lista.StrictDelimiter := True;
    lista.DelimitedText := querySelect(
      'SELECT CAST(LIST(DISTINCT COALESCE(NULLIF(TRIM(mon$remote_process), ''''), NULLIF(TRIM(mon$remote_address), ''''), ' +
      '''programa desconhecido''), ''|'') AS VARCHAR(4000)) FROM mon$attachments ' +
      'WHERE mon$attachment_id <> CURRENT_CONNECTION AND mon$system_flag = 0');
    Result := '';
    for i := 0 to lista.Count - 1 do
      if Trim(lista[i]) <> '' then
      begin
        if Result <> '' then Result := Result + ', ';
        Result := Result + ExtractFileName(Trim(lista[i]));
      end;
  finally
    lista.Free;
  end;
end;

function TamanhoArquivo(arquivo: string): Int64;
var
  sr: TSearchRec;
begin
  Result := -1;
  if FindFirst(arquivo, faAnyFile, sr) = 0 then
  begin
    Result := sr.Size;
    FindClose(sr);
  end;
end;

//Copia a base para o arquivo de backup; com outro programa conectado, pede para fechar antes
function CopiaBase(origem, destino: string): Boolean;
var
  conexoes, erro: string;
begin
  Result := False;
  repeat
    conexoes := ConexoesAbertas;
    if conexoes = '' then Break;
    case Mensagem('A base está aberta em outro programa: ' + conexoes + '.' + sLineBreak + sLineBreak +
                  'Feche o ADMERP (e os outros programas conectados nesta base) e clique em Tentar novamente,' + sLineBreak +
                  'para a cópia não pegar uma gravação pela metade.' + sLineBreak + sLineBreak +
                  'Base: ' + origem,
                  mtWarning, [mbYes, mbNo, mbCancel], ['Tentar novamente', 'Copiar assim mesmo', 'Cancelar'], TITULO) of
      mrYes: Continue;
      mrNo: Break;
    else
      Exit;
    end;
  until False;

  erro := '';
  if not ForceDirectories(ExtractFileDir(destino)) then
    erro := 'Não foi possível criar a pasta ' + ExtractFileDir(destino) + '.'
  else begin
    Screen.Cursor := crHourGlass;
    try
      if not CopyFile(PChar(origem), PChar(destino), True) then
        erro := SysErrorMessage(GetLastError)
      else if TamanhoArquivo(destino) <> TamanhoArquivo(origem) then
      begin
        DeleteFile(destino);
        erro := 'A cópia ficou com tamanho diferente da base.';
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  end;

  if erro <> '' then
  begin
    Mensagem('Não foi possível copiar a base.' + sLineBreak + sLineBreak + erro + sLineBreak + sLineBreak +
             'Base: ' + origem + sLineBreak + 'Cópia: ' + destino, mtError, [mbOK], ['OK'], TITULO);
    Exit;
  end;
  Mensagem('Cópia da base gravada em:' + sLineBreak + sLineBreak + destino, mtInformation, [mbOK], ['OK'], TITULO);
  Result := True;
end;

//Antes da primeira importação da empresa (pasta de backup ainda sem cópias): pergunta se guarda a base como está.
//Retorna False se a cópia foi pedida e não foi feita (a importação não começa).
function BackupAntesDaImportacao(tipo: string): Boolean;
var
  origem, motivo, pasta, destino: string;
begin
  Result := True;
  if not ArquivoLocal(origem, motivo) then Exit;
  pasta := PastaEmpresa;
  if TemCopias(pasta) then Exit;
  destino := NomeCopia(pasta, 0, 'ANTES_' + tipo);
  if Mensagem('Ainda não há backup desta empresa em ' + PastaBackup + '.' + sLineBreak + sLineBreak +
              'Deseja guardar uma cópia da base ANTES desta importação?' + sLineBreak + sLineBreak +
              'Base: ' + origem + sLineBreak + 'Cópia: ' + destino,
              mtConfirmation, [mbYes, mbNo], ['Sim', 'Não'], TITULO) <> mrYes then Exit;
  if not CopiaBase(origem, destino) then
  begin
    Mensagem('A importação não foi iniciada, pois a cópia da base não foi feita.', mtWarning, [mbOK], ['OK'], TITULO);
    Result := False;
  end;
end;

//Depois da importação concluída: pergunta se faz o backup
procedure BackupDepoisDaImportacao(tipo: string);
var
  origem, motivo, pasta, destino: string;
begin
  if UpperCase(ExtractFileExt(frmPrinc.DBPath.Text)) <> '.FDB' then Exit;
  if not ArquivoLocal(origem, motivo) then
  begin
    Mensagem('Não é possível fazer o backup desta base.' + sLineBreak + sLineBreak + motivo, mtWarning, [mbOK], ['OK'], TITULO);
    Exit;
  end;
  pasta := PastaEmpresa;
  destino := NomeCopia(pasta, ProximoNumero(pasta), 'APOS_' + tipo);
  if Mensagem('Deseja realizar backup (cópia da base de dados)?' + sLineBreak + sLineBreak +
              'Base: ' + origem + sLineBreak + 'Cópia: ' + destino,
              mtConfirmation, [mbYes, mbNo], ['Sim', 'Não'], TITULO) <> mrYes then Exit;
  CopiaBase(origem, destino);
end;

//Menu Backup: cópia da base na hora
procedure BackupManual;
var
  origem, motivo, pasta: string;
begin
  if not ArquivoLocal(origem, motivo) then
  begin
    Mensagem(motivo, mtWarning, [mbOK], ['OK'], TITULO);
    Exit;
  end;
  pasta := PastaEmpresa;
  CopiaBase(origem, NomeCopia(pasta, ProximoNumero(pasta), 'MANUAL'));
end;

//Menu Backup: abre a pasta dos backups da empresa da base de destino (ou a pasta geral)
procedure AbrePastaBackup;
var
  origem, motivo, pasta, empresa: string;
begin
  pasta := PastaBackup;
  if ArquivoLocal(origem, motivo) then
  begin
    empresa := PastaEmpresa;
    if DirectoryExists(empresa) then pasta := empresa;
  end;
  if not DirectoryExists(pasta) then
  begin
    Mensagem('Nenhum backup foi feito ainda (a pasta ' + pasta + ' não existe).', mtInformation, [mbOK], ['OK'], TITULO);
    Exit;
  end;
  ShellExecute(Application.Handle, 'open', PChar(pasta), nil, nil, SW_SHOWNORMAL);
end;

end.
