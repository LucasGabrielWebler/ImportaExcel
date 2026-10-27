unit uTema;

{ Tema da aplicação (claro do Windows ou um dos temas VCL embutidos no executável)
  A escolha fica salva por usuário em %APPDATA%\ImportaExcel\ImportaExcel.ini e volta ao abrir a aplicação.
  As partes desenhadas no código (faixas da grade e da tela Colunas, cores de status) pegam as cores daqui. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.IniFiles, System.UITypes, Vcl.Graphics, Vcl.Menus, Vcl.Themes, Vcl.Styles;

const
  //Tema padrão do Windows (sem estilo)
  TEMA_CLARO = 'Windows';

procedure IniciaTema;
procedure AplicaTema(nome: string);
procedure MontaMenuTema(menu: TMenuItem; aoClicar: TNotifyEvent);
function TemaEscuro: Boolean;

//Cores das partes desenhadas no código
function CorFaixa(linha: Integer): TColor;
function CorTexto: TColor;
function CorCabecalho: TColor;
function CorUpdate: TColor;
function CorDescricao: TColor;
function CorSucesso: TColor;
function CorErro: TColor;
function CorAndamento: TColor;

implementation

function ArquivoConfig: string;
begin
  Result := IncludeTrailingPathDelimiter(GetEnvironmentVariable('APPDATA')) + 'ImportaExcel';
  ForceDirectories(Result);
  Result := Result + '\ImportaExcel.ini';
end;

//Aplica o tema salvo (antes de criar as telas); tema que não existe mais volta para o claro
procedure IniciaTema;
var
  ini: TIniFile;
  nome: string;
begin
  nome := TEMA_CLARO;
  try
    ini := TIniFile.Create(ArquivoConfig);
    try
      nome := ini.ReadString('Aparencia', 'Tema', TEMA_CLARO);
    finally
      ini.Free;
    end;
  except
    //Sem acesso à pasta do usuário: segue com o tema claro
  end;
  if not TStyleManager.TrySetStyle(nome, False) then
    TStyleManager.TrySetStyle(TEMA_CLARO, False);
end;

//Troca o tema na hora (todas as telas abertas) e salva a escolha
procedure AplicaTema(nome: string);
var
  ini: TIniFile;
begin
  if not TStyleManager.TrySetStyle(nome, False) then Exit;
  try
    ini := TIniFile.Create(ArquivoConfig);
    try
      ini.WriteString('Aparencia', 'Tema', nome);
    finally
      ini.Free;
    end;
  except
    //Sem acesso à pasta do usuário: o tema vale só até fechar a aplicação
  end;
end;

//Itens do menu Tema: o claro do Windows e os temas embutidos, marcando o atual
procedure MontaMenuTema(menu: TMenuItem; aoClicar: TNotifyEvent);
var
  nomes: TStringList;
  item: TMenuItem;
  nome: string;
begin
  menu.Clear;
  nomes := TStringList.Create;
  try
    for nome in TStyleManager.StyleNames do
      if nome <> TEMA_CLARO then nomes.Add(nome);
    nomes.Sort;
    nomes.Insert(0, TEMA_CLARO);
    for nome in nomes do
    begin
      item := TMenuItem.Create(menu);
      if nome = TEMA_CLARO then item.Caption := 'Claro (padrão do Windows)'
      else item.Caption := nome;
      item.Hint := nome;
      item.RadioItem := True;
      item.GroupIndex := 1;
      item.Checked := nome = TStyleManager.ActiveStyle.Name;
      item.OnClick := aoClicar;
      menu.Add(item);
    end;
  finally
    nomes.Free;
  end;
end;

//Mistura duas cores (porcentagem da segunda)
function Mistura(c1, c2: TColor; pct: Integer): TColor;
var
  r1, g1, b1, r2, g2, b2: Integer;
begin
  c1 := ColorToRGB(c1);
  c2 := ColorToRGB(c2);
  r1 := GetRValue(c1); g1 := GetGValue(c1); b1 := GetBValue(c1);
  r2 := GetRValue(c2); g2 := GetGValue(c2); b2 := GetBValue(c2);
  Result := RGB(r1 + (r2 - r1) * pct div 100, g1 + (g2 - g1) * pct div 100, b1 + (b2 - b1) * pct div 100);
end;

function Cor(c: TColor): TColor;
begin
  Result := ColorToRGB(StyleServices.GetSystemColor(c));
end;

//Tema escuro: fundo das janelas com pouca luminosidade
function TemaEscuro: Boolean;
var
  c: TColor;
begin
  c := Cor(clWindow);
  Result := (GetRValue(c) * 299 + GetGValue(c) * 587 + GetBValue(c) * 114) div 1000 < 128;
end;

//Faixas da grade: azul e amarelo claros no tema claro; no escuro, o fundo do tema levemente puxado para azul e amarelo
function CorFaixa(linha: Integer): TColor;
begin
  if not TemaEscuro then
  begin
    if linha mod 2 = 0 then Result := $00E1FFF9 else Result := $00FFEBDF;
  end
  else begin
    if linha mod 2 = 0 then Result := Mistura(Cor(clWindow), $0030B0C0, 14)
    else Result := Mistura(Cor(clWindow), $00C06030, 22);
  end;
end;

function CorTexto: TColor;
begin
  if TemaEscuro then Result := Cor(clWindowText) else Result := clBlack;
end;

//Linha de títulos e coluna de números da grade
function CorCabecalho: TColor;
begin
  if TemaEscuro then Result := Cor(clBtnFace) else Result := clWhite;
end;

//Coluna marcada como Update
function CorUpdate: TColor;
begin
  if TemaEscuro then Result := $00FFB060 else Result := clBlue;
end;

//Descrição em cinza na tela Colunas
function CorDescricao: TColor;
begin
  if TemaEscuro then Result := Mistura(Cor(clWindowText), Cor(clWindow), 40) else Result := clGrayText;
end;

function CorSucesso: TColor;
begin
  if TemaEscuro then Result := $0060D060 else Result := clGreen;
end;

function CorErro: TColor;
begin
  if TemaEscuro then Result := $006060FF else Result := clRed;
end;

function CorAndamento: TColor;
begin
  if TemaEscuro then Result := $00FFB060 else Result := clHotLight;
end;

end.
