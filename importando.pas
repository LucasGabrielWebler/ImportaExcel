unit importando;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons;


type
  TfrmImportando = class(TForm)
    BitBtn1: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure fim(status: integer);
    procedure atualizaItens(min: integer; max: integer);
    procedure atualizaStatus(status: string);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }
    erro: Boolean;
    linha: Integer;
  public
    { Public declarations }
  end;

var
  frmImportando: TfrmImportando;

implementation

uses importa_excel, uTema;

{$R *.dfm}

procedure TfrmImportando.atualizaItens(min: integer; max: integer);
begin
  linha := min;
  frmImportando.Label2.Caption := IntToStr(min) + ' / ' + IntToStr(max);
  frmImportando.Label2.Left := (Width - frmImportando.Label2.Width ) div 2;
  //frmImportando.Label2.Top := (Height - frmImportando.Label2.Height) div 2;
  Application.ProcessMessages; //Isso aqui faz a magica da atualizacao na tela acontecer
end;

procedure TfrmImportando.atualizaStatus(status: string);
begin
  frmImportando.Label3.Caption := status;
  frmImportando.Label3.Left := (Width - frmImportando.Label3.Width ) div 2;
  //frmImportando.Label3.Top := (Height - frmImportando.Label3.Height) div 2;
  Application.ProcessMessages; //Isso aqui faz a magica da atualizacao na tela acontecer
end;

procedure TfrmImportando.BitBtn1Click(Sender: TObject);
var
  i: Integer;
begin
  //Se deu erro, setar foco na linha que deu erro
  if erro then begin
    frmPrinc.StringGrid1.Row := linha;
    frmPrinc.StringGrid1.SetFocus;
  end;

  frmImportando.Close;
end;

procedure TfrmImportando.fim(status: integer);
begin
  if status = 1 then
  begin
    frmImportando.Label2.Caption := frmImportando.Label2.Caption + ' - Concluído!';
    frmImportando.Label2.Font.Color := CorSucesso;
    frmImportando.BitBtn1.Caption := 'Fechar';
  end
  else begin
    frmImportando.Label2.Caption := frmImportando.Label2.Caption + ' - Erro.';
    frmImportando.Label2.Font.Color := CorErro;
    frmImportando.BitBtn1.Caption := 'Ir para Erro';
    erro := True;
  end;
  frmImportando.Label2.Left := (Width - frmImportando.Label2.Width ) div 2;
  //frmImportando.Label2.Top := (Height - frmImportando.Label2.Height) div 2;
end;

//Título e status com as cores do tema (o tema não troca a cor das fontes próprias destes rótulos)
procedure TfrmImportando.FormShow(Sender: TObject);
begin
  Label1.Font.Color := CorAndamento;
  Label3.Font.Color := CorTexto;
end;

procedure TfrmImportando.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmImportando.Label2.Font.Color := CorAndamento;
  frmImportando.BitBtn1.Caption := 'Cancelar';
end;

end.
