unit uSubstituir;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmSubstituir = class(TForm)
    rgSelecao: TRadioGroup;
    Label1: TLabel;
    edtOldValue: TEdit;
    Label2: TLabel;
    edtNewValue: TEdit;
    btnReplace: TButton;
    gbExtras: TGroupBox;
    btnRemoveAcentos: TButton;
    btnMinuscula: TButton;
    btnMaiusculas: TButton;
    btnPrimeiraMaiuscula: TButton;
    procedure btnReplaceClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnRemoveAcentosClick(Sender: TObject);
    procedure btnMinusculaClick(Sender: TObject);
    procedure btnMaiusculasClick(Sender: TObject);
    procedure btnPrimeiraMaiusculaClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSubstituir: TfrmSubstituir;

implementation

uses
  importa_excel, uUtil;

{$R *.dfm}

procedure TfrmSubstituir.btnMaiusculasClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  edtOldValue.Text := AnsiUpperCase(edtOldValue.Text);

  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := AnsiUpperCase(temp);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := AnsiUpperCase(temp);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := AnsiUpperCase(temp);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.btnMinusculaClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  edtOldValue.Text := AnsiLowerCase(edtOldValue.Text);

  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := AnsiLowerCase(temp);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := AnsiLowerCase(temp);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := AnsiLowerCase(temp);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.btnPrimeiraMaiusculaClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  edtOldValue.Text := PrimeiraMaiuscula(edtOldValue.Text);

  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := PrimeiraMaiuscula(temp);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := PrimeiraMaiuscula(temp);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := PrimeiraMaiuscula(temp);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.btnRemoveAcentosClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  edtOldValue.Text := RemoveAcento(edtOldValue.Text);

  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := RemoveAcento(temp);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := RemoveAcento(temp);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := RemoveAcento(temp);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.btnReplaceClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.FormShow(Sender: TObject);
begin
  edtOldValue.Text := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];

  if edtOldValue.Text = '' then
    edtOldValue.SetFocus
  else edtNewValue.SetFocus;
end;

end.
