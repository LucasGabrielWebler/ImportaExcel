unit empresa;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Data.SqlExpr;

type
  TfrmEmpr = class(TForm)
    LabelRazao: TLabel;
    razaoSocial: TEdit;
    LabelFantasia: TLabel;
    nomeFantasia: TEdit;
    Label1: TLabel;
    cnpj: TEdit;
    Label2: TLabel;
    empr: TEdit;
    Label3: TLabel;
    inscr: TEdit;
    simples: TRadioButton;
    normal: TRadioButton;
    Label4: TLabel;
    ende: TEdit;
    Label5: TLabel;
    bair: TEdit;
    Label6: TLabel;
    cep: TEdit;
    Label7: TLabel;
    cida: TComboBox;
    Label8: TLabel;
    fone: TEdit;
    mail: TEdit;
    Label9: TLabel;
    save: TBitBtn;
    cancel: TBitBtn;
    uf: TEdit;
    Label10: TLabel;
    procedure cancelClick(Sender: TObject);
    procedure saveClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure emprChange(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEmpr: TfrmEmpr;

implementation


{$R *.dfm}

uses importa_excel;


//Botão Cancelar (Sair)
procedure TfrmEmpr.cancelClick(Sender: TObject);
begin
  Close;
end;


//Função para executar SQL e retornar algo
function query(comando: string): string;
var
  queryTemp: TSQLQuery;
begin
  try
    queryTemp := TSQLQuery.Create(nil);
    frmPrinc.conDestino.Open;
    queryTemp.SQLConnection := frmPrinc.conDestino;
    queryTemp.SQL.Clear;
    queryTemp.SQL.Add(comando);
    queryTemp.Open;

    Result := queryTemp.Fields[0].AsString;
  finally
    queryTemp.Free;
    frmPrinc.conDestino.Close;
  end;
end;


//Ao criar formulario, carregar dados
procedure TfrmEmpr.emprChange(Sender: TObject);
begin
  if empr.Text <> '' then
    FormActivate(Self);
end;

procedure TfrmEmpr.FormActivate(Sender: TObject);
var
  crt,cidade: string;
begin
  //Carregar valores do banco
  razaoSocial.Text := query('select d.razao from dados_empre d where d.codi='+empr.Text+';');
  nomeFantasia.Text := query('select d.nome_fantasia from dados_empre d where d.codi='+empr.Text+';');
  ende.Text := query('select d.endereco from dados_empre d where d.codi='+empr.Text+';');
  inscr.Text := query('select d.ie from dados_empre d where d.codi='+empr.Text+';');
  fone.Text := query('select d.fone1 from dados_empre d where d.codi='+empr.Text+';');
  mail.Text := query('select d.mail from dados_empre d where d.codi='+empr.Text+';');
  cnpj.Text := query('select d.cgc from dados_empre d where d.codi='+empr.Text+';');
  cep.Text := query('select d.cep from dados_empre d where d.codi='+empr.Text+';');
  bair.Text := query('select d.bairro from dados_empre d where d.codi='+empr.Text+';');
  crt := query('select d.crt from dados_empre d where d.codi='+empr.Text+';');

  cidade := query('select d.codi_cida from dados_empre d where d.codi='+empr.Text+';');
  cida.Text := query('select c.cid_desc from cidade c where c.cid_codi='+cidade+';');
  uf.Text := query('select c.cid_uf from cidade c where c.cid_codi='+cidade+';');

  if (crt='1') then
  begin
    simples.Checked := True;
  end
  else begin
    normal.Checked := True;
  end;

end;


//Botão para salvar no banco
procedure TfrmEmpr.saveClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  crt: Integer;
  cgc,razao,fant: string;
  today : TDateTime;
begin

  //Simples ou Normal
  if simples.Checked then
  begin
    crt:=1;
  end
  else begin
    crt:=3;
  end;

  //Formatar CNPJ
  cgc := Trim(cnpj.Text);
  cgc := stringreplace(cgc, '-', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := stringreplace(cgc, '/', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := stringreplace(cgc, '.', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := (Copy(cgc,1,2))+ '.' + (Copy(cgc,3,3)) + '.' + (Copy(cgc,6,3)) + '/' + (Copy(cgc,9,4)) + '-' + (Copy(cgc,13,2));

  //Tratar ' na razao social e nome fantasia
  razao := stringreplace(UpperCase(razaoSocial.Text), '''', ' ',[rfReplaceAll, rfIgnoreCase]);
  fant := stringreplace(UpperCase(nomeFantasia.Text), '''', ' ',[rfReplaceAll, rfIgnoreCase]);

  try
    try
      //Abrir conexoes
      SQL := TSQLDataSet.Create(Nil);
      frmPrinc.conDestino.Open;
      SQL.SQLConnection := frmPrinc.conDestino;

      //Executar COMANDO
      SQL.CommandText := 'UPDATE DADOS_EMPRE SET '+
                          'RAZAO = '+''''+razao+''''+','+
                          'ENDERECO = '+''''+UpperCase(ende.Text)+''''+','+
                          'CIDADE = '+''''+UpperCase(cida.Text)+''''+','+
                          'UF = '+''''+UpperCase(uf.Text)+''''+','+
                          'IE = '+''''+UpperCase(inscr.Text)+''''+','+
                          'FONE1 = '+''''+UpperCase(fone.Text)+''''+','+
                          'MAIL = '+''''+LowerCase(mail.Text)+''''+','+
                          'CGC = '+''''+cgc+''''+','+
                          'CEP = '+''''+UpperCase(cep.Text)+''''+','+
                          'NOME_FANTASIA = '+''''+fant+''''+','+
                          'CODI_CIDA = '+''''+TfrmPrinc.buscaCidade(UpperCase(cida.Text), UpperCase(uf.Text))+''''+','+
                          'BAIRRO = '+''''+UpperCase(bair.Text)+''''+','+
                          'CRT = '+IntToStr(crt)+' '+
                      'WHERE (CODI = '+''''+empr.Text+''''+');';
      SQL.ExecSQL;

      //Comando para tributação de entrada
      SQL.CommandText := 'delete from ASSOC_CST_CSOSN_XML;';
      SQL.ExecSQL;

      if simples.Checked then
      begin
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (1, '+''''+'00'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (2, '+''''+'10'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (3, '+''''+'20'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (4, '+''''+'30'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (5, '+''''+'40'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (6, '+''''+'41'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (7, '+''''+'50'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (8, '+''''+'51'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (9, '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (10, '+''''+'70'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (11, '+''''+'90'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (12, '+''''+'101'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (13, '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (14, '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (15, '+''''+'201'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (16, '+''''+'202'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (17, '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (18, '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (19, '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (20, '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (21, '+''''+'900'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
      end
      else begin
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                         'VALUES (1, '+''''+'00'+''''+', '+''''+'00'+''''+', NULL, '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (2, '+''''+'10'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (3, '+''''+'20'+''''+', '+''''+'20'+''''+', NULL, '+''''+'20'+''''+', '+''''+'20'+''''+', '+''''+'20'+''''+', '+''''+'20'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (4, '+''''+'30'+''''+', '+''''+'30'+''''+', NULL, '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (5, '+''''+'40'+''''+', '+''''+'40'+''''+', NULL, '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (6, '+''''+'41'+''''+', '+''''+'41'+''''+', NULL, '+''''+'41'+''''+', '+''''+'41'+''''+', '+''''+'41'+''''+', '+''''+'41'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (7, '+''''+'50'+''''+', '+''''+'50'+''''+', NULL, '+''''+'50'+''''+', '+''''+'50'+''''+', '+''''+'50'+''''+', '+''''+'50'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (8, '+''''+'51'+''''+', '+''''+'51'+''''+', NULL, '+''''+'51'+''''+', '+''''+'51'+''''+', '+''''+'51'+''''+', '+''''+'51'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (9, '+''''+'60'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (10, '+''''+'70'+''''+', '+''''+'70'+''''+', NULL, '+''''+'70'+''''+', '+''''+'70'+''''+', '+''''+'70'+''''+', '+''''+'70'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (11, '+''''+'90'+''''+', '+''''+'90'+''''+', NULL, '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (12, '+''''+'101'+''''+', '+''''+'00'+''''+', NULL, '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (13, '+''''+'102'+''''+', '+''''+'00'+''''+', NULL, '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (14, '+''''+'103'+''''+', '+''''+'00'+''''+', NULL, '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'00'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', '+''''+'103'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (15, '+''''+'201'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (16, '+''''+'202'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (17, '+''''+'203'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', '+''''+'203'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (18, '+''''+'300'+''''+', '+''''+'30'+''''+', NULL, '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'30'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', '+''''+'300'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (19, '+''''+'400'+''''+', '+''''+'40'+''''+', NULL, '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'40'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', '+''''+'400'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (20, '+''''+'500'+''''+', '+''''+'60'+''''+', NULL, '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'60'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', '+''''+'500'+''''+', 1);';
        SQL.ExecSQL;
        SQL.CommandText := 'INSERT INTO ASSOC_CST_CSOSN_XML (CODI, XML_CST_CSOSN, ENTRADA_CST, ENTRADA_CSOSN, SAIDA_CST_EST, SAIDA_CST_INTER, SAIDA_CST_EST_CF, SAIDA_CST_INTER_CF, SAIDA_CSOSN_EST, SAIDA_CSOSN_INTER, SAIDA_CSOSN_EST_CF, SAIDA_CSOSN_INTER_CF, EMPR) '+
                                 'VALUES (21, '+''''+'900'+''''+', '+''''+'90'+''''+', NULL, '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'90'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', '+''''+'102'+''''+', 1);';
      end;

      //Liberação níveis de acesso
      today := Now;
      SQL.CommandText := 'UPDATE VDS set VDS.ulti_venc = ' + QuotedStr(stringreplace(DateToStr(today), '/', '.',[rfReplaceAll, rfIgnoreCase]));
      SQL.ExecSQL;

    except
      on e: exception do
      begin
        ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
      end;
    end;

  finally
    SQL.Free;
    frmPrinc.conDestino.Close;
    ShowMessage('Dados da empresa salvos.');
    Close;
  end;
end;

end.
