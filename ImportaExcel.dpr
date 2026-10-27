program ImportaExcel;

uses
  Vcl.Forms,
  importa_excel in 'importa_excel.pas' {frmPrinc},
  importando in 'importando.pas' {frmImportando},
  empresa in 'empresa.pas' {frmEmpr},
  Colunas in 'Colunas.pas' {frmColunas},
  uUtil in 'Units\uUtil.pas',
  uImportaClieForn in 'Units\uImportaClieForn.pas',
  uImportaProduto in 'Units\uImportaProduto.pas',
  uImportaGrupo in 'Units\uImportaGrupo.pas',
  uImportaSubGrupo in 'Units\uImportaSubGrupo.pas',
  uImportaMarca in 'Units\uImportaMarca.pas',
  uImportaTran in 'Units\uImportaTran.pas',
  uImportaTituP in 'Units\uImportaTituP.pas',
  uImportaTituR in 'Units\uImportaTituR.pas',
  uImportaCUSTOM in 'Units\uImportaCUSTOM.pas',
  uSubstituir in 'Units\uSubstituir.pas' {frmSubstituir},
  uProcurar in 'Units\uProcurar.pas' {frmProcurar},
  uDividir in 'Units\uDividir.pas' {frmDividir},
  uTema in 'Units\uTema.pas';

{$R *.res}
//Temas embutidos (menu Tema)
{$R 'Temas.res' 'Temas.rc'}

begin
  Application.Initialize;
  //Tema salvo pelo usuário, antes de criar as telas
  IniciaTema;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrinc, frmPrinc);
  Application.CreateForm(TfrmImportando, frmImportando);
  Application.CreateForm(TfrmEmpr, frmEmpr);
  Application.CreateForm(TfrmColunas, frmColunas);
  Application.CreateForm(TfrmSubstituir, frmSubstituir);
  Application.CreateForm(TfrmProcurar, frmProcurar);
  Application.CreateForm(TfrmDividir, frmDividir);
  Application.Run;
end.
