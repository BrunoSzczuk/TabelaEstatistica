program TabelaEstatistica;

uses
  Forms,
  uFrmMain in 'uFrmMain.pas' {frmMain},
  uFrmTabela in 'uFrmTabela.pas' {frmTabela};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmTabela, frmTabela);
  Application.Run;
end.
