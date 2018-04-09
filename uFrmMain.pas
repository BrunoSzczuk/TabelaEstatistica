unit uFrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TfrmMain = class(TForm)
    edVlInicial: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    edAmplitude: TEdit;
    Label3: TLabel;
    edNrIntervalos: TEdit;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses uFrmTabela;

{$R *.dfm}

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  frmTabela.edVlInicial.Text := edVlInicial.Text;
  frmTabela.edAmplitude.Text := edAmplitude.Text;
  frmTabela.edNrIntervalos.Text := edNrIntervalos.Text;
  frmTabela.btCalcularClick(self);
  frmMain.Hide;
  frmTabela.ShowModal;
end;

end.
