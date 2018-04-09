unit uFrmTabela;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Math,
  Dialogs, Grids, Wwdbigrd, Wwdbgrid, DB, DBClient, StdCtrls, Mask, DBCtrls,
  ComObj, Buttons, ExtCtrls;

type
  TfrmTabela = class(TForm)
    dsoTabela: TDataSource;
    gridTabela: TwwDBGrid;
    GroupBox1: TGroupBox;
    edVlInicial: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    edAmplitude: TEdit;
    Label3: TLabel;
    edNrIntervalos: TEdit;
    btCalcular: TButton;
    tabTabela: TClientDataSet;
    tabTabelanr_indice: TIntegerField;
    tabTabelavl_inicial: TFloatField;
    tabTabelavl_final: TFloatField;
    tabTabelavl_frequencia: TIntegerField;
    tabTabelavl_xifi: TFloatField;
    tabTabelavl_fri: TFloatField;
    tabTabelavl_sumFi: TFloatField;
    tabTabelapc_fri: TFloatField;
    tabTabelavl_yi: TIntegerField;
    tabTabelavl_fiyi2: TFloatField;
    tabTabelavl_yifi: TFloatField;
    tabTabelavl_fixi2: TFloatField;
    tabTabelavl_xi: TFloatField;
    GroupBox2: TGroupBox;
    edEfi: TEdit;
    Label4: TLabel;
    edXo: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    edExifi: TEdit;
    Label7: TLabel;
    edEyifi: TEdit;
    edEfixi2: TEdit;
    edEfiyi2: TEdit;
    edMedia: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    edMediana: TEdit;
    Label11: TLabel;
    Label12: TLabel;
    edModa: TEdit;
    Label13: TLabel;
    edQ1: TEdit;
    Label14: TLabel;
    edQ3: TEdit;
    Label15: TLabel;
    edCV: TEdit;
    btExportar: TSpeedButton;
    btLimpar: TButton;
    Label16: TLabel;
    edDesvioP: TEdit;
    cbCor: TColorBox;
    procedure btCalcularClick(Sender: TObject);
    procedure tabTabelavl_inicialValidate(Sender: TField);
    procedure tabTabelavl_frequenciaValidate(Sender: TField);
    procedure tabTabelaAfterPost(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure tabTabelavl_yiChange(Sender: TField);
    procedure btExportarClick(Sender: TObject);
    procedure btLimparClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure gridTabelaDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
    procedure cbCorSelect(Sender: TObject);
  private
    { Private declarations }
    procedure attLinha();
    procedure SendToExcel(aDataSet: TDataSet);
    function ReplaceStr(Text, OldString, NewString: String): String;
    { Public declarations }
  end;

var
  frmTabela: TfrmTabela;
  valida: boolean;

implementation

uses uFrmMain;
{$R *.dfm}

procedure TfrmTabela.attLinha;
var
  evento: TFieldNotifyEvent;
  vlFi, f0, vlFRI: real;
  book: TBookmark;
  soma_fixi, soma_fixi2, soma_yifi, soma_yifi2: real;
  med_x, med_anterior: real;
  q1_x, q1_anterior: real;
  q3_x, q3_anterior: real;
  mediana, q1, q3: boolean;
begin
  if not valida then begin
    try
      try
        valida := true;
        vlFi := 0;
        f0 := 0;
        q1_x := 0;
        med_anterior := 0;
        vlFRI := 0;
        soma_fixi := 0;
        soma_fixi2 := 0;
        soma_yifi := 0;
        soma_yifi2 := 0;
        med_x := 0;
        evento := tabTabelavl_frequencia.OnValidate;
        tabTabelavl_frequencia.OnValidate := nil;
        book := tabTabela.GetBookmark;
        tabTabela.First;
        if tabTabela.RecordCount > 0 then
          // Laço para carregar o Efi
          while not tabTabela.Eof do begin
            tabTabela.Edit;
            if tabTabelavl_frequencia.IsNull then
              break;
            tabTabelavl_sumFi.AsFloat := vlFi + tabTabelavl_frequencia.AsFloat;
            if tabTabelavl_frequencia.AsFloat > f0 then begin
              f0 := tabTabelavl_frequencia.AsFloat;
              edXo.Text := FormatFloat('#,##0.00',tabTabelavl_xi.AsFloat);
            end;
            vlFi := tabTabelavl_sumFi.AsFloat;
            tabTabela.post;
            tabTabela.Next;
          end;
        edEfi.Text := FloatToStr(vlFi);
        // Laço para lançar a % sobre o montante total, Fri
        tabTabela.First;
        while not tabTabela.Eof do begin
          if tabTabelavl_frequencia.IsNull or tabTabelavl_sumFi.IsNull then
            break;
          tabTabela.Edit;
          if StrToFloat(edEfi.Text) > 0 then
            tabTabelavl_fri.AsFloat := (tabTabelavl_frequencia.AsFloat / StrToFloat(edEfi.Text))
          else
            tabTabelavl_fri.AsFloat := 0;
          tabTabelapc_fri.AsFloat := vlFRI + tabTabelavl_fri.AsFloat;
          vlFRI := tabTabelapc_fri.AsFloat;
          tabTabelavl_yi.AsInteger := Trunc((tabTabelavl_xi.AsFloat - StrToFloat(edXo.Text)) / StrToInt(edAmplitude.Text));
          soma_fixi := soma_fixi + tabTabelavl_xifi.AsFloat;
          soma_fixi2 := soma_fixi2 + tabTabelavl_fixi2.AsFloat;
          soma_yifi := soma_yifi + tabTabelavl_yifi.AsFloat;
          soma_yifi2 := soma_yifi2 + tabTabelavl_fiyi2.AsFloat;
          tabTabela.post;
          tabTabela.Next;
        end;
        edModa.Text := edXo.Text;
        edExifi.Text := FloatToStr(soma_fixi);
        edEfixi2.Text := FloatToStr(soma_fixi2);
        edEyifi.Text := FloatToStr(soma_yifi);
        edEfiyi2.Text := FloatToStr(soma_yifi2);
        // Media
        if StrToInt(edEfi.Text) > 0 then begin
          edMedia.Text := FormatFloat('#,##0.00', StrToFloat(edXo.Text) + ((StrToFloat(edEyifi.Text) * StrToFloat(edAmplitude.Text))) / StrToFloat(edEfi.Text));
          med_x := StrToFloat(edEfi.Text) / 2;
          q1_x := StrToFloat(edEfi.Text) / 4;
          q3_x := (3 * StrToFloat(edEfi.Text)) / 4;
        end;

        // Mediana
        tabTabela.First;
        while not tabTabela.Eof do begin
          if not mediana then
            if (tabTabelavl_sumFi.AsFloat > med_x) then begin
              if tabTabelavl_frequencia.AsFloat > 0 then
                med_x := tabTabelavl_inicial.AsFloat + ((med_x - med_anterior) * StrToFloat(edAmplitude.Text)) / tabTabelavl_frequencia.AsFloat
              else
                med_x := tabTabelavl_inicial.AsFloat + 0;
              mediana := true;
            end;
          if not q1 then
            if (tabTabelavl_sumFi.AsFloat > q1_x) then begin
              if tabTabelavl_frequencia.AsFloat > 0 then
                q1_x := tabTabelavl_inicial.AsFloat + ((q1_x - q1_anterior) * StrToFloat(edAmplitude.Text)) / tabTabelavl_frequencia.AsFloat
              else
                q1_x := tabTabelavl_inicial.AsFloat + 0;
              q1 := true;
            end;
          if not q3 then
            if (tabTabelavl_sumFi.AsFloat > q3_x) then begin
              if tabTabelavl_frequencia.AsFloat > 0 then
                q3_x := tabTabelavl_inicial.AsFloat + ((q3_x - q3_anterior) * StrToFloat(edAmplitude.Text)) / tabTabelavl_frequencia.AsFloat
              else
                q3_x := tabTabelavl_inicial.AsFloat + 0;
              q3 := true;
            end;

          q1_anterior := tabTabelavl_sumFi.AsFloat;
          med_anterior := tabTabelavl_sumFi.AsFloat;
          q3_anterior := tabTabelavl_sumFi.AsFloat;
          tabTabela.Next;
        end;
        if vlFi > 0 then begin
          edDesvioP.Text := FormatFloat('#,##0.00', StrToFloat(edAmplitude.Text) * sqrt((soma_yifi2 / vlFi) - power(soma_yifi / vlFi, 2)));
          edCV.Text := FormatFloat('#,##0.00%', (StrToFloat(edDesvioP.Text) / StrToFloat(edMedia.Text)) * 100)
        end;
        edMediana.Text := FormatFloat('#,##0.00', med_x);
        edQ1.Text := FormatFloat('#,##0.00', q1_x);
        edQ3.Text := FormatFloat('#,##0.00', q3_x);
        tabTabela.GotoBookmark(book);
        tabTabela.FreeBookmark(book);
        tabTabelavl_frequencia.OnValidate := evento;
      except
        on e: Exception do
          raise Exception.Create('ERROR: ' + e.Message);
      end;
    finally
      valida := false;
    end;
  end;
end;

function TfrmTabela.ReplaceStr(Text, OldString, NewString: String): String;
Var
  Atual, StrToFind, OriginalStr: PChar;
  NewText: string;
  LenOldString, LenNewString, m, Index: Integer;
begin // ReplaceStr
  NewText := Text;
  OriginalStr := PChar(Text);
  StrToFind := PChar(OldString);
  LenOldString := Length(OldString);
  LenNewString := Length(NewString);
  Atual := StrPos(OriginalStr, StrToFind);
  Index := 0;
  While Atual <> Nil do Begin // Atual<>nil
    m := Atual - OriginalStr - Index + 1;
    Delete(NewText, m, LenOldString);
    Insert(NewString, NewText, m);
    Inc(Index, LenOldString - LenNewString);
    Atual := StrPos(Atual + LenOldString, StrToFind);
  end; // Atual<>nil
  Result := NewText;
end; // ReplaceStr

procedure TfrmTabela.SendToExcel(aDataSet: TDataSet);
var
  coluna, linha: Integer;
  excel: variant;
  valor: string;
begin
  excel := CreateOleObject('Excel.Application');
  excel.Workbooks.add(1);
  aDataSet.First;
  try
    for linha := 0 to aDataSet.RecordCount - 1 do begin
      for coluna := 1 to aDataSet.FieldCount do begin // eliminei a coluna 0 da relação do Excel
        // trato o campo data
        if aDataSet.Fields[coluna - 1].DataType = ftDate then begin
          excel.cells[linha + 2, coluna] := aDataSet.Fields[coluna - 1].AsDateTime;
        end
        else begin
          valor := ReplaceStr(aDataSet.Fields[coluna - 1].AsString, ',', '.');
          excel.cells[linha + 2, coluna] := valor;
        end;
      end;
      aDataSet.Next;
    end;
    linha := linha + 3;
    // preencher os labels
    excel.cells[linha, 1] := 'Saídas';
    excel.cells[linha + 1, 1] := Label4.Caption;
    excel.cells[linha + 2, 1] := Label5.Caption;
    excel.cells[linha + 3, 1] := Label6.Caption;
    excel.cells[linha + 4, 1] := Label7.Caption;
    excel.cells[linha + 5, 1] := Label8.Caption;
    excel.cells[linha + 6, 1] := Label9.Caption;
    excel.cells[linha + 7, 1] := Label10.Caption;
    excel.cells[linha + 8, 1] := Label11.Caption;
    excel.cells[linha + 9, 1] := Label12.Caption;
    excel.cells[linha + 10, 1] := Label13.Caption;
    excel.cells[linha + 11, 1] := Label14.Caption;
    excel.cells[linha + 12, 1] := Label15.Caption;
    excel.cells[linha + 13, 1] := Label16.Caption;
    // preencher os valores
    excel.cells[linha + 1, 2] := edEfi.Text;
    excel.cells[linha + 2, 2] := edXo.Text;
    excel.cells[linha + 3, 2] := edExifi.Text;
    excel.cells[linha + 4, 2] := edEyifi.Text;
    excel.cells[linha + 5, 2] := edEfixi2.Text;
    excel.cells[linha + 6, 2] := edEfiyi2.Text;
    excel.cells[linha + 7, 2] := edMedia.Text;
    excel.cells[linha + 8, 2] := edModa.Text;
    excel.cells[linha + 9, 2] := edMediana.Text;
    excel.cells[linha + 10, 2] := edQ1.Text;
    excel.cells[linha + 11, 2] := edQ3.Text;
    excel.cells[linha + 12, 2] := edCV.Text;
    excel.cells[linha + 13, 2] := ReplaceStr(edDesvioP.Text, '.',',');

    for coluna := 1 to aDataSet.FieldCount do begin // eliminei a coluna 0 da relação do Excel
      valor := aDataSet.Fields[coluna - 1].DisplayLabel;
      excel.cells[1, coluna] := valor;
    end;

    excel.columns.AutoFit; // esta linha é para fazer com que o Excel dimencione as células adequadamente.
    excel.visible := true;
    // se um dia precisar salvar a planilha
    // excel.ActiveWorkbook.SaveAs('c:\teste.xls');
  except
    Application.MessageBox('Aconteceu um erro desconhecido durante a conversão' + 'da tabela para o Ms-Excel', 'Erro', MB_OK + MB_ICONEXCLAMATION);
  end;
  aDataSet.First;
end;

procedure TfrmTabela.btCalcularClick(Sender: TObject);
var
  i: Integer;
  vlanterior: real;
begin
  tabTabela.EmptyDataSet;
  vlanterior := StrToFloat(edVlInicial.Text);
  for i := 0 to StrToInt(edNrIntervalos.Text) - 1 do begin
    tabTabela.Append;
    tabTabelanr_indice.AsInteger := i + 1;
    tabTabelavl_xi.AsFloat := (vlanterior + vlanterior + StrToFloat(edAmplitude.Text)) / 2;
    tabTabelavl_inicial.AsFloat := vlanterior;
    tabTabela.post;
    vlanterior := vlanterior + StrToFloat(edAmplitude.Text);
    tabTabela.Next;
  end;
  btCalcular.Enabled := false;
  edVlInicial.Enabled := false;
  edAmplitude.Enabled := false;
  edNrIntervalos.Enabled := false;
  btLimpar.Enabled := true;
end;

procedure TfrmTabela.btExportarClick(Sender: TObject);
begin
  SendToExcel(tabTabela);
end;

procedure TfrmTabela.btLimparClick(Sender: TObject);
begin
  tabTabela.EmptyDataSet;
  btLimpar.Enabled := false;
  edVlInicial.Enabled := true;
  edAmplitude.Enabled := true;
  edNrIntervalos.Enabled := true;
  btCalcular.Enabled := true;
end;

procedure TfrmTabela.cbCorSelect(Sender: TObject);
begin
  gridTabela.Repaint;
end;

procedure TfrmTabela.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmMain.show;
end;

procedure TfrmTabela.FormShow(Sender: TObject);
begin
  if tabTabela.Active then
    tabTabela.First;
end;

procedure TfrmTabela.tabTabelaAfterPost(DataSet: TDataSet);
begin
  attLinha;
end;

procedure TfrmTabela.tabTabelavl_frequenciaValidate(Sender: TField);
begin
  tabTabelavl_xifi.AsFloat := tabTabelavl_frequencia.AsFloat * tabTabelavl_xi.AsFloat;
  tabTabelavl_fixi2.AsFloat := tabTabelavl_xifi.AsFloat * tabTabelavl_xi.AsFloat;
end;

procedure TfrmTabela.tabTabelavl_inicialValidate(Sender: TField);
begin
  if not tabTabelavl_inicial.IsNull then
    tabTabelavl_final.AsFloat := tabTabelavl_inicial.AsFloat + StrToFloat(edAmplitude.Text);
end;

procedure TfrmTabela.tabTabelavl_yiChange(Sender: TField);
begin
  if not tabTabelavl_frequencia.IsNull then begin
    tabTabelavl_yifi.AsFloat := tabTabelavl_frequencia.AsInteger * tabTabelavl_yi.AsFloat;
    tabTabelavl_fiyi2.AsFloat := tabTabelavl_yifi.AsFloat * tabTabelavl_yi.AsFloat;
  end;

end;

procedure TfrmTabela.gridTabelaDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
begin
  if Field = tabTabelavl_frequencia then begin
    gridTabela.Canvas.Font.Color := clBlack;
    gridTabela.Canvas.Brush.Color := cbCor.Selected;
    gridTabela.Canvas.FillRect(Rect);
    gridTabela.DefaultDrawDataCell(Rect,field,State);
  end;
end;

end.
