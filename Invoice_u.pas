unit Invoice_u;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, Buttons, Spin, StdCtrls, ComCtrls,
  DBCtrls, ExtCtrls, clsTheInvoice, jpeg;

type
  TForm2 = class(TForm)
    shpInvoice: TShape;
    Shape2: TShape;
    Shape1: TShape;
    shpDivider: TShape;
    shpHeader: TShape;
    imgLogo: TImage;
    shpProducts: TShape;
    btnFrozen: TSpeedButton;
    btnPrepared: TSpeedButton;
    btnGroceries: TSpeedButton;
    btnPackaging: TSpeedButton;
    btnVegetables: TSpeedButton;
    lblSearch: TLabel;
    lblQuantity: TLabel;
    lblDescription: TLabel;
    lblTotal: TLabel;
    lblProduct: TLabel;
    lblQuantityName: TLabel;
    lblTotalPrice: TLabel;
    lblEqualName: TLabel;
    lblPriceTag: TLabel;
    lblInvoice: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    dbtProduct: TDBText;
    dbtPrice: TDBText;
    Label3: TLabel;
    Shape3: TShape;
    lblFinalTotal: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    redOut: TRichEdit;
    edtSearch: TEdit;
    sedQuantity: TSpinEdit;
    btnAdd: TBitBtn;
    btnUndo: TBitBtn;
    dbgProducts: TDBGrid;
    edtInvoice: TEdit;
    edtOrder: TEdit;
    edtDelivery: TEdit;
    DataSource1: TDataSource;
    OpenDialog1: TOpenDialog;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    btnNew: TButton;
    btnSave: TButton;
    procedure btnAddClick(Sender: TObject);
    procedure btnFrozenClick(Sender: TObject);
    procedure btnGroceriesClick(Sender: TObject);
    procedure btnNewClick(Sender: TObject);
    procedure btnPackagingClick(Sender: TObject);
    procedure btnPreparedClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnUndoClick(Sender: TObject);
    procedure btnVegetablesClick(Sender: TObject);
    procedure dbgProductsCellClick(Column: TColumn);
    procedure edtSearchKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
    procedure imgLogoClick(Sender: TObject);
    procedure sedQuantityChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    //helper methods
    procedure MenuShow(sMenu : string);
    procedure PriceChange;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  sPath : string;
  TheInvoice : TTheInvoice;

implementation

uses Menu_u;

{$R *.dfm}

procedure TForm2.btnAddClick(Sender: TObject);
var
  sDescription, sUnit, sTotal, sChange : string;
  rUnit, rTOtal : real;
  iQuantity, iOn : integer;
begin
  //can undo now
  btnUndo.Enabled := true;
  //get all variables ready for create
  iQuantity := sedQuantity.Value;
  sDescription := dbtProduct.Caption;
  sTotal := lblPriceTag.Caption;
  sTotal := copy(sTotal, 3, length(sTotal)-2);
  sChange := StringReplace(sTotal, ',', '.',
              [rfReplaceAll, rfIgnoreCase]);
  rUnit := strtofloat(sChange);
  rTotal := rUnit;
  rUnit := rUnit/sedQuantity.Value;
  sUnit := floattostrf(rUnit, ffFixed, 5,2);
  sUnit := StringReplace(sUnit, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
  //add item
  TheInvoice.AddItem(iQuantity, sDescription, rUnit, rTotal);
  //setup RedOut
  redOut.Paragraph.TabCount := 4;
  redOut.Paragraph.Tab[0] := 25;     
  redOut.Paragraph.Tab[1] := 80;
  redOut.Paragraph.Tab[2] := 218;
  redOut.Paragraph.Tab[3] := 300;
  iOn := TheInvoice.getSet-1;
  redOut.Lines.Add(TheInvoice.getItem(iOn).toString);
  //display new total
  lblFinalTotal.Caption := TheInvoice.getFinalTotal;
end;

procedure TForm2.btnFrozenClick(Sender: TObject);
begin
  MenuShow('FROZEN AND FRESH SEAFOOD');
end;

procedure TForm2.btnGroceriesClick(Sender: TObject);
begin
  MenuShow('GROCERIES AND STATIONERY');
end;

procedure TForm2.btnNewClick(Sender: TObject);
begin
  btnUndo.Enabled := false;
  redOut.Clear;
  edtInvoice.Clear;
  edtOrder.Clear;
  edtDelivery.Clear;
  lblFinalTotal.Caption := 'R 0,00';
  TheInvoice.Reset;
  edtInvoice.SetFocus;
end;

procedure TForm2.btnPackagingClick(Sender: TObject);
begin
  MenuShow('PACKAGING');
end;

procedure TForm2.btnPreparedClick(Sender: TObject);
begin
  MenuShow('PREPARED FOOD, MARINADES, DRESS');
end;

procedure TForm2.btnSaveClick(Sender: TObject);
var
  iInvNo, iOrderNo, iDeliveryNote : integer;
  sDate, sInvNo, sOrderNo, sDeliveryNote : string;
begin
  //get variables
  sDate := inputbox('Date', 'Please enter the date in the format(mm/dd/yyyy) without any extra 0s e.g. (03):','');
  //defensive programming
  try
    iInvNo := strtoint(edtInvoice.Text);
    iOrderNo := strtoint(edtOrder.Text);
    iDeliveryNote := strtoint(edtDelivery.Text);
  except
    ShowMessage('Please make sure invoice number, order number, and delivery note' +
                ' are all filled in and integers!');
    Exit;
  end;
  //make sure invoice isn't empty
  if redOut.Text = '' then
    begin
      ShowMessage('Invoice is empty. Nothing to save.');
      Exit;
    end;
  //make sure date is given
  if sDate = '' then
    begin
      ShowMessage('The date must be given.');
      Exit;
    end;
  //get variables ready
  sInvNo := inttostr(iInvNo);
  sOrderNo := inttostr(iOrderNo);
  sDeliveryNote := inttostr(iDeliveryNote);
  //save
  TheInvoice.SaveToFile(sPath, sInvNo, sOrderNo, sDeliveryNote, sDate);
  //clear invoice
  btnNew.Click;  
end;

procedure TForm2.btnUndoClick(Sender: TObject);
var
  buttonselected : integer;
begin
  //confirmation
  buttonselected := Messagedlg('Are you sure?', mtConfirmation, [mbOK, mbCancel], 0);
  if buttonselected = mrCancel then Exit;
  TheInvoice.UndoItem;
  redOut.Lines.Delete(redOut.Lines.Count-1);
  lblFinalTotal.Caption := TheInvoice.getFinalTotal;
  //decide whether undo should be displayed
  if redOut.Text = ''
    then btnUndo.Enabled := false;
end;

procedure TForm2.btnVegetablesClick(Sender: TObject);
begin
  MenuShow('VEGETABLES AND DESERTS');
end;

procedure TForm2.dbgProductsCellClick(Column: TColumn);
begin
  PriceChange;
end;

procedure TForm2.edtSearchKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  sSearch, sText : string;
  iPos : integer;
begin
  //check if anything is searched
  sSearch := edtSearch.Text;
  if edtSearch.Text = '' then
    begin
      Exit;
    end;

  if ADOQuery1.SQL.Text = '' then
    begin
      ShowMessage('Select a category please');
      Exit;
    end;
  //start searching
  sSearch := '%' + sSearch + '%';
  sSearch := quotedstr(sSearch);
  if ADOQuery1.SQL.Text <> '' then
    begin
      sText := ADOQuery1.SQL.Text;
      if pos('WHERE', sText) > 0 then
        begin
          iPos := pos('WHERE', sText);
          sText := copy(sText, 1, iPos + 5);
          sText := sText + '(PRODUCT LIKE ' + sSearch + ') OR (PRICE LIKE ' + sSearch + ')';
          ADOQuery1.Close;
          ADOQuery1.SQL.Text := sText;
          ADOQuery1.Open;
        end
      else
        begin
          ADOQuery1.Close;
          ADOQuery1.SQL.Add('WHERE (PRODUCT LIKE ' + sSearch + ') OR (PRICE LIKE ' + sSearch + ')');
          ADOQuery1.Open;
        end;
    end;
  PriceChange;
end;

procedure TForm2.imgLogoClick(Sender: TObject);
begin
  //return to menu
  Form2.Hide;
  Form1.Show;
  ADOQuery1.Close;
  ADOQuery1.ConnectionString := '';
  TheInvoice.destroy;
end;

procedure TForm2.sedQuantityChange(Sender: TObject);
begin
  PriceChange;
end;

procedure TForm2.FormActivate(Sender: TObject);
begin
  btnUndo.Enabled := false;
  ADOQuery1.Close;
  ADOQuery1.ConnectionString := '';
  sPath := '';
  ShowMessage('Select Prices of Products Excel File');
  if opendialog1.Execute
    then sPath := opendialog1.FileName
    else Exit;
  //prices of products.xlsx
  ADOQuery1.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+ sPath + ';Extended Properties="Excel 12.0 Xml;HDR=yes";Persist Security Info=False';
  TheInvoice := TTheInvoice.create;
end;

//helper method
procedure TForm2.MenuShow(sMenu: string);
begin
  //Show correct category with the products
  with ADOQuery1 do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT *');
      SQL.Add('FROM [' + sMenu + '$]');
      Open;
    end;
  PriceChange;
end;

procedure TForm2.PriceChange;
var
  sPrice, sCut, sConvert, sFinal : string;
  rPrice, rTotal : real;
  iPos, iKos, iQuantity : integer;
begin
  //change the pricetag of the product when changed
  sPrice := dbtPrice.Caption;
  if sPrice = '' then
    begin
      lblPriceTag.Caption := 'R 0,00';
      Exit;
    end;
  iPos := pos('R ', sPrice);
  iKos := pos(',', sPrice);
  sCut := copy(sPrice, iPos +2, iKos-iPos+2);
  sConvert := StringReplace(sCut, ',', '.',
              [rfReplaceAll, rfIgnoreCase]);
  rPrice := strtofloat(sConvert);
  iQuantity := sedQuantity.Value;
  rTotal := rPrice * iQuantity;
  sFinal := floattostrf(rTotal, ffFixed, 5, 2);
  sFinal :=   StringReplace(sFinal, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
  lblPriceTag.Caption := 'R ' + sFinal;
end;

end.
