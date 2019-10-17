unit Statement_u;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Buttons, ComCtrls, ExtCtrls, ComObj,
  clsTheStatement, jpeg;

type
  TForm3 = class(TForm)
    shpHeader: TShape;
    imgLogo: TImage;
    shpProducts: TShape;
    lblStatement: TLabel;
    lblDescription: TLabel;
    lblTotal: TLabel;
    lblCapTotal: TLabel;
    shpTotal: TShape;
    lblFinalTotal: TLabel;
    lblNameOfFile: TLabel;
    redOut: TRichEdit;
    btnAdd: TBitBtn;
    btnUndo: TBitBtn;
    edtMonth: TEdit;
    btnDone: TBitBtn;
    btnNew: TButton;
    OpenDialog2: TOpenDialog;
    OpenDialog1: TOpenDialog;
    procedure btnAddClick(Sender: TObject);
    procedure btnDoneClick(Sender: TObject);
    procedure btnNewClick(Sender: TObject);
    procedure btnUndoClick(Sender: TObject);
    procedure imgLogoClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
  TheStatement : TTheStatement;

implementation

uses Menu_u;

{$R *.dfm}

procedure TForm3.btnAddClick(Sender: TObject);
var
  sPAth, sInvoice, sPrice : string;
  iPos, iKos, iOn : integer;
  XL : OLEVariant;
begin
  //enable undo
  btnUndo.Enabled := true;
  //get invoice file
  sInvoice := '';
  if opendialog1.Execute
    then sPath := opendialog1.FileName;
  XL := CreateOleObject('Excel.Application');
  XL.WorkBooks.Open(sPath);
  sPrice := XL.ActiveSheet.Cells[49,7].Value;
  XL.WorkBooks.Close;
  //get variables
  iPos := pos('Invoice ', sPath);
  iKos := pos('.xlsx', sPath);
  sInvoice := copy(sPath, iPos, iKos-iPos);
  //add invoice
  TheStatement.AddInvoice(sInvoice, sPrice);
  iOn := TheStatement.getSet-1;
  //get redOut ready
  redOUt.Paragraph.TabCount := 1;
  redOut.Paragraph.Tab[0] := 150;
  //display it
  redOUt.Lines.Add(TheStatement.getInvoice(iOn).toString);
  //display price
  lblFinalTotal.Caption := TheStatement.getFinalTotal;
  //free from memory
  XL.Quit;
  XL := Unassigned;
end;

procedure TForm3.btnDoneClick(Sender: TObject);
var
  sDate, ssPath : string;
begin
  //defensive programming
  if redOut.Text = '' then
    begin
      Showmessage('Your statement is empty');
      Exit;
    end;
  if edtMonth.Text = '' then
    begin
      Showmessage('Please enter the month and year');
      Exit;
    end;
  //get variables ready
  sDate := edtMonth.Text;
  if opendialog2.Execute then
    begin
      ssPath := opendialog2.FileName;       //formatted statement
    end;

  TheStatement.Done(sDate, ssPath);
  //clear statmenet
  btnNew.Click;
end;

procedure TForm3.btnNewClick(Sender: TObject);
begin
  redOut.Clear;
  edtMonth.Clear;
  lblFinalTotal.Caption := 'R 0,00';
  btnUndo.Enabled := false;
  TheStatement.Reset;
  edtMonth.SetFocus;
end;

procedure TForm3.btnUndoClick(Sender: TObject);
var
  buttonselected : integer;
begin
  //confirmation
  buttonselected := Messagedlg('Are you sure?', mtConfirmation, [mbOK, mbCancel], 0);
  if buttonselected = mrCancel then Exit;
  TheStatement.UndoInvoice;         
  redOut.Lines.Delete(redOut.Lines.Count-1);
  lblFinalTotal.Caption := TheStatement.getFinalTotal;
  //decide whether undo should still be enabled or not
  if redOut.Text = ''
    then btnUndo.Enabled := false;
end;

procedure TForm3.imgLogoClick(Sender: TObject);
begin
  Form3.Hide;
  Form1.Show;
  TheStatement.destroy;
end;

procedure TForm3.FormActivate(Sender: TObject);
begin
  //create statement
  TheStatement := TTheStatement.create;
end;

end.
