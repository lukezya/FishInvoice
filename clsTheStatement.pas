unit clsTheStatement;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, clsInvoice, ADODB;

type
  TTheStatement = class
  private
    fSet : integer;
    fArrInvoice : array of TInvoice;
    fFinalTotal : string;
  public
    constructor create;
    destructor destroy; override;
    function getSet : integer;
    function getFinalTotal : string;
    function getInvoice(iNo:integer):TInvoice;
    procedure setSet(iSet : integer);
    procedure setFinalTotal(sFinalTotal : string);
    procedure AddInvoice(sInvoice, sPrice : string);
    procedure UndoInvoice;
    procedure Reset;
    procedure Done(sDate, sPath : string);
    //helper methods
    procedure ConvertPoint(var sComma:string);
    procedure ConvertComma(var sPoint:string);
  end;

implementation

{ TTheStatement }

procedure TTheStatement.AddInvoice(sInvoice, sPrice: string);
var
  iAdd : integer;
  sT, sAdd : string;
  rGlobal, rAdd : real;
begin
  //add item to array
  inc(fSet);
  //max 27
  if fSet > 27 then
    begin
      ShowMessage('Item could not be added, invoice sheet item space has reached ' +
                  'its maximum!');
      Exit;
    end;
  //carry on
  SetLength(fArrInvoice, fSet);
  iAdd := fSet-1;
  fArrInvoice[iAdd] := TInvoice.create(sInvoice, sPrice);
  //change finaltotal
  sT := fFinalTotal;
  sT := copy(sT, 3, length(sT)-2);
  ConvertPoint(sT);

  sAdd := copy(fArrInvoice[iAdd].getPrice, 3, length(fArrInvoice[iAdd].getPrice)-2);
  ConvertPOint(sAdd);

  rGlobal := strtofloat(sT);
  rAdd := strtofloat(sAdd);
  rGlobal := rGlobal + rAdd;

  fFinalTotal := 'R ' + floattostrf(rGlobal, ffFixed, 5, 2);
  ConvertComma(fFinalTotal);
end;

procedure TTheStatement.ConvertComma(var sPoint: string);
begin
  sPoint := StringReplace(sPoint, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
end;

procedure TTheStatement.ConvertPoint(var sComma: string);
begin
   sComma := StringReplace(sComma, ',', '.',
              [rfReplaceAll, rfIgnoreCase]);
end;

constructor TTheStatement.create;
begin
  //default
  fSet := 0;
  fFinalTotal := 'R 0.00';
end;

destructor TTheStatement.destroy;
begin

  inherited;
end;

procedure TTheStatement.Done(sDate, sPath: string);
var
  sAPath, sc1, sc2, sd1, sHead : string;
  ADOQuery : TADOQuery;
  ic1, iLoop : integer;
begin
  //connections and paths
  sAPath := copy(sPath, 1, length(sPath)-24);
  sAPath := SAPath + 'Statements\' + sDate + '.xlsx';

  CopyFile(PChar(sPath), pChar(sAPath), true);

  ADOQuery := TADOQuery.Create(Application);
  ADOQuery.ConnectionString := '';
  ADOQuery.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' +
    sAPath + ';Extended Properties="Excel 12.0 Xml;HDR=yes";Persist Security Info=False';
  ADOQuery.ParamCheck := False;
  //items
  ic1 := 9;
  for iLoop := 0 to fSet-1 do
    begin
      inc(ic1);
      sc1 := 'B';
      sc2 := sc1 + inttostr(ic1); //B22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' + quotedstr(fArrInvoice[iLoop].getInvoice) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;

      sc1 := 'C';
      sc2 := sc1 + inttostr(ic1); //C22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' + quotedstr(fArrInvoice[iLoop].getPrice) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;
    end;
  //other info enter
  sHead := 'INVOICE STATEMENT FOR ' + uppercase(sDate);
  sd1 := 'INSERT INTO [Sheet1$B8:B8] VALUES(' + quotedstr(sHead) + ')';
  ADOQuery.SQL.Text := sd1;
  ADOQuery.ExecSQL;

  sd1 := 'INSERT INTO [Sheet1$C37:C37] VALUES(' + quotedstr((fFinalTotal)) + ')';
  ADOQuery.SQL.Text := sd1;
  ADOQuery.ExecSQL;
  ShowMessage('Your statement has been processed, you will find it under the folder Statements');
  ADOQuery.Close;
  ADOQuery.ConnectionString := '';
end;

function TTheStatement.getFinalTotal: string;
begin
  result := fFinalTotal;
end;

function TTheStatement.getInvoice(iNo: integer): TInvoice;
begin
  result := fArrInvoice[iNo];
end;

function TTheStatement.getSet: integer;
begin
  result := fSet;
end;

procedure TTheStatement.Reset;
begin
  fSet := 0;
  fFinalTotal := 'R 0.00';
  SetLength(fArrInvoice, 0);
end;

procedure TTheStatement.setFinalTotal(sFinalTotal: string);
begin
  fFinalTotal := sFinalTotal;
end;

procedure TTheStatement.setSet(iSet: integer);
begin
  fSet := iSet;
end;

procedure TTheStatement.UndoInvoice;
var
  rGlobal, rMinus : real;
  sT, sMinus : string;
begin
  //change setLength
  fSet := fSet-1;
  sMinus := fArrInvoice[fSet].getPrice;
  SetLength(fArrInvoice, fSet);
  //change finaltotal
  sT := fFinalTotal;
  sT := copy(sT, 3, length(sT)-2);
  ConvertPoint(sT);

  Delete(sMinus, 1, 2);
  ConvertPoint(sMinus);

  rGlobal := strtofloat(sT);
  rMinus := strtofloat(sMinus);
  rGlobal := rGlobal - rMinus;

  fFinalTotal := 'R ' + floattostrf(rGlobal, ffFixed, 5, 2);
  ConvertComma(fFinalTotal);
end;

end.
 