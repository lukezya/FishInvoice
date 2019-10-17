unit clsTheInvoice;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, clsItem, ADODB, ComObj;

type
  TTheInvoice = class
  private
    fSet : integer;  //max 26
    fArrItem : array of TItem;
    fFinalTotal : string;
  public
    constructor create;
    destructor destroy; override;
    function getSet : integer;
    function getFinalTotal : string;
    procedure setSet(iSet:integer);
    procedure setFinalTotal(sFinalTotal:string);
    function getItem(iNo : integer):TItem;
    procedure AddItem(iQuantity : integer;sDescription:string;rUnit, rTotal:real);
    procedure UndoItem;
    procedure SaveToFile(sPath, sInvNo, sOrderNo, sDeliveryNote, sDate : string);
    procedure Reset;
    //helper methods
    procedure ConvertPoint(var sComma:string);
    procedure ConvertComma(var sPoint:string);
  end;

implementation

{ TTheInvoice }

procedure TTheInvoice.AddItem(iQuantity: integer; sDescription: string;
  rUnit, rTotal: real);
var
  iAdd : integer;
  sT : string;
  rGlobal : real;
begin
  //add item to array
  inc(fSet);
  //max 26
  if fSet > 26 then
    begin
      ShowMessage('Item could not be added, invoice sheet item space has reached ' +
                  'its maximum!');
      Exit;
    end;
  //carry on
  SetLength(fArrItem, fSet);
  iAdd := fSet-1;
  fArrItem[iAdd] := TItem.create(iQuantity, sDescription, rUnit, rTotal);
  //change finaltotal
  sT := fFinalTotal;
  sT := copy(sT, 3, length(sT)-2);
  ConvertPoint(sT);

  rGlobal := strtofloat(sT);
  rGlobal := rGlobal + fArrItem[iAdd].getTotal;

  fFinalTotal := 'R ' + floattostrf(rGlobal, ffFixed, 5, 2);
  ConvertComma(fFinalTotal);
end;

procedure TTheInvoice.ConvertComma(var sPoint: string);
begin
  sPoint := StringReplace(sPoint, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
end;

procedure TTheInvoice.ConvertPoint(var sComma: string);
begin
   sComma := StringReplace(sComma, ',', '.',
              [rfReplaceAll, rfIgnoreCase]);
end;

constructor TTheInvoice.create;
begin
  //default
  fSet := 0;
  fFinalTotal := 'R 0.00';
end;

destructor TTheInvoice.destroy;
begin

  inherited;
end;

function TTheInvoice.getFinalTotal: string;
begin
  result := fFinalTotal;
end;

function TTheInvoice.getItem(iNo: integer): TItem;
begin
  result := fArrItem[iNo];
end;

function TTheInvoice.getSet: integer;
begin
  result := fSet;
end;

procedure TTheInvoice.Reset;
begin
  fSet := 0;
  fFinalTotal := 'R 0.00';
  SetLength(fArrItem, 0);
end;

procedure TTheInvoice.SaveToFile(sPath, sInvNo, sOrderNo, sDeliveryNote,
  sDate: string);
var
  //closefile
  XC : OLEVariant;
  //date
  sMonth : string;
  iMonth, iDay, iYear, iPos : integer;
  //saving
  ADOQuery : TADOQuery;
  ic1, iLoop : integer;
  sc1, sc2, sd1, sExcel, ssPath, sAPath, sQuery, sNow, sCUnit, sCTotal, sCheck : string;
begin
  //Date, year, month, day
  iPos := pos('/', sDate);
  iMonth := strtoint(copy(sDate, 1, iPos-1));
  Delete(sDate, 1, iPos);

  iPos := pos('/', sDate);
  iDay := strtoint(copy(sDate, 1, iPOs-1));
  Delete(sDate, 1, iPos);

  iYear := strtoint(sDate);

  case iMonth of
    1: sMonth := 'January';
    2: sMonth := 'February';
    3: sMonth := 'March';
    4: sMonth := 'April';
    5: sMonth := 'May';
    6: sMonth := 'June';
    7: sMonth := 'July';
    8: sMonth := 'August';
    9: sMonth := 'September';
    10: sMonth := 'October';
    11: sMonth := 'November';
    12: sMonth := 'December';
    end;

  //excel name
  sExcel := '\' + inttostr(iDay) + ' ' + sMonth + ' ' + inttostr(iyear) +
            ' Invoice ' + sInvNo + '.xlsx';
  //get paths ready
  ssPath := copy(sPath, 1, length(sPath) - 23);
  sAPath := copy(sPath, 1, length(sPath) - 23)+ 'Invoices\';
  sQuery := sAPath + sMonth + ' ' + inttostr(iYear) + sExcel;
  //check if folder exists for month and year
  if DirectoryExists(sMonth + ' ' + inttostr(iYear)) = false
    then CreateDir(sAPath + sMonth + ' ' + inttostr(iYear));
  CopyFile(PChar(ssPath + 'Formatted Invoice.xlsx'), pChar(sAPath + sMonth
           + ' ' + inttostr(iYear) + sExcel), true);
  //start connection
  ADOQuery := TADOQuery.Create(Application);
  ADOQuery.ConnectionString := '';
  ADOQuery.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' +
    sQuery + ';Extended Properties="Excel 12.0 Xml;HDR=yes";Persist Security Info=False';
  ADOQuery.ParamCheck := False;

  ic1 := 21;
  //write to file - items
  for iLoop := 0 to fSet-1 do
    begin
      inc(ic1);
      sc1 := 'B';
      sc2 := sc1 + inttostr(ic1); //B22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' +
              quotedstr(inttostr(fArrItem[iLoop].getQuantity)) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;

      sc1 := 'C';
      sc2 := sc1 + inttostr(ic1); //C22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' +
              quotedstr(fArrItem[iLoop].getDescription) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;

      //get unit and total into rand format R x.xx
      sCUnit := 'R ' + floattostrf(fArrItem[iLoop].getUnit, ffFixed, 5, 2);

      sCTotal := 'R ' + floattostrf(fArrItem[iLoop].getTotal, ffFixed, 5, 2);

      sCUnit := StringReplace(sCunit, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
      sCTotal := StringReplace(sCTotal, '.', ',',
              [rfReplaceAll, rfIgnoreCase]);
      //carry on
      sc1 := 'F';
      sc2 := sc1 + inttostr(ic1); //F22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' +
              quotedstr(sCUnit) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;

      sc1 := 'G';
      sc2 := sc1 + inttostr(ic1); //G22
      sd1 := 'INSERT INTO [Sheet1$' + sc2 + ':' + sc2 + '] VALUES(' +
              quotedstr(sCTotal) + ')';
      ADOQuery.SQL.Text := sd1;
      ADOQuery.ExecSQL;
    end;

  //other details of the invoice
  sNow := inttostr(iDay) + ' ' + sMonth + ' ' + inttostr(iYear);
  ADOQuery.SQL.Text := 'INSERT INTO [Sheet1$F9:F9] VALUES (' + quotedstr(sNow) + ')';
  ADOQuery.ExecSQL;

  ADOQuery.SQL.Text := 'INSERT INTO [Sheet1$G48:G48] VALUES (' +
                        quotedstr(fFinalTotal) + ')';
  ADOQuery.ExecSQL;

  ADOQuery.SQL.Text := 'INSERT INTO [Sheet1$G12:G12] VALUES (' +
                        quotedstr(sInvNo) + ')';
  ADOQuery.ExecSQL;

  ADOQuery.SQL.Text := 'INSERT INTO [Sheet1$G13:G13] VALUES (' +
                        quotedstr(sOrderNo) + ')';
  ADOQuery.ExecSQL;

  ADOQuery.SQL.Text := 'INSERT INTO [Sheet1$G14:G14] VALUES (' +
                        quotedstr(sDeliveryNote) + ')';
  ADOQuery.ExecSQL;

  ShowMessage('Invoice saved, you will find your file in the folder Invoices');
  ADOQuery.Close;
  ADOQuery.ConnectionString := '';
  //close it
  sCheck := sAPath + sMonth + ' ' + inttostr(iYear) + sExcel;
  XC := CreateOleObject('Excel.Application');
  XC.Workbooks.Open(sCheck);
  XC.Workbooks.Close;
  XC.Quit;
  XC := Unassigned;
end;

procedure TTheInvoice.setFinalTotal(sFinalTotal: string);
begin
  fFinalTotal := sFinalTotal;
end;

procedure TTheInvoice.setSet(iSet: integer);
begin
  fSet := iSet;
end;

procedure TTheInvoice.UndoItem;
var
  rGlobal, rMinus : real;
  sT : string;
begin
  //change setLength
  fSet := fSet-1;
  rMinus := fArrItem[fSet].getTotal;
  SetLength(fArrItem, fSet);
  //change finaltotal
  sT := fFinalTotal;
  sT := copy(sT, 3, length(sT)-2);
  ConvertPoint(sT);
  rGlobal := strtofloat(sT);
  rGlobal := rGlobal - rMinus;

  fFinalTotal := 'R ' + floattostrf(rGlobal, ffFixed, 5, 2);
  ConvertComma(fFinalTotal);
end;

end.
