unit clsItem;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TItem = class
  private
    fUnit : real;
    fQuantity : integer;
    fDescription : string;
    fTotal : real;
  public
    constructor create(iQuantity : integer; sDescr : string; rUnit, rTotal : real);
    destructor destroy; override;
    function getUnit : real;
    function getQuantity : integer;
    function getDescription : string;
    function getTotal : real;
    function toString : string;
    procedure setUnit(rUnit : real);
    procedure setQuantity(iQuantity : integer);
    procedure setDescription(sDescr : string);
    procedure setTotal(rTotal : real);
  end;

implementation

{ TItem }

constructor TItem.Create(iQuantity: integer; sDescr : string;
  rUnit, rTotal: real);
begin
  fQuantity := iQuantity;
  fUnit := rUnit;
  fDescription := sDescr;
  fTotal := rTotal;
end;

destructor TItem.Destroy;
begin

  inherited;
end;

function TItem.getDescription: string;
begin
  result := fDescription;
end;

function TItem.getQuantity: integer;
begin
  result := fQuantity;
end;

function TItem.getTotal: real;
begin
  result := fTotal;
end;

function TItem.getUnit: real;
begin
  result := fUnit;
end;

procedure TItem.setDescription(sDescr: string);
begin
  fDescription := sDescr;
end;

procedure TItem.setQuantity(iQuantity: integer);
begin
  fQuantity := iQuantity;
end;

procedure TItem.setTotal(rTotal: real);
begin
  fTotal := rTotal;
end;

procedure TItem.setUnit(rUnit: real);
begin
  fUnit := rUnit;
end;

function TItem.toString: string;
begin
  result := #9 + inttostr(fQuantity) + #9 + fDescription + #9 +
            floattostrf(fUnit, ffFixed, 5,2) + #9 + floattostrf(fTotal,ffFixed,5,2);
end;

end.
 