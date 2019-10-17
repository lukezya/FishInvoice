unit clsInvoice;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TInvoice = class
  private
    fInvoice : string;
    fPrice : string;
  public
    constructor create(sInvoice, sPrice : string);
    destructor destroy; override;
    function getInvoice : string;
    function getPrice : string;
    function toString : string;
    procedure setInvoice(sInvoice : string);
    procedure setPrice(sPrice : string);
  end;

implementation

{ TInvoice }

constructor TInvoice.create(sInvoice, sPrice: string);
begin
  fInvoice := sInvoice;
  fPrice := sPrice;
end;

destructor TInvoice.destroy;
begin

  inherited;
end;

function TInvoice.getInvoice: string;
begin
  result := fInvoice;
end;

function TInvoice.getPrice: string;
begin
  result := fPrice;
end;

procedure TInvoice.setInvoice(sInvoice: string);
begin
  fInvoice := sInvoice;
end;

procedure TInvoice.setPrice(sPrice: string);
begin
  fPrice := sPrice;
end;

function TInvoice.toString: string;
begin
  result := fInvoice + #9 + fPrice;
end;

end.
 