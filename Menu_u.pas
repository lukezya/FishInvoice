unit Menu_u;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, jpeg;

type
  TForm1 = class(TForm)
    shpHeader: TShape;
    imgLogo: TImage;
    shpMenu: TShape;
    btnHelp: TButton;
    btnExit: TButton;
    btnInvoice: TButton;
    btnStatement: TButton;
    procedure btnInvoiceClick(Sender: TObject);
    procedure btnStatementClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Invoice_u, Statement_u;

{$R *.dfm}

procedure TForm1.btnInvoiceClick(Sender: TObject);
begin
  //Invoice Form
  Form1.Hide;
  Form2.Show;
end;

procedure TForm1.btnStatementClick(Sender: TObject);
begin
  //Statement Form
  Form1.Hide;
  Form3.Show;
end;

procedure TForm1.btnExitClick(Sender: TObject);
begin
  //Exit
  Application.Terminate;
end;

end.
