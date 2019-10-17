program TheFishInvoice;

uses
  Forms,
  Menu_u in 'Menu_u.pas' {Form1},
  Invoice_u in 'Invoice_u.pas' {Form2},
  Statement_u in 'Statement_u.pas' {Form3},
  clsItem in 'clsItem.pas',
  clsTheInvoice in 'clsTheInvoice.pas',
  clsTheStatement in 'clsTheStatement.pas',
  clsInvoice in 'clsInvoice.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
