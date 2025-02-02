unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ShellApi, Vcl.StdCtrls, ComObj,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls;

type
  TForm2 = class(TForm)
    Image1: TImage;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Timer1: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
const
   xlCellTypeLastCell = $0000000B;
var
  Form2: TForm2;
  Outlook, MailItem: Variant;
  user,date,host: string;
   ExcelApp, ExcelSheet: OLEVariant;
   MyMass: Variant;
   x, y,i,j: Integer;
   live: Boolean;

implementation

{$R *.dfm}

procedure TakeData();
begin


end;

procedure TForm2.Button2Click(Sender: TObject);
const
  olMailItem = 0;
begin

var i: integer;
Label1.Caption:='';
       // �������� OLE-������� Excel
   ExcelApp := CreateOleObject('Excel.Application');

   // �������� ����� Excel
   ExcelApp.Workbooks.Open('C:\Users\ssaprankov\Desktop\��������\1��������.xlsx');

   // �������� ����� �����
   ExcelSheet := ExcelApp.Workbooks[1].WorkSheets[1];

   // ��������� ��������� ��������������� ������ �� �����
   ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

   // ��������� �������� ������� ���������� ���������
   x := ExcelApp.ActiveCell.Row;
   y := ExcelApp.ActiveCell.Column;


   // ���������� ������� ��������� ����� �� ����� � �����
   MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

   for i := 1 to x do
   begin
   for j := 1 to y do
      begin
      case j of                   //creating variables by data from file
       1: user:= MyMass[i,j];
       2: date:= MyMass[i,j];
       3: host:= MyMass[i,j];
       else break;
      end;
     end;
   if (user='') or (date='') or (host='') then continue     //cheking for empty data
   else
    begin
      try                                        //creating outlook window
        Outlook := GetActiveOleObject('Outlook.Application');
      except
        Outlook := CreateOleObject('Outlook.Application');
      end;

      MailItem := Outlook.CreateItem(olMailItem);            // creating mail
      MailItem.Recipients.Add(host);
      MailItem.Subject := '������������ �������� �����������';
      MailItem.Body := '������ ����.'+#10#13+'��������� ����������� �������� ����������� ����� �� �������� �������������� ������������, �������� ������� ������������ ��������� ��� ������������ �� 06.08.2020 �1175.'+#10#13+'����������� ������ �������� �'+date+'.'+#10#13+'��������� ����������� ����� ��������� ��������: '+user+'.'+#10#13+'�����, �������� �� ��� ������ ���� �� �����, ����, � ������ ���� ���-�� �� ������������� ����������� �� ������ ������ �������� �� �����-���� �������, ��������� ����, � � �������� ��� ��������.'+#10#13+'�������.'
;
      MailItem.Display;
      Label1.Caption:=Label1.Caption+ MailItem.Subject+' '+MailItem.Body;
      Button1.click;
     end;
   end;

   // �������� ����� � ������� ����������
   ExcelApp.Quit;
   ExcelApp := Unassigned;
   ExcelSheet := Unassigned;
end;

procedure TForm2.Button1Click(Sender: TObject);
begin

  MailItem.Send;
  Outlook := Unassigned;
end;

procedure TForm2.FormClick(Sender: TObject);
begin
live:=true;
end;

procedure TForm2.FormCreate(Sender: TObject);     // form
 var  rgn: HRGN;
 begin
    Form2.Borderstyle := bsNone;
     rgn := CreateRoundRectRgn(0,0,ClientWidth,ClientHeight,40, 40);
   SetWindowRgn(Handle, rgn, True);
   button2.Click;
end;

procedure TForm2.Image1Click(Sender: TObject);
begin
close;
end;



procedure TForm2.Timer1Timer(Sender: TObject);     // auto-closing program
begin
if live=true then Timer1.Interval:=9000000
else close;
end;

end.
