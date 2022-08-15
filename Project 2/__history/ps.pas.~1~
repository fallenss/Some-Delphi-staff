unit ps;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ShellApi, Vcl.StdCtrls, ComObj, ActiveX,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls;

type
  TForm2 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Image1: TImage;
    Timer1: TTimer;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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
  command,user: string;
   ExcelApp, ExcelSheet: OLEVariant;
   MyMass: Variant;
   x, y,i,j,i1,j1: Integer;
   CurrentDate, date1,date2: TDateTime;
   live,fek: Boolean;


implementation

{$R *.dfm}

procedure TakeData();
begin

           // создание OLE-объекта Excel
   ExcelApp := CreateOleObject('Excel.Application');

   // открытие книги Excel
   ExcelApp.Workbooks.Open('Q:\”правление безопасности\!ќбщие документы\∆урнал учета отпусков.xlsx');

   // открытие листа книги
   ExcelSheet := ExcelApp.Workbooks[1].WorkSheets[1];

   // выделение последней задействованной €чейки на листе
   ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

   // получение значений размера выбранного диапазона
   x := ExcelApp.ActiveCell.Row;
   y := ExcelApp.ActiveCell.Column;

   CurrentDate:=Date;
   // присвоение массиву диапазона €чеек на листе и вывод
   MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;
   command:='';
   for i := 1 to x do
   begin
   fek:=false;



    for j := 1 to y do
    begin
      case j of
         1: user:= MyMass[i,j];
         2: date1:= MyMass[i,j];
         3: date2:= MyMass[i,j];
         else break;
      end;
    end;
    if (user='') or (date1=0) or (date2=0) then continue;
   if i>1 then
     begin
       for i1 := 1 to i-1 do
         begin
          if (user=MyMass[i1,1]) and (CurrentDate>=MyMass[i1,2]) and (CurrentDate<=MyMass[i1,3]) then
           begin
           fek:=true;
           break;
           end;
         end;
          if fek=true then continue;
     end;

    if (CurrentDate>=date1) and (CurrentDate<=date2) then  command:=command+'Disable-ADAccount '+user
    else command:=command+'Enable-ADAccount '+user;


    command:=command+'; ';
   end;

   // закрытие книги и очистка переменных
   ExcelApp.Quit;
   ExcelApp := Unassigned;
   ExcelSheet := Unassigned;
end;

procedure WriteLog();
  var f: TextFile;
  begin
    AssignFile(f,'C:\Users\ssaprankov\Documents\Embarcadero\Studio\Projects\Win32\Debug\log.txt');

    Append(f);
    Writeln(f,DateTimeToStr(Now)+':');
    Writeln(f,command);
    Writeln(f,'');

    CloseFile(f);
end;

procedure TForm2.Button1Click(Sender: TObject);        //PS
begin
    ShellExecute(handle,'open', 'powershell.exe',pchar('Import-Module ActiveDirectory'+ #10#13+command) ,nil, Sw_ShowNormal);
end;



procedure TForm2.Button2Click(Sender: TObject);        //Excel
begin
    live:=true;
    TakeData();
    label1.Caption:=command;
    WriteLog();
end;



procedure TForm2.FormClick(Sender: TObject);
begin
      live:= true;
end;



procedure TForm2.FormCreate(Sender: TObject);     //√лавное окно
 var  rgn: HRGN;
 begin
    Form2.Borderstyle := bsNone;
     rgn := CreateRoundRectRgn(0,0,ClientWidth,ClientHeight,40, 40);
   SetWindowRgn(Handle, rgn, True);

     TakeData();
     Form2.Button1.Click;
     WriteLog();
end;



procedure TForm2.Image1Click(Sender: TObject);
begin
close;
end;

procedure TForm2.Timer1Timer(Sender: TObject);
begin
if live=true then Timer1.Interval:=9000000
else close;
end;

end.
