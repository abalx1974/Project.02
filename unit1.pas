unit Unit1;

{$mode objfpc}{$H+}



interface

uses
  Classes, SysUtils, IBConnection, sqldb, db, fpstdexports, FileUtil, LR_DBSet,
  LR_E_HTM, LR_Class, LR_View, lrAddFunctionLibrary, Forms, Controls, Graphics,
  Dialogs, Menus, ComCtrls, StdCtrls, DBGrids, fpsexport, INIFiles,
  lconvencoding, unit2;

type

  { TForm1 }

  TForm1 = class(TForm)
    DataSource1: TDataSource;
    DataSource2: TDataSource;
    DBGrid1: TDBGrid;
    FPSExport1: TFPSExport;
    frDBDataSet1: TfrDBDataSet;
    frHTMExport1: TfrHTMExport;
    frPreview1: TfrPreview;
    frReport1: TfrReport;
    IBConnection1: TIBConnection;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItem43: TMenuItem;
    MenuItem44: TMenuItem;
    MenuItem45: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem47: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    MenuItem52: TMenuItem;
    MenuItem53: TMenuItem;
    MenuItem54: TMenuItem;
    MenuItem55: TMenuItem;
    MenuItem56: TMenuItem;
    MenuItem57: TMenuItem;
    MenuItem58: TMenuItem;
    MenuItem59: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem60: TMenuItem;
    MenuItem61: TMenuItem;
    MenuItem62: TMenuItem;
    MenuItem63: TMenuItem;
    MenuItem64: TMenuItem;
    MenuItem65: TMenuItem;
    MenuItem66: TMenuItem;
    MenuItem67: TMenuItem;
    MenuItem68: TMenuItem;
    MenuItem69: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem70: TMenuItem;
    MenuItem71: TMenuItem;
    MenuItem72: TMenuItem;
    MenuItem73: TMenuItem;
    MenuItem74: TMenuItem;
    MenuItem75: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem80: TMenuItem;
    MenuItem81: TMenuItem;
    MenuItem82: TMenuItem;
    MenuItem83: TMenuItem;
    MenuItem84: TMenuItem;
    MenuItem9: TMenuItem;
    SQLQuery1: TSQLQuery;
    SQLQuery2: TSQLQuery;
    SQLTransaction1: TSQLTransaction;
    procedure FormCreate(Sender: TObject);
    procedure frPreview1Click(Sender: TObject);
    procedure IBConnection1AfterConnect(Sender: TObject);
    procedure Label1Click(Sender: TObject);
    procedure MenuItem14Click(Sender: TObject);
    procedure MenuItem15Click(Sender: TObject);
    procedure MenuItem16Click(Sender: TObject);
    procedure MenuItem17Click(Sender: TObject);
    procedure MenuItem18Click(Sender: TObject);
    procedure MenuItem19Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem20Click(Sender: TObject);
    procedure MenuItem21Click(Sender: TObject);
    procedure MenuItem22Click(Sender: TObject);
    procedure MenuItem23Click(Sender: TObject);
    procedure MenuItem24Click(Sender: TObject);
    procedure MenuItem25Click(Sender: TObject);
    procedure MenuItem26Click(Sender: TObject);
    procedure MenuItem27Click(Sender: TObject);
    procedure MenuItem28Click(Sender: TObject);
    procedure MenuItem29Click(Sender: TObject);
    procedure MenuItem30Click(Sender: TObject);
    procedure MenuItem31Click(Sender: TObject);
    procedure MenuItem32Click(Sender: TObject);
    procedure MenuItem33Click(Sender: TObject);
    procedure MenuItem34Click(Sender: TObject);
    procedure MenuItem35Click(Sender: TObject);
    procedure MenuItem36Click(Sender: TObject);
    procedure MenuItem37Click(Sender: TObject);
    procedure MenuItem39Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem40Click(Sender: TObject);
    procedure MenuItem42Click(Sender: TObject);
    procedure MenuItem43Click(Sender: TObject);
    procedure MenuItem44Click(Sender: TObject);
    procedure MenuItem45Click(Sender: TObject);
    procedure MenuItem47Click(Sender: TObject);
    procedure MenuItem48Click(Sender: TObject);
    procedure MenuItem49Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem46Click(Sender: TObject);
    procedure MenuItem50Click(Sender: TObject);
    procedure MenuItem51Click(Sender: TObject);
    procedure MenuItem52Click(Sender: TObject);
    procedure MenuItem53Click(Sender: TObject);
    procedure MenuItem54Click(Sender: TObject);
    procedure MenuItem55Click(Sender: TObject);
    procedure MenuItem56Click(Sender: TObject);
    procedure MenuItem57Click(Sender: TObject);
    procedure MenuItem58Click(Sender: TObject);
    procedure MenuItem59Click(Sender: TObject);
    procedure MenuItem60Click(Sender: TObject);
    procedure MenuItem61Click(Sender: TObject);
    procedure MenuItem62Click(Sender: TObject);
    procedure MenuItem63Click(Sender: TObject);
    procedure MenuItem64Click(Sender: TObject);
    procedure MenuItem65Click(Sender: TObject);
    procedure MenuItem66Click(Sender: TObject);
    procedure MenuItem67Click(Sender: TObject);
    procedure MenuItem68Click(Sender: TObject);
    procedure MenuItem69Click(Sender: TObject);
    procedure MenuItem70Click(Sender: TObject);
    procedure MenuItem71Click(Sender: TObject);
    procedure MenuItem72Click(Sender: TObject);
    procedure MenuItem73Click(Sender: TObject);
    procedure MenuItem74Click(Sender: TObject);
    procedure MenuItem75Click(Sender: TObject);
    procedure MenuItem76Click(Sender: TObject);
    procedure MenuItem77Click(Sender: TObject);
    procedure MenuItem78Click(Sender: TObject);
    procedure MenuItem79Click(Sender: TObject);
    procedure MenuItem80Click(Sender: TObject);
    procedure MenuItem81Click(Sender: TObject);
    procedure MenuItem82Click(Sender: TObject);
    procedure MenuItem83Click(Sender: TObject);
    procedure MenuItem84Click(Sender: TObject);
  private
    { private declarations }
    procedure AVDObject();
    procedure AllObject();
    procedure ObslugaObject(name1:string);
    procedure AllObjectXozorg();
  public
    { public declarations }
    Obsluga:      string;
    Osoba:        string;
    filename:     string;
    zaprdate:     string;


  end;




var
  Form1: TForm1;


implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.AVDObject();
begin
    SQLQuery1.Close;
    SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1 FROM OBJECTS a WHERE POSITION(:obslug,a.LOCATION1) > 0 AND POSITION(:otkl,a.OBJSHORTNAME1) = 0 AND character_length(GSMPHONE)=0 AND OBJN > 999 ORDER BY a.OBJN';
    SQLQuery1.ParamByName('otkl').AsString:=Form1.Obsluga;
    SQLQuery1.ParamByName('obslug').AsString:=Form1.Osoba;
    SQLQuery1.Open;

    frReport1.LoadFromFile('proba3.lrf');
    frReport1.ShowReport;
    if frReport1.PrepareReport then
      frReport1.ExportTo(TfrHTMExportFilter, 'avtodozvon.html');
    IBConnection1.Close;
    IBConnection1.Open;
end;

procedure TForm1.AllObjectXozorg();
var
    Exp: TFPSExport;
    ExpSettings: TFPSExportFormatSettings;
    TheDate: TDateTime;

begin

    SQLQuery2.Close;

    SQLQuery2.sql.text:='SELECT a.OBJN,a.OBJFULLNAME1, a.OBJSHORTNAME1,a.ADDRESS1, a.PHONES1,b.SURNAME1, b.NAME1, b.SECNAME1, b.ADDRESS1, b.PHONES1, b.STATUS1, a.CONTRACT1, a.LOCATION1, a.NOTES1, a.GSMPHONE FROM OBJECTS a,PERSONAL b WHERE a.OBJUIN = b.OBJUIN AND POSITION(:obslug,a.LOCATION1) > 0 AND POSITION(:otkl,a.OBJSHORTNAME1) = 0 AND a.OBJN > 999 ORDER BY a.OBJN';

    SQLQuery2.ParamByName('otkl').AsString:=Form1.Obsluga;
    SQLQuery2.ParamByName('obslug').AsString:=Form1.Osoba;
    SQLQuery2.Open;


    Exp := TFPSExport.Create(nil);
    ExpSettings := TFPSExportFormatSettings.Create(true);

    ExpSettings.ExportFormat := efXLS; // choose file format
    ExpSettings.HeaderRow := true; // include header row with field names
    Exp.FormatSettings := ExpSettings; // apply settings to export object
    Exp.Dataset:=SQLQuery2;
    Exp.FileName := Form1.filename;
    Exp.Execute; // run the export

    Exp.Free;
    ExpSettings.Free;


end;

procedure TForm1.AllObject();
begin
    SQLQuery1.Close;
    SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1 FROM OBJECTS a WHERE POSITION(:obslug,a.LOCATION1) > 0 AND POSITION(:otkl,a.OBJSHORTNAME1) = 0 AND OBJN > 999 ORDER BY a.OBJN';
    SQLQuery1.ParamByName('otkl').AsString:=Form1.Obsluga;
    SQLQuery1.ParamByName('obslug').AsString:=Form1.Osoba;
    SQLQuery1.Open;

    frReport1.LoadFromFile('proba4.lrf');
    frReport1.ShowReport;
    if frReport1.PrepareReport then
      frReport1.ExportTo(TfrHTMExportFilter,Form1.filename);
    IBConnection1.Close;
    IBConnection1.Open;

end;
procedure TForm1.ObslugaObject(name1:string);
begin
    SQLQuery1.Close;
    SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1,a.NOTES1,a.GSMPHONE, a.NEXTTESTTIME1, a.LASTTESTTIME1 FROM OBJECTS a WHERE POSITION(:obslug,a.NOTES1) > 0 AND POSITION(:otkl,a.OBJSHORTNAME1) = 0 AND OBJN > 999 ORDER BY a.OBJN';
    SQLQuery1.ParamByName('otkl').AsString:=Form1.Obsluga;
    SQLQuery1.ParamByName('obslug').AsString:=Form1.Osoba;
    SQLQuery1.Open;
    with frReport1 do begin
           LoadFromFile('proba7.lrf');
           Variables.Add(' Zagol');
           frVariables['Zagol']:=name1;
           ShowReport;
    end;

    if frReport1.PrepareReport then
      frReport1.ExportTo(TfrHTMExportFilter,Form1.filename);
    IBConnection1.Close;
    IBConnection1.Open;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin

end;

procedure TForm1.MenuItem20Click(Sender: TObject);
{Объекты автодозвон Віт}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Віт';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem21Click(Sender: TObject);
{Объекты автодозвон Гудима}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гудима';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem22Click(Sender: TObject);
{Объекты автодозвон Лозінський}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Лозінський';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem23Click(Sender: TObject);
{Объекты автодозвон Федорчук}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Федорчук';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem24Click(Sender: TObject);
{Объекты автодозвон Кармаліта}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кармаліта';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem25Click(Sender: TObject);
{Объекты автодозвон Бабій}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Бабій';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem26Click(Sender: TObject);
{Объекты автодозвон Доля}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Доля';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem27Click(Sender: TObject);
{Объекты автодозвон Квасовський}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Квасовський';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem28Click(Sender: TObject);
{Федорчук всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Федорчук';
  Form1.filename:='fedorchuk.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem29Click(Sender: TObject);
{Кармаліта всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кармаліта';
  Form1.filename:='karmalita.html';
  Form1.AllObject();

end;

procedure TForm1.MenuItem14Click(Sender: TObject);
begin
  Halt(0);
end;

procedure TForm1.MenuItem15Click(Sender: TObject);
{Действующие объекты GSM/GPRS}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  {Form1.obsluga:='РЕМОНТ';}
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) = 0 AND character_length(a.GSMPHONE) > 0 ORDER BY a.OBJN';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;
  SQLQuery1.Open;

  frReport1.LoadFromFile('proba2.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'gsmobjects.html');
  IBConnection1.Close;
  IBConnection1.Open;
end;

procedure TForm1.MenuItem16Click(Sender: TObject);
{Всі Діючі об`екти автодозвон}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem17Click(Sender: TObject);
{Всі відключені об`екти}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.zaprdate:='now';
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) > 0 ORDER BY a.LASTTESTTIME1 DESC';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;
  {SQLQuery1.ParamByName('zaprdate').AsString:=Form1.zaprdate;}
  SQLQuery1.Open;

  frReport1.LoadFromFile('proba5.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'otkl_total.html');
  IBConnection1.Close;
  IBConnection1.Open;
end;

procedure TForm1.MenuItem18Click(Sender: TObject);
{Объекты автодозвон Кушнир}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кушнір';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem19Click(Sender: TObject);
{Объекты автодозвон Гуцалюк}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гуцалюк';
  Form1.AVDObject();
end;

procedure TForm1.IBConnection1AfterConnect(Sender: TObject);
begin

end;

procedure TForm1.Label1Click(Sender: TObject);
begin

end;

procedure TForm1.frPreview1Click(Sender: TObject);
begin

end;

procedure TForm1.FormCreate(Sender: TObject);
var
    IniF : TINIFile;
    host_address : string;
    database_name : string;

begin
     if FileExists(ExtractFilePath(ParamStr(0))+'project02.ini') then begin
       Inif :=TINIFile.Create(ExtractFilePath(ParamStr(0))+'project02.ini');
       host_address:=Inif.ReadString('most_host','host_address','localhost');
       database_name:=Inif.ReadString('most_host','database_name','most5');
     end else begin
       Inif := TINIFile.Create(ExtractFilePath(ParamStr(0))+'project02.ini');
       Inif.WriteString('most_host','host_address','10.10.11.2');
       host_address:='10.10.11.12';
       Inif.WriteString('most_host','database_name','most5');
       database_name:='most5';

     end;
       Inif.Free;

       IBConnection1.DatabaseName:=database_name;
       IBConnection1.HostName:=host_address;
       IBConnection1.Connected:=false;
       Label1.Caption:=host_address;
       Label1.Visible:=true;
       Label2.Caption:=database_name;
       Label2.Visible:=true;

end;

procedure TForm1.MenuItem46Click(Sender: TObject);
{Техн. обслуг. ТОВ *Пожежне спостереження-Хмельницький*  *  706519}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='ТОВ *Пожежне спостереження-Хмельницький*';
  Form1.filename:='to_pshm.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Пожежне спостереження-Хмельницький*  *  706519');
end;

procedure TForm1.MenuItem50Click(Sender: TObject);
{Техн. обслуг. Аварійно-рятувальний загін спец. призначення МНС України в Хмельн.  обл.  *   63-01-03}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Аварійно-рятувальний';
  Form1.filename:='to_mns.html';
  Form1.ObslugaObject('Техн. обслуг. Аварійно-рятувальний загін спец. призначення МНС України в Хмельн.  обл.  *   63-01-03');
end;

procedure TForm1.MenuItem51Click(Sender: TObject);
{Техн. обслуг. ПП *Ботеза*  *  74-31-20}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Ботеза';
  Form1.filename:='to_boteza.html';
  Form1.ObslugaObject('Техн. обслуг. ПП *Ботеза*  *  74-31-20');
end;

procedure TForm1.MenuItem52Click(Sender: TObject);
{Техн. обслуг. ПБП *СПЕЦТЕХНІКА ГРУП*   *  65-98-13}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='СПЕЦТЕХНІКА';
  Form1.filename:='to_sptehgrp.html';
  Form1.ObslugaObject('Техн. обслуг. ПБП *СПЕЦТЕХНІКА ГРУП*   *  65-98-13');
end;

procedure TForm1.MenuItem53Click(Sender: TObject);
{Техн. обслуг. ТОВ *Хмельницькпожспецсервіс*}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Хмельницькпожспецсервіс';
  Form1.filename:='to_hmpozspezsrv.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Хмельницькпожспецсервіс*');
end;

procedure TForm1.MenuItem54Click(Sender: TObject);
{Техн. обслуг. ДП „Центр „ Інновації та технології*  *  72-01-04, 72-01-05}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Інновації та технології';
  Form1.filename:='to_innovacii.html';
  Form1.ObslugaObject('Техн. обслуг. ДП „Центр „ Інновації та технології*  *  72-01-04, 72-01-05');
end;

procedure TForm1.MenuItem55Click(Sender: TObject);
{Техн. обслуг. ТОВ *Агробудсервіс*  *  72-05-69, 79-50-66}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Агробудсервіс';
  Form1.filename:='to_agrobud.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Агробудсервіс*  *  72-05-69, 79-50-66');
end;

procedure TForm1.MenuItem56Click(Sender: TObject);
{Техн. обслуг. БМ ТОВ *Поділля*  *   65-72-03}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Поділля';
  Form1.filename:='to_bmpodillya.html';
  Form1.ObslugaObject('Техн. обслуг. БМ ТОВ *Поділля*  *   65-72-03');

end;

procedure TForm1.MenuItem57Click(Sender: TObject);
{Тех.обсл. ПП Омега-Центр-70-44-04,067-164-10-32 Геннадий Петрович}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Омега';
  Form1.filename:='to_omega.html';
  Form1.ObslugaObject('Тех.обсл. ПП Омега-Центр-70-44-04,067-164-10-32 Геннадий Петрович');

end;

procedure TForm1.MenuItem58Click(Sender: TObject);
{Техн. обслуг. ДСО УМВС  *  72-56-92}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='УМВС';
  Form1.filename:='to_dsoumvs.html';
  Form1.ObslugaObject('Техн. обслуг. ДСО УМВС  *  72-56-92');

end;

procedure TForm1.MenuItem59Click(Sender: TObject);
{Техн. обслуг. не проводиться}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='не проводиться';
  Form1.filename:='to_net.html';
  Form1.ObslugaObject('Техн. обслуг. не проводиться');
end;

procedure TForm1.MenuItem60Click(Sender: TObject);
{Техн. обслуг. ТОВ *Інтервідеосервіс*вул. Прибузька, 4  *  65-53-98}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Інтервідео';
  Form1.filename:='to_intervid.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Інтервідеосервіс*вул. Прибузька, 4  *  65-53-98');
end;

procedure TForm1.MenuItem61Click(Sender: TObject);
{Техн. обслуг. Приватне підприємство „Стражник*  *  66-50-50}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Стражник';
  Form1.filename:='to_stragnik.html';
  Form1.ObslugaObject('Техн. обслуг. Приватне підприємство „Стражник*  *  66-50-50');

end;

procedure TForm1.MenuItem62Click(Sender: TObject);
{Техн. обслуг. Хмельницька дирекція ВАТ *УКРТЕЛЕКОМ*  *  72-07-04- приймальна 79-34-72, 79-32-82}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='УКРТЕЛЕКОМ';
  Form1.filename:='to_ukrtelecom.html';
  Form1.ObslugaObject('Техн. обслуг. Хмельницька дирекція ВАТ *УКРТЕЛЕКОМ*  *  72-07-04- приймальна 79-34-72, 79-32-82');

end;

procedure TForm1.MenuItem63Click(Sender: TObject);
{Техн. обслуг. Шепетівська Дільниця протипожежних робіт  * 03840- 5-88-23  *}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Шепетівськ';
  Form1.filename:='to_shepdiln.html';
  Form1.ObslugaObject('Техн. обслуг. Шепетівська Дільниця протипожежних робіт  * 03840- 5-88-23  *');
end;

procedure TForm1.MenuItem64Click(Sender: TObject);
{Техн. обслуг. Лабораторія МТОЗ Грінник Сергій Павлович 097-299-0766}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Грінник';
  Form1.filename:='to_hnu.html';
  Form1.ObslugaObject('Техн. обслуг. Лабораторія МТОЗ Грінник Сергій Павлович 097-299-0766');
end;

procedure TForm1.MenuItem65Click(Sender: TObject);
{Техн. обслуг. ТОВ *Сигнал* (м.Кам-Под.)  Сідлецький Петро Михайлович *  067-600-16-16}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Сигнал';
  Form1.filename:='to_signal.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Сигнал* (м.Кам-Под.)  Сідлецький Петро Михайлович *  067-600-16-16');
end;

procedure TForm1.MenuItem66Click(Sender: TObject);
{Техн. обслуг. Хм.обл.спец.рем-буд.підпр.протипож.робіт  *  55-04-03, 55-18-43 }
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='55-18-43';
  Form1.filename:='to_hmspecrembud.html';
  Form1.ObslugaObject('Техн. обслуг. Хм.обл.спец.рем-буд.підпр.протипож.робіт  *  55-04-03, 55-18-43');
end;

procedure TForm1.MenuItem67Click(Sender: TObject);
{Техн. обслуг. ПП *Хмельницьк-Антисептік-БМУ*}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Антисептік';
  Form1.filename:='to_antiseptik.html';
  Form1.ObslugaObject('Техн. обслуг. ПП *Хмельницьк-Антисептік-БМУ*');
end;

procedure TForm1.MenuItem68Click(Sender: TObject);
{Техн. обслуг.  Хмельн. професійний ліцей електроніки  *  67-17-39}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='ліцей електроніки';
  Form1.filename:='to_liceyelectr.html';
  Form1.ObslugaObject('Техн. обслуг.  Хмельн. професійний ліцей електроніки  *  67-17-39');
end;

procedure TForm1.MenuItem69Click(Sender: TObject);
{Техн. обслуг.  ДП *Хмельницьке БВП ОСС*  *  55-13-73}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='БВП ОСС';
  Form1.filename:='to_hmbvp_occ.html';
  Form1.ObslugaObject('Техн. обслуг.  ДП *Хмельницьке БВП ОСС*  *  55-13-73');
end;

procedure TForm1.MenuItem70Click(Sender: TObject);
{Техн. обслуг. ТОВ *Науково-виробнича фірма Бранд майстер плюс*}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Бранд майстер';
  Form1.filename:='to_brand.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ *Науково-виробнича фірма Бранд майстер плюс*');
end;

procedure TForm1.MenuItem71Click(Sender: TObject);
{Техн. обслуг. *Професійні системи безпеки м. Львів*}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Львів';
  Form1.filename:='to_lviv.html';
  Form1.ObslugaObject('Техн. обслуг. *Професійні системи безпеки м. Львів*');
end;

procedure TForm1.MenuItem72Click(Sender: TObject);
{ТО вони самі роблять}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='ТО вони самі роблять';
  Form1.filename:='to_sami.html';
  Form1.ObslugaObject('ТО вони самі роблять');
end;

procedure TForm1.MenuItem73Click(Sender: TObject);
{Пошук по ТО}
var
    str2 : string;
begin
  Form2.Show;
  str2:=trim(Form2.str2);
  if length(str2)>0 then begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:=str2;
  Form1.filename:='to_poisk.html';
  Form1.ObslugaObject(str2);
  end;

end;

procedure TForm1.MenuItem74Click(Sender: TObject);
{ Об'екти на яких встановлені наші плати або прилади з підмінного фонда }
begin
   Form1.obsluga:='ОТКЛЮЧИТЬ';
  {Form1.obsluga:='РЕМОНТ';}
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) = 0 AND POSITION(:ownstring,a.LOCATION1)>0 ORDER BY a.OBJN';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;
  SQLQuery1.ParamByName('ownstring').AsString:=' наш';
  SQLQuery1.Open;

  frReport1.LoadFromFile('proba8.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'obmenfond.html');
  IBConnection1.Close;
  IBConnection1.Open;

end;

procedure TForm1.MenuItem75Click(Sender: TObject);
{ Об`екти без зв`язку }
begin
      Form1.obsluga:='ОТКЛЮЧИТЬ';
  {Form1.obsluga:='РЕМОНТ';}
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) = 0 AND a.ISCONNSTATE1=0 AND a.OBJN >100 ORDER BY a.OBJN';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;

  SQLQuery1.Open;

  frReport1.LoadFromFile('proba9.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'bezsvjazi.html');
  IBConnection1.Close;
  IBConnection1.Open;

end;

procedure TForm1.MenuItem76Click(Sender: TObject);
{ Об`екти в отладке }
begin
   Form1.Obsluga:='ОТКЛЮЧИТЬ';
   SQLQuery1.Close;
   SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) = 0 AND a.ENG1=1 AND a.OBJN >100 ORDER BY a.OBJN';
   SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;

   SQLQuery1.Open;

   frReport1.LoadFromFile('proba10.lrf');
    frReport1.ShowReport;
   if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'votladke.html');
   IBConnection1.Close;
   IBConnection1.Open;

end;

procedure TForm1.MenuItem77Click(Sender: TObject);
{ Доля Объекты полный список с хозорганами }
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Доля';
  Form1.filename:='dolya_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem78Click(Sender: TObject);
{Бабій всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Бабій';
  Form1.filename:='babij_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem79Click(Sender: TObject);
{Гуцалюк всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гуцалюк';
  Form1.filename:='guzaluk_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem80Click(Sender: TObject);
{Віт всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Віт';
  Form1.filename:='vit_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem81Click(Sender: TObject);
{Кармаліта всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кармаліта';
  Form1.filename:='karmalita_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem82Click(Sender: TObject);
{Гудима всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гудима';
  Form1.filename:='gudima_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem83Click(Sender: TObject);
{Лозінський всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Лозінський';
  Form1.filename:='lozinski_xozorg.xls';
  Form1.AllObjectXozorg();
end;

procedure TForm1.MenuItem84Click(Sender: TObject);
{Кушнір всі об`єкти з хозорганами}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кушнір';
  Form1.filename:='kushnir_xozorg.xls';
  Form1.AllObjectXozorg();
end;


procedure TForm1.MenuItem30Click(Sender: TObject);
{Бабій всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Бабій';
  Form1.filename:='babij.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem31Click(Sender: TObject);
{Доля всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Доля';
  Form1.filename:='dolya.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem32Click(Sender: TObject);
{Лозинський всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Лозінський';
  Form1.filename:='lozinski.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem33Click(Sender: TObject);
{Гудима всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гудима';
  Form1.filename:='gudima.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem34Click(Sender: TObject);
{Віт всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Віт';
  Form1.filename:='gudima.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem35Click(Sender: TObject);
{Гуцалюк всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Гуцалюк';
  Form1.filename:='guzaluk.html';
  Form1.AllObject();

end;

procedure TForm1.MenuItem36Click(Sender: TObject);
{Кушнір всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Кушнір';
  Form1.filename:='kushir.html';
  Form1.AllObject();

end;

procedure TForm1.MenuItem37Click(Sender: TObject);
{Квасовський всі об`єкти}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Квасовський';
  Form1.filename:='kvasovski.html';
  Form1.AllObject();

end;

procedure TForm1.MenuItem39Click(Sender: TObject);
{Михайлиця автодозвон}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Михайлица';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem3Click(Sender: TObject);
{Отключенные объекты оставшиеся на связи}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.zaprdate:='now';
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT FIRST 20 a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) > 0 ORDER BY a.LASTTESTTIME1 DESC';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;
  {SQLQuery1.ParamByName('zaprdate').AsString:=Form1.zaprdate;}
  SQLQuery1.Open;

  frReport1.LoadFromFile('proba5.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'otkl_active.html');
  IBConnection1.Close;
  IBConnection1.Open;
end;

procedure TForm1.MenuItem40Click(Sender: TObject);
 { Михайлица повний перелік }
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Михайлица';
  Form1.filename:='mihailiza.html';
  Form1.AllObject();

end;

procedure TForm1.MenuItem42Click(Sender: TObject);
 {Нижний автодозвон}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Нижний';
  Form1.AVDObject();
end;

procedure TForm1.MenuItem43Click(Sender: TObject);
 {Нижний повний перелік}
begin
  Form1.Obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Нижний';
  Form1.filename:='nizniy.html';
  Form1.AllObject();
end;

procedure TForm1.MenuItem44Click(Sender: TObject);
 {Діючі об`єкти МТС}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  {Form1.obsluga:='РЕМОНТ';}
  SQLQuery1.Close;
  SQLQuery1.sql.text:='SELECT a.OBJN, a.OBJFULLNAME1, a.OBJSHORTNAME1, a.OBJTYPEID, a.ADDRESS1, a.CONTRACT1, a.LOCATION1, a.NOTES1,  a.NEXTTESTTIME1, a.LASTTESTTIME1, a.GSMPHONE FROM OBJECTS a WHERE POSITION(:obslug,a.OBJSHORTNAME1) = 0 AND POSITION(:gsmoperator,a.GSMPHONE) >0 AND character_length(a.GSMPHONE) > 0 ORDER BY a.OBJN';
  SQLQuery1.ParamByName('obslug').AsString:=Form1.obsluga;
  SQLQuery1.ParamByName('gsmoperator').AsString:='050';
  SQLQuery1.Open;

  frReport1.LoadFromFile('proba6.lrf');
  frReport1.ShowReport;
  if frReport1.PrepareReport then
    frReport1.ExportTo(TfrHTMExportFilter, 'mtcgsmobjects.html');
  IBConnection1.Close;
  IBConnection1.Open;
end;

procedure TForm1.MenuItem45Click(Sender: TObject);
begin

end;

procedure TForm1.MenuItem47Click(Sender: TObject);
{Техн. обслуг. ПП *ХМЕЛЬНИЦКПОЖСЕРВІС -Дяченко*}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='ХМЕЛЬНИЦКПОЖСЕРВІС';
  Form1.filename:='to_hmpzserv.html';
  Form1.ObslugaObject('Техн. обслуг. ПП *ХМЕЛЬНИЦКПОЖСЕРВІС -Дяченко*');
end;

procedure TForm1.MenuItem48Click(Sender: TObject);
{Техн. обслуг.  ВАТ ЕК *Хмельницькобленерго*  *  78-78-67, 78-78-68}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Хмельницькобленерго';
  Form1.filename:='to_hmoblenergo.html';
  Form1.ObslugaObject('Техн. обслуг.  ВАТ ЕК *Хмельницькобленерго*  *  78-78-67, 78-78-68');
end;

procedure TForm1.MenuItem49Click(Sender: TObject);
{Техн. обслуг. ТОВ * Спецпожгазмонтаж*  *  665053}
begin
  Form1.obsluga:='ОТКЛЮЧИТЬ';
  Form1.Osoba:='Спецпожгазмонтаж';
  Form1.filename:='to_spezgaz.html';
  Form1.ObslugaObject('Техн. обслуг. ТОВ * Спецпожгазмонтаж*  *  665053');

end;

procedure TForm1.MenuItem4Click(Sender: TObject);
begin

end;


end.

