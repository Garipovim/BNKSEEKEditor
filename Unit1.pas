unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, FileCtrl, DB, ADODB, Grids, DBGrids, ExtCtrls, DBCtrls, DBTables,
  Buttons ;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Button2: TButton;
    CNDBF: TADOConnection;
    ADOQuery1: TADOQuery;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    DataSource2: TDataSource;
    ADOQuery2: TADOQuery;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    DataSource4: TDataSource;
    ADOQuery4: TADOQuery;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    DBLookupComboBox1: TDBLookupComboBox;
    Button3: TButton;
    BitBtn1: TBitBtn;
    Button4: TButton;
    DBGrid2: TDBGrid;
    Button5: TButton;
    Button6: TButton;
    ADOQuery5: TADOQuery;
    DataSource5: TDataSource;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Path : string;   //переменная в которой будем хранить путь к исполняемому файлу.
implementation

{$R *.dfm}



// возвращает каталог, выбранный пользователем
function GetPath(mes: string):string;
var
  Root: string;      // корневой каталог
  pwRoot : PWideChar;
  Dir: string;
begin

  Root := ''; //Корневой каталог
  Dir:= ExtractFileDir(Application.ExeName);
  GetMem(pwRoot, (Length(Root)+1) * 2);
  pwRoot := StringToWideChar(Root,pwRoot,MAX_PATH*2);
  if SelectDirectory(mes, pwRoot, Dir)
     then
          if length(Dir) = 2  // пользователь выбрал корневой каталог
              then GetPath := Dir+'\'
              else GetPath := Dir
     else
          GetPath := '';
end;


procedure TForm1.Button1Click(Sender: TObject);
var
  Path: string;
begin
  Path := GetPath('Выберите папку');
  if Path <> ''
     then Edit1.Text := Path;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
DBgrid2.Visible:=false; //Скроем таблицу в которой будем редактировать данные и вводить новые
CnDBF.Connected:=false;
CnDBF.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' +Edit1.Text + ';Extended Properties=dBASE IV;;;';
CnDBF.LoginPrompt:=false;
cnDBF.Mode:=cmReadWrite;
CnDBF.Connected:=true;
//Заполняем таблицу в которой отображаются наши данные
try
       ADOQuery1.SQL.Clear;
       ADOQuery1.SQL.Add(' SELECT TBnkseek.REAL, TPzn.NAME as TPZN_NAME, ');
       ADOQuery1.SQL.Add(' UERn.UERNAME as UER_NAME, REgn.Name as REG_name, TBnkseek.IND, TNPn.FULLNAME as Tnpn_fullname , TBnkseek.NNP, TBnkseek.ADR, TBnkseek.RKC, TBnkseek.NAMEP, TBnkseek.NEWNUM,  ');
       ADOQuery1.SQL.Add(' TBnkseek.TELEF, TBnkseek.REGN, TBnkseek.OKPO, TBnkseek.DT_IZM, TBnkseek.KSNP, TBnkseek.DATE_IN, TBnkseek.DATE_CH ');
       ADOQuery1.SQL.Add(' FROM (((( bnkseek.dbf as TBnkseek ');
       ADOQuery1.SQL.Add(' left join tnp.dbf as TNPn on TNPn.TNP=TBnkseek.TNP)' );
       ADOQuery1.SQL.Add(' left join pzn.dbf as TPzn on TPzn.PZN=TBnkseek.PZN)' );
       ADOQuery1.SQL.Add(' left join UER.dbf as UERn on UERn.UER=TBnkseek.UER)' );
       ADOQuery1.SQL.Add(' left join REG.dbf as REGn on REGn.RGN=TBnkseek.RGN)' );
       ADOQuery1.Active:=true;


      except on e:exception do
        Showmessage(e.ClassName + ':' + e.Message);
    end;


 //Заполним данными выпадающий список для фильтра
    try
       ADOQuery2.SQL.Clear;
       ADOQuery2.SQL.Add(' SELECT Name FROM PZN.dbf ');
       ADOQuery2.Active:=true;

        ADOQuery2.First;
           while not ADOQuery2.Eof do
            Begin
             DBGrid1.Columns[1].PickList.add(ADOQuery2.FieldByName('NAME').AsString);
             ADOQuery2.Next;
            End;
 //Поскольку данные в выпадающий список загружены удачно, разблокируем кнопку фильтра.
     Form1.Button3.Enabled:=True;
     Form1.Button4.Enabled:=True;
    except on e:exception do
        Showmessage(e.ClassName + ':' + e.Message);
    end;

//Защитим нашу основную таблицу от редактирования.
    DBgrid1.ReadOnly:=True;

//Сделаем видимыми кнопки редактирования и добавления записей
   Button5.Visible:=true;
   Button6.Visible:=true;
   Bitbtn1.Visible:=true;
//Связываем таблицу редактирования и источник данных
   DBgrid2.DataSource:=  DataSource4;
end;


procedure TForm1.LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
begin
// Настроим маску в поле БИК  фильтра
 if not (Key in ['0'..'9', #8])then Key:=#0;
end;


procedure TForm1.Button3Click(Sender: TObject);
var

FNewNUM, FRgn, FPZN: string;
begin


 FNewNUM  := quotedStr('%'+ LabeledEdit1.text +'%');
 FRgn     := quotedStr('%'+ LabeledEdit2.text +'%');

 FPZN     :=quotedStr(DBLookupcombobox1.text);

 try
       ADOQuery1.SQL.Clear;
       ADOQuery1.SQL.Add(' SELECT 	TBnkseek.REAL, TPzn.NAME as TPZN_NAME, ');
       ADOQuery1.SQL.Add(' UERn.UERNAME as UER_NAME, REgn.Name as REG_name, TBnkseek.IND, TNPn.FULLNAME as Tnpn_fullname , TBnkseek.NNP, TBnkseek.ADR, TBnkseek.RKC, TBnkseek.NAMEP, TBnkseek.NEWNUM,  ');
       ADOQuery1.SQL.Add(' TBnkseek.TELEF, TBnkseek.REGN, TBnkseek.OKPO, TBnkseek.DT_IZM, TBnkseek.KSNP, TBnkseek.DATE_IN, TBnkseek.DATE_CH ');
       ADOQuery1.SQL.Add(' FROM (((( bnkseek.dbf as TBnkseek ');
       ADOQuery1.SQL.Add(' left join tnp.dbf as TNPn on TNPn.TNP=TBnkseek.TNP)' );
       ADOQuery1.SQL.Add(' left join pzn.dbf as TPzn on TPzn.PZN=TBnkseek.PZN)' );
       ADOQuery1.SQL.Add(' left join UER.dbf as UERn on UERn.UER=TBnkseek.UER)' );
       ADOQuery1.SQL.Add(' left join REG.dbf as REGn on REGn.RGN=TBnkseek.RGN)' );
       ADOQuery1.SQL.Add(' where TBnkseek.NEWNUM like ' + FNewNUM);
       ADOQuery1.SQL.Add(' and  REgn.Name like ' + FRgn);             // По полям БИК, РЕГИОН поиск идет по частичному включению
       IF  DBLookupcombobox1.text <> '' then
         ADOQuery1.SQL.Add(' and  TPzn.NAME = ' + FPZN);              //по выпадающему списку -  поиск идет по точному совпадению

       ADOQuery1.Active:=true;
except on e:exception do
        Showmessage(e.ClassName + ':' + e.Message);
end;

end;

procedure TForm1.Button4Click(Sender: TObject);
begin
//очистим поля фильтра
LabeledEdit1.Text:='';
LabeledEdit2.Text:='';
DBLookUpComboBox1.KeyValue := NULL;
//после очистки полей фильтра нам надо перезаполнить таблицу.
 Button3.Click;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
//Заполним таблицу редактирования активной строчкой из основной таблицы
       ADOQuery4.SQL.Clear;
       ADOQuery4.SQL.Add(' SELECT TBnkseek.VKEY,	TBnkseek.REAL, TPzn.NAME as TPZN_NAME, ');
       ADOQuery4.SQL.Add(' UERn.UERNAME as UER_NAME, REgn.Name as REG_name, TBnkseek.IND, TNPn.FULLNAME as Tnpn_fullname , TBnkseek.NNP, TBnkseek.ADR, TBnkseek.RKC, TBnkseek.NAMEP,TBnkseek.NAMEN, TBnkseek.NEWNUM,  ');
       ADOQuery4.SQL.Add(' TBnkseek.NEWKS, TBnkseek.PERMFO,	TBnkseek.SROK,	TBnkseek.AT1,	TBnkseek.AT2, ');
       ADOQuery4.SQL.Add('  TBnkseek.TELEF, TBnkseek.REGN, TBnkseek.OKPO, TBnkseek.DT_IZM, TBnkseek.CKS, TBnkseek.KSNP, TBnkseek.DATE_IN, TBnkseek.DATE_CH, ');
       ADOQuery4.SQL.Add('  TBnkseek.VKEYDEL,	TBnkseek.DT_IZMR');
       ADOQuery4.SQL.Add(' FROM (((( bnkseek.dbf as TBnkseek ');
       ADOQuery4.SQL.Add(' left join tnp.dbf as TNPn on TNPn.TNP=TBnkseek.TNP)' );
       ADOQuery4.SQL.Add(' left join pzn.dbf as TPzn on TPzn.PZN=TBnkseek.PZN)' );
       ADOQuery4.SQL.Add(' left join UER.dbf as UERn on UERn.UER=TBnkseek.UER)' );
       ADOQuery4.SQL.Add(' left join REG.dbf as REGn on REGn.RGN=TBnkseek.RGN)' );
       ADOQuery4.SQL.Add('  where TBnkseek.newnum = ' + quotedStr(ADOquery1.fieldbyname('newnum').AsString));

       ADOQuery4.Active:=true;
       DBgrid2.Visible:=true;
       DBgrid2.Columns[0].Visible:=false;
       BitBtn1.Visible:=false;

       //   Заполним выпадающие списки в полях PZN, TNP, UER,REG  соответствующими значениями из таблиц.
    try
       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT Name FROM PZN.dbf ');
       ADOQuery3.Active:=true;

        ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[2].PickList.add(ADOQuery3.FieldByName('NAME').AsString);
             ADOQuery3.Next;
            End;
        ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM TNP.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[6].PickList.add(ADOQuery3.FieldByName('FullName').AsString);
             ADOQuery3.Next;
            End;
       ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM UER.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[3].PickList.add(ADOQuery3.FieldByName('UERName').AsString);
             ADOQuery3.Next;
            End;
       ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM REG.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[4].PickList.add(ADOQuery3.FieldByName('Name').AsString);
             ADOQuery3.Next;
            End;
    except on e:exception do
        Showmessage(e.ClassName + ':' + e.Message);
end;

  Form1.Bitbtn1.Visible:=True;
end;

procedure TForm1.Button6Click(Sender: TObject);
Var
n:integer;
begin
   //перед внесением изменений в базу, проверим на заполнение обязательных реквизитов
   If  (ADOQuery4.Fields[0].AsString = '') or ( ADOQuery4.FieldbyName('newnum').AsString = '') or (( ADOQuery4.FieldbyName('uer_name').AsString = ''))
      OR (ADOQuery4.FieldbyName('reg_name').AsString = '') OR (ADOQuery4.FieldbyName('namep').AsString = '') OR (ADOQuery4.FieldbyName('namen').AsString = '')
      or (ADOQuery4.FieldbyName('SROK').AsString = '') OR (ADOQuery4.FieldbyName('DT_IZM').AsString = '') OR (ADOQuery4.FieldbyName('DATE_IN').AsString = '')
   Then
        begin
          Showmessage('Не заполнены обязательные реквизиты') ;
          exit;
        end;

  AdoQuery3.SQL.Clear;
  AdoQuery3.SQL.Add('SELECT   * from  BNKSEEK.dbf WHERE newnum = ' + quotedStr(ADOQuery4.FieldbyName('newnum').AsString));
  ADOQuery3.Active:=true;
  ADOQuery3.Last;
  n := ADOQuery3.RecordCount;
  If n >0 Then  //Если поле с заданным БИК уже существует
   Begin
    If n >1 Then //Если таких полей больше чем 1
      Begin
        Showmessage('Обнаружено дублирование БИК ' + ADOQuery4.FieldbyName('newnum').AsString);
        exit;
      end;
//Если вводимое новое значение БИК уже встречается в другой строке таблицы. Сравниваем по VKEY.
    If  ADOQuery3.fieldbyname('VKEY').asstring <> ADOQuery4.FieldbyName('VKEY').AsString  then
      Begin
        Showmessage('Обнаружено дублирование БИК ' + ADOQuery4.FieldbyName('newnum').AsString);
        exit;
      end;
   end;

//В новый запрос собираем значения из подставных полей
  ADOQuery5.sql.Clear;
  ADOQuery5.sql.Add(' SELECT * From BNKSEEK.dbf');
  ADOQuery5.sql.add(' where VKEY = ' +  quotedStr(ADOQuery4.FieldbyName('VKEY').AsString));
  ADOQuery5.Active:=true;

  ADOQuery4.FieldbyName('newnum').AsString;
  ADOQuery5.Edit;
  ADOQuery5.Fields[0].AsString := ADOQuery4.Fields[0].AsString;
  ADOQuery5.Fields[1].AsString := ADOQuery4.Fields[1].AsString;

    ADOQuery3.sql.Clear;
    ADOQuery3.SQL.Add(' SELECT PZN FROM PZN.dbf where NAME = ' + quotedStr(ADOQuery4.FieldbyName('TPZN_Name').AsString));
    ADOQuery3.Active:=true;
  ADOQuery5.Fields[2].AsString := ADOQuery3.fieldByName('PZN').asstring;

    ADOQuery3.sql.Clear;
    ADOQuery3.SQL.Add(' SELECT UER FROM UER.dbf where UERName = ' + quotedStr(ADOQuery4.FieldbyName('UER_Name').AsString));
    ADOQuery3.Active:=true;

  ADOQuery5.Fields[3].AsString := ADOQuery3.fieldByName('UER').asstring;

    ADOQuery3.sql.Clear;
    ADOQuery3.SQL.Add(' SELECT RGN FROM REG.dbf where NAME = ' + quotedStr(ADOQuery4.FieldbyName('REG_name').AsString));
    ADOQuery3.Active:=true;
  ADOQuery5.Fields[4].AsString := ADOQuery3.fieldByName('RGN').asstring;;


  ADOQuery5.Fields[5].AsString := ADOQuery4.Fields[5].AsString;


    ADOQuery3.sql.Clear;
    ADOQuery3.SQL.Add(' SELECT TNP FROM TNP.dbf where FULLNAME = ' + quotedStr(ADOQuery4.FieldbyName('Tnpn_fullname').AsString));
    ADOQuery3.Active:=true;
  ADOQuery5.Fields[6].AsString := ADOQuery3.fieldByName('TNP').asstring;

 //соберем совпадающие поля

  ADOQuery5.Fields[7].AsString := ADOQuery4.Fields[7].AsString;
  ADOQuery5.Fields[8].AsString := ADOQuery4.Fields[8].AsString;
  ADOQuery5.Fields[9].AsString := ADOQuery4.Fields[9].AsString;
  ADOQuery5.Fields[10].AsString := ADOQuery4.Fields[10].AsString;
  ADOQuery5.Fields[11].AsString := ADOQuery4.Fields[11].AsString;
  ADOQuery5.Fields[12].AsString := ADOQuery4.Fields[12].AsString;
  ADOQuery5.Fields[13].AsString := ADOQuery4.Fields[13].AsString;
  ADOQuery5.Fields[14].AsString := ADOQuery4.Fields[14].AsString;
  ADOQuery5.Fields[15].AsString := ADOQuery4.Fields[15].AsString;
  ADOQuery5.Fields[16].AsString := ADOQuery4.Fields[16].AsString;
  ADOQuery5.Fields[17].AsString := ADOQuery4.Fields[17].AsString;
  ADOQuery5.Fields[18].AsString := ADOQuery4.Fields[18].AsString;
  ADOQuery5.Fields[19].AsString := ADOQuery4.Fields[19].AsString;
  ADOQuery5.Fields[20].AsString := ADOQuery4.Fields[20].AsString;
  ADOQuery5.Fields[21].AsString := ADOQuery4.Fields[21].AsString;
  ADOQuery5.Fields[22].AsString := ADOQuery4.Fields[22].AsString;
  ADOQuery5.Fields[23].AsString := ADOQuery4.Fields[23].AsString;
  ADOQuery5.Fields[24].AsString := ADOQuery4.Fields[24].AsString;
  ADOQuery5.Fields[25].AsString := ADOQuery4.Fields[25].AsString;
  ADOQuery5.Fields[26].AsString := ADOQuery4.Fields[26].AsString;
  ADOQuery5.Fields[27].AsString := ADOQuery4.Fields[27].AsString;
//  Внесем изменения в базу
  ADOQuery5.Post;
//Перечитаем основную таблицу, чтоб там отразились новые данные
  DBGrid1.datasource.dataset.close;
  DBGrid1.datasource.dataset.open;
//Спрячем таблицу редактирования
  DBgrid2.Visible:=false;
//закроем запрос
 ADOQuery4.Close;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
       //В таблицу для редактирования добавим пустую строку и заголовки столбцов
       ADOQuery4.SQL.Clear;
       ADOQuery4.SQL.Add(' SELECT TBnkseek.VKEY,	TBnkseek.REAL, TPzn.NAME as TPZN_NAME, ');
       ADOQuery4.SQL.Add(' UERn.UERNAME as UER_NAME, REgn.Name as REG_name, TBnkseek.IND, TNPn.FULLNAME as Tnpn_fullname , TBnkseek.NNP, TBnkseek.ADR, TBnkseek.RKC, TBnkseek.NAMEP,TBnkseek.NAMEN, TBnkseek.NEWNUM,  ');
       ADOQuery4.SQL.Add(' TBnkseek.NEWKS, TBnkseek.PERMFO,	TBnkseek.SROK,	TBnkseek.AT1,	TBnkseek.AT2, ');
       ADOQuery4.SQL.Add('  TBnkseek.TELEF, TBnkseek.REGN, TBnkseek.OKPO, TBnkseek.DT_IZM, TBnkseek.CKS, TBnkseek.KSNP, TBnkseek.DATE_IN, TBnkseek.DATE_CH, ');
       ADOQuery4.SQL.Add('  TBnkseek.VKEYDEL,	TBnkseek.DT_IZMR');
       ADOQuery4.SQL.Add(' FROM (((( bnkseek.dbf as TBnkseek ');
       ADOQuery4.SQL.Add(' left join tnp.dbf as TNPn on TNPn.TNP=TBnkseek.TNP)' );
       ADOQuery4.SQL.Add(' left join pzn.dbf as TPzn on TPzn.PZN=TBnkseek.PZN)' );
       ADOQuery4.SQL.Add(' left join UER.dbf as UERn on UERn.UER=TBnkseek.UER)' );
       ADOQuery4.SQL.Add(' left join REG.dbf as REGn on REGn.RGN=TBnkseek.RGN)' );
       ADOQuery4.SQL.Add('  where TBnkseek.VKEY = NULL' );

       ADOQuery4.Active:=true;
       DBgrid2.Visible:=true;
       Button5.Visible:=false;
    //Заполним значениями выпадающие списки
    try
       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT Name FROM PZN.dbf ');
       ADOQuery3.Active:=true;

        ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[2].PickList.add(ADOQuery3.FieldByName('NAME').AsString);
             ADOQuery3.Next;
            End;
        ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM TNP.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[6].PickList.add(ADOQuery3.FieldByName('FullName').AsString);
             ADOQuery3.Next;
            End;
       ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM UER.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[3].PickList.add(ADOQuery3.FieldByName('UERName').AsString);
             ADOQuery3.Next;
            End;
       ADOQuery3.Active:=False;

       ADOQuery3.SQL.Clear;
       ADOQuery3.SQL.Add(' SELECT * FROM REG.dbf ');
       ADOQuery3.Active:=true;
       ADOQuery3.First;
           while not ADOQuery3.Eof do
            Begin
             DBGrid2.Columns[4].PickList.add(ADOQuery3.FieldByName('Name').AsString);
             ADOQuery3.Next;
            End;
    except on e:exception do
        Showmessage(e.ClassName + ':' + e.Message);

end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
Path:=ExtractFileDir(Application.ExeName);
Edit1.Text:= Path;
end;

end.
