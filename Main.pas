unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComObj, ShellAPI, WinProcs, WinSock, Mask, ComCtrls,
  ExtCtrls;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Button4: TButton;
    Button1: TButton;
    SG1: TStringGrid;
    Button3: TButton;
    Button5: TButton;
    Label1: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    PB2: TProgressBar;
    Label7: TLabel;
    PB3: TProgressBar;
    Label8: TLabel;
    Label9: TLabel;
    SG2: TStringGrid;
    Button2: TButton;
    Label2: TLabel;
    Label4: TLabel;
    Button6: TButton;
    GroupBox1: TGroupBox;
    Button7: TButton;
    MaskEdit1: TMaskEdit;
    MaskEdit2: TMaskEdit;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    SG3: TStringGrid;
    SG4: TStringGrid;
    Image9: TImage;
    Image10: TImage;
    Image11: TImage;
    Image12: TImage;
    Image13: TImage;
    Image14: TImage;
    MaskEdit3: TMaskEdit;
    Label10: TLabel;
    Label11: TLabel;
    PB1: TProgressBar;
    TabSheet3: TTabSheet;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    StringGrid1: TStringGrid;
    StringGrid2: TStringGrid;
    Button14: TButton;
    Label12: TLabel;
    TabSheet4: TTabSheet;
    GroupBox4: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    Image15: TImage;
    Image16: TImage;
    Label15: TLabel;
    Button15: TButton;
    MaskEdit4: TMaskEdit;
    MaskEdit5: TMaskEdit;
    MaskEdit6: TMaskEdit;
    GroupBox5: TGroupBox;
    Image17: TImage;
    Image18: TImage;
    Button16: TButton;
    Button17: TButton;
    Button18: TButton;
    GroupBox6: TGroupBox;
    Image19: TImage;
    Image20: TImage;
    Button19: TButton;
    PB4: TProgressBar;
    Label16: TLabel;
    SG6: TStringGrid;
    SG5: TStringGrid;
    Label17: TLabel;
    TabSheet5: TTabSheet;
    GroupBox7: TGroupBox;
    Label18: TLabel;
    Label19: TLabel;
    Image21: TImage;
    Image22: TImage;
    Label20: TLabel;
    Button20: TButton;
    MaskEdit7: TMaskEdit;
    MaskEdit8: TMaskEdit;
    MaskEdit9: TMaskEdit;
    GroupBox8: TGroupBox;
    Image23: TImage;
    Image24: TImage;
    Button23: TButton;
    Button22: TButton;
    Button21: TButton;
    GroupBox9: TGroupBox;
    Image25: TImage;
    Image26: TImage;
    Button24: TButton;
    PB5: TProgressBar;
    Label21: TLabel;
    SG7: TStringGrid;
    SG8: TStringGrid;
    Label22: TLabel;
    ProgressBar1: TProgressBar;
//    OpenDialog1: TOpenDialog;
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure Button21Click(Sender: TObject);
    procedure Button22Click(Sender: TObject);
    procedure Button20Click(Sender: TObject);
    procedure Button23Click(Sender: TObject);
    procedure Button24Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    var
    prise, site: string;  // prise - ���� ������, site - ���� � �����
// kurs - ���� ������, proc - ������� � ���������, nd - ������� � ���������� �����, cena - ���� � ������ ��������
    kurs, proc, nd: real;
  end;

var
  Form1: TForm1;
  //Excel: Variant;

implementation

{$R *.dfm}
// ������ ������ ����� ���������
procedure Insert_d(fail:string; lab1,lab2:TLabel; Grid:TStringGrid; bar:TProgressBar; imgz,imgk:TImage);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, j, r, q:integer;

  d: TDateTime;
begin
  d:=now;
  bar.Position:=0;
  bar.Visible:=true;
  lab1.Visible:=true;
  // ������������ ���� � �����
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(fail); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(fail)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  bar.Max:=r; // ???
  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    sheet.cells[i,19]:=strtofloat(Grid.Cells[1,i-1]); // ����
    ExLApp.columns[30].NumberFormat:='@';
    sheet.cells[i,30]:=Grid.Cells[2,i-1];   // ������� "true/false"
    ExLApp.columns[5].NumberFormat:='@';
    sheet.cells[i,5]:=Grid.Cells[3,i-1];    // ������� nalichie
    bar.Position:=i;
  End;
  Except

  end;

 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close(SaveChanges :=true);
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  lab2.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d);

  if q<>2 then
  Begin
    imgz.Visible:=true;
    imgk.Visible:=false;
  End
  Else
    imgk.Visible:=true;
  bar.Visible:=false;
  lab1.Visible:=false;
  lab2.Visible:=true;
end;

// ����������� ����������� ������
procedure Open_file(fail,proverka:string; imgz,imgk:TImage; bt:TButton; open:TOpenDialog);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet,ss : OLEVariant;
  q:integer;
  s:string;
begin
  fail:='';
  if open.Execute then fail:=open.FileName;
  if fail<>'' then
    Begin
      try
      ExLApp:=CreateOleObject('Excel.Application');
      ExLApp.Visible:=false;
      ExLApp.Workbooks.Open(fail); // ��������� ���� � �����
      Sheet:=ExLApp.Workbooks[ExtractFileName(fail)].WorkSheets[1];
      Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
      s:='';
      q:=0;
      s:=sheet.cells[1,1];
      if s=proverka then q:=1;
      if not VarIsEmpty(ExLApp) then
      begin
        ExLApp.DisplayAlerts := False; // <---
        ExLApp.Quit;
        ExLApp:=Unassigned;
      end;
      Except
        q:=2;
      end;
    End
    Else
    q:=3;

  if q=1 then
  Begin
    Bt.Enabled:=true;
    imgz.Visible:=true;
    imgk.Visible:=false;
    //prise:= fail;
  End
  Else
    imgk.Visible:=true;
end;

// ������������ ���� � �����
procedure Xsl_Open_site(XLSFile:string; Grid:TStringGrid; bar:TProgressBar);
Const
xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, r, q:integer;
Begin
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(XLSFile); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(XLSFile)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  bar.Max:=r*3; // ???
  q:=1;
  for i := 2 to r do      // ������
    Begin
          Grid.Cells[0,q]:=sheet.cells[i,7];  // ��� ������
          Grid.Cells[1,q]:=sheet.cells[i,19];  // ����
          Grid.Cells[2,q]:=sheet.cells[i,30];  // �������
          Grid.Cells[3,q]:=sheet.cells[i,5];  // ������� nalichie
//���� ����� ������ ����� ������ �� ������� �����������
//          Grid.Cells[4,q]:=sheet.cells[i,2];  // ������������ ������
//  !!!
          Grid.RowCount:=q+1;
          bar.Position:=q;
          q:=q+1;
    End;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.DisplayAlerts := False; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;
End;

procedure TForm1.Button10Click(Sender: TObject);
var
  s:string;
begin
  site:='';
  s:='product_id';
  Open_file(site,s, Image12, Image11, Button9, OpenDialog1);
  Image11.Visible:=false;
  Image12.Visible:=false;
  site:=OpenDialog1.FileName;
end;

procedure TForm1.Button13Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet,ss : OLEVariant;
  i,j,r,q,k:integer;

  s:string;
begin
  // ������������ ���� � �����
  Xsl_Open_site(site, StringGrid1, ProgressBar1);
  StringGrid1.Visible:=true;

 // ��������� ����� � ������� ��� � ������ ����� 2
  StringGrid2.Visible:=true;
    // ����� ��� ����� � ������
    StringGrid2.Cells[0,0]:='���';
    StringGrid2.Cells[1,0]:='����';
    StringGrid2.Cells[2,0]:='������� "nalichie"';
    StringGrid2.Cells[3,0]:='������� "true/false"';
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  StringGrid2.RowCount:=r;
  s:='';
  i:=0;
  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    StringGrid2.Cells[0,i-1]:=sheet.cells[i,2];  //KOD
    s:=sheet.cells[i,4];
    s:=StringReplace(s, '.', ',',[rfReplaceAll, rfIgnoreCase]);
    StringGrid2.Cells[1,i-1]:=s;    // ����
    s:='';
    try
      s:=sheet.cells[i,5];
    Except
      s:='0';
    end;
    if (s<>'0') and (s<>'') then
    Begin
      StringGrid2.Cells[3,i-1]:='���� � �������';   // ������� "nalichie"
      StringGrid2.Cells[2,i-1]:='true';             // ������� "true/false"
    End
    Else
    Begin
       StringGrid2.Cells[3,i-1]:='����p ��� �����.&lt;br&gt; �������� 1-3 ���.';   // ������� "nalichie"
       StringGrid2.Cells[2,i-1]:='true';   // ������� "true/false"
    End;
    StringGrid2.Cells[4,i-1]:=sheet.cells[i,3];
  End;
  Except
    Image11.Visible:=true;
  end;

 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;


  // ������������ StringGrid1 � StringGrid2
  // ��������� ����������� ������
  r:=StringGrid1.RowCount-1;
  ProgressBar1.Max:= StringGrid2.RowCount-1;
  StringGrid1.RowCount:=StringGrid1.RowCount+StringGrid2.RowCount;
  for i := 1 to StringGrid2.RowCount do      // ������ � ������
  Begin
    StringGrid1.Cells[0,r+i]:=StringGrid2.Cells[0,i];   // ���
    StringGrid1.Cells[1,r+i]:=StringGrid2.Cells[1,i];   // ����
    StringGrid1.Cells[2,r+i]:=StringGrid2.Cells[2,i];   // ������� "true/false"
    StringGrid1.Cells[3,r+i]:=StringGrid2.Cells[3,i];   // ������� nalichie
    StringGrid1.Cells[4,r+i]:=StringGrid2.Cells[4,i];   // ������������ ������
    ProgressBar1.Position:=i;
  End;
  StringGrid1.RowCount:=StringGrid1.RowCount-1;
end;

procedure TForm1.Button14Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, j, r, q:integer;
  s:string;
begin
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(site); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(site)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=StringGrid1.RowCount;
  ProgressBar1.Max:=StringGrid1.RowCount;

  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    sheet.cells[i,7]:=StringGrid1.Cells[0,i-1];       // kod
    sheet.cells[i,15]:=StringGrid1.Cells[0,i-1];      // model
    sheet.cells[i,19]:=strtofloat(StringGrid1.Cells[1,i-1]); // ����
    ExLApp.columns[30].NumberFormat:='@';
    sheet.cells[i,30]:=StringGrid1.Cells[2,i-1];             // ������� "true/false"
    ExLApp.columns[5].NumberFormat:='@';
    sheet.cells[i,5]:=StringGrid1.Cells[3,i-1];              // ������� nalichie
    sheet.cells[i,2]:=StringGrid1.Cells[4,i-1];
    ProgressBar1.Position:=i;
  End;
  Except

  end;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close(SaveChanges :=true);
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

end;

procedure TForm1.Button15Click(Sender: TObject);
var
  w:string;
begin
  try
    kurs:= StrtoFloat(MaskEdit4.Text);
    proc:= StrtoFloat(MaskEdit5.Text);
    nd:= StrtoFloat(MaskEdit6.Text);
    w:=floattostr(kurs)+'||'+floattostr(proc)+'||'+floattostr(nd);
  Except
     kurs:=0;
  end;
//  Label2.Caption:=w;

  // ��������
  if (kurs<>0) then
  Begin
    Image15.Visible:=true;
    Button18.Enabled:=true;
    Image16.Visible:=false;
    MaskEdit4.Enabled:=false;
    MaskEdit5.Enabled:=false;
    MaskEdit6.Enabled:=false;
  End
  Else
    Begin
    Image16.Visible:=true;
    Button18.Enabled:=false;
    Image15.Visible:=false;
    End;

end;

procedure TForm1.Button16Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, j, r, q:integer;
  s,l:string;
  cena:real;
  d: TDateTime;
begin
  d:=now;
  Label16.Visible:=true;
  PB4.Visible:=true;

  // ������������ ���� � �����
  Xsl_Open_site(site, Sg5, PB4);

  // ��������� ����� � ������� ��� � ������ ����� 6
  Sg6.Visible:=true;
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  PB4.Max:=r;
  s:='';
  i:=0;
  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    Sg6.Cells[0,q]:=sheet.cells[i,2];                   // ��� ������
    s:='';
    s:=sheet.cells[i,6];
    s:=StringReplace(s, '.', ',',[rfReplaceAll, rfIgnoreCase]);

    // ���� ������� ��������� �� 0
    if proc<>0 then
    cena:=strtofloat(s)*kurs+(strtofloat(s)*kurs)*(proc/100)+nd;
    // ���� ������� ��������� 0
    if proc=0 then cena:=strtofloat(s)*kurs+nd;

    Sg6.Cells[1,q]:=floattostr(cena);    // ����
    s:='';
    s:=sheet.cells[i,7];
    if s<>'0' then
    Begin
      Sg6.Cells[3,q]:='���� � �������';   // ������� "nalichie"
      Sg6.Cells[2,q]:='true';             // ������� "true/false"
    End;
    if s='0' then
    Begin
        s:='';
        s:=sheet.cells[i,8];
        l:='';
        l:=sheet.cells[i,9];
        if (s<>'0') or (l<>'0') then
        Begin
          Sg6.Cells[3,q]:='����p ��� �����.&lt;br&gt; �������� 1-3 ���.';   // ������� "nalichie"
          Sg6.Cells[2,q]:='true';   // ������� "true/false"
        End;
        if (s='0') and (l='0') then
        Begin
          Sg6.Cells[3,q]:='��� �����';   // ������� "nalichie"
          Sg6.Cells[2,q]:='false';   // ������� "true/false"
        End;
    End;
    Sg6.RowCount:=q+1;
    q:=q+1;
    PB4.Position:=i;
  End;
  Except
    Image17.Visible:=true;
  end;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close;
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  // ������������ ���������� � ������ ���� ������
  // ������ ��������� ������ � SG5
  PB4.Max:=Sg6.RowCount-1;
  for i := 1 to Sg6.RowCount-1  do      // ������ � ������
  Begin
      for j := 1 to Sg5.RowCount-1  do  // ������ � ����� � �����
      Begin
          if Sg6.Cells[0,i]=Sg5.Cells[0,j] then
          Begin
            Sg5.Cells[1,j]:=Sg6.Cells[1,i];   // ����
            Sg5.Cells[2,j]:=Sg6.Cells[2,i];   // ������� "true/false"
            Sg5.Cells[3,j]:=Sg6.Cells[3,i];   // ������� nalichie
          End;
      End;
      PB4.Position:=i;
  End;
  //


  PB4.Visible:=false;
  Label16.Visible:=false;
  Label17.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d);
  if Image17.Visible<>true then Image18.Visible:=true;
  Label17.Visible:=true;
  Button19.Enabled:=true;

end;

procedure TForm1.Button17Click(Sender: TObject);
var
  s:string;
begin
  prise:='';
  s:='�������� ��� ELIT';
  Open_file(prise,s, Image18, Image17, Button16, OpenDialog1);
  Image18.Visible:=false;
  Image17.Visible:=false;
  prise:=OpenDialog1.FileName;
end;

procedure TForm1.Button18Click(Sender: TObject);
var
  s:string;
Begin
  site:='';
  s:='product_id';
  Open_file(site,s, Image18, Image17, Button17, OpenDialog1);
  Image18.Visible:=false;
  Image17.Visible:=false;
  site:=OpenDialog1.FileName;
End;

procedure TForm1.Button19Click(Sender: TObject);
begin
  Insert_d(site, Label16, Label17, Sg5, PB4, Image20, Image19);
end;

procedure TForm1.Button1Click(Sender: TObject);   // ���������� ��� ����� ���� � �����
var
  s:string;
begin
  site:='';
  s:='product_id';
  Open_file(site,s, Image1, Image5, Button4, OpenDialog1);
  site:=OpenDialog1.FileName;
end;

procedure TForm1.Button20Click(Sender: TObject);
var
  w:string;
begin
  try
    kurs:= StrtoFloat(MaskEdit7.Text);
    proc:= StrtoFloat(MaskEdit8.Text);
    nd:= StrtoFloat(MaskEdit9.Text);
    w:=floattostr(kurs)+'||'+floattostr(proc)+'||'+floattostr(nd);
  Except
     kurs:=0;
  end;
  // ��������
  if (kurs<>0) then
  Begin
    Image21.Visible:=true;
    Button21.Enabled:=true;
    Image22.Visible:=false;
    MaskEdit7.Enabled:=false;
    MaskEdit8.Enabled:=false;
    MaskEdit9.Enabled:=false;
  End
  Else
    Begin
    Image22.Visible:=true;
    Button21.Enabled:=false;
    Image21.Visible:=false;
    End;
end;

procedure TForm1.Button21Click(Sender: TObject);
var
  s:string;
Begin
  site:='';
  s:='product_id';
  Open_file(site,s, Image24, Image23, Button22, OpenDialog1);
  Image24.Visible:=false;
  Image23.Visible:=false;
  site:=OpenDialog1.FileName;
End;

procedure TForm1.Button22Click(Sender: TObject);
var
  s:string;
Begin
  prise:='';
  s:='�����';
  Open_file(prise,s, Image24, Image23, Button23, OpenDialog1);
  Image24.Visible:=false;
  Image23.Visible:=false;
  prise:=OpenDialog1.FileName;
end;

procedure TForm1.Button23Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, j, r, q:integer;
  s,l,k1,k2:string;
  cena:real;
  d: TDateTime;
Begin
  d:=now;
  Label21.Visible:=true;
  PB5.Visible:=true;

  // ������������ ���� � �����
  Xsl_Open_site(site, Sg7, PB5);

  // ��������� ����� � ������� ��� � ������ �����8

ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  PB5.Max:=r;
  s:='';
  i:=0;
  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    Sg8.Cells[0,q]:=sheet.cells[i,2];                   // ��� ������
    s:='';
    s:=sheet.cells[i,4];
    s:=StringReplace(s, '.', ',',[rfReplaceAll, rfIgnoreCase]);
    // ���� ������� ��������� �� 0
    if proc<>0 then
    cena:=strtofloat(s)*kurs+(strtofloat(s)*kurs)*(proc/100)+nd;
    // ���� ������� ��������� 0
    if proc=0 then cena:=strtofloat(s)*kurs+nd;
    Sg8.Cells[1,q]:=floattostr(cena);    // ����
    s:='';
    try
      s:=sheet.cells[i,5];
    Except
      s:='0';
    end;

    if (s<>'0') and (s<>'') then
    Begin
      Sg8.Cells[3,q]:='���� � �������';   // ������� "nalichie"
      Sg8.Cells[2,q]:='true';             // ������� "true/false"
    End
    Else
    Begin
       Sg8.Cells[3,q]:='����p ��� �����.&lt;br&gt; �������� 1-3 ���.';   // ������� "nalichie"
       Sg8.Cells[2,q]:='true';   // ������� "true/false"
    End;
    Sg8.RowCount:=q+1;
    q:=q+1;
    PB5.Position:=i;
  End;
  Except
    Image23.Visible:=true;
  end;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close;
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  // ������������ ���������� � ������ ����7 ������
  // ������ ��������� ������ � SG7
  PB5.Max:=Sg8.RowCount-1;
  for i := 1 to Sg7.RowCount-1  do      // ������ � ������
  Begin
      for j := 1 to Sg8.RowCount-1  do  // ������ � ����� � �����
      Begin
          if Sg7.Cells[0,i]=Sg8.Cells[0,j] then
          Begin
            Sg7.Cells[1,i]:=Sg8.Cells[1,j];   // ����
            Sg7.Cells[2,i]:=Sg8.Cells[2,j];   // ������� "true/false"
            Sg7.Cells[3,i]:=Sg8.Cells[3,j];   // ������� nalichie
          End;
      End;
      PB5.Position:=i;
  End;

  //  end
  PB5.Visible:=false;
  Label21.Visible:=false;
  Label22.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d);
  if Image23.Visible<>true then Image24.Visible:=true;
  Label22.Visible:=true;
  Button24.Enabled:=true;

End;

procedure TForm1.Button24Click(Sender: TObject);
begin
  Insert_d(site, Label21,Label22, Sg7, PB5, Image26,Image25);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.Button3Click(Sender: TObject); // ������������ ���� � ����� (����������� ��� ������ ������� ���������� ��������)
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet,ss : OLEVariant;
  i, j, r1,r2,r3, q:integer;
  s:string;
  d: TDateTime;
begin
  d:=now;
  Sg1.Visible:=true;
  PB2.Visible:=true;
  Label7.Visible:=true;
  // ������������ ���� � �����
  Xsl_Open_site(site, Sg1, PB2);

  // ��������� ����� � ������� ��� � ������ ����� 2
  Sg2.Visible:=true;
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r2:=ExLApp.ActiveCell.Row;
  PB2.Max:=r1+r2*2;
  s:='';
  i:=0;
  q:=1;
  try
  for i := 2 to r2 do      // ������ � ������
  Begin
    s:=sheet.cells[i,2];
    if s<>'' then
    Begin
         Sg2.Cells[0,q]:=sheet.cells[i,2];   // ��� ������
         Sg2.Cells[1,q]:=sheet.cells[i,6];   // ����
         Sg2.Cells[2,q]:=sheet.cells[i,7];   // �������
         Sg2.Cells[3,q]:='���� � �������';   // ������� "nalichie"
         Sg2.RowCount:=q+1;
         q:=q+1;
    End;
    PB2.Position:=PB2.Position+1;
  End;
  Except
    Image7.Visible:=true;
  end;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

 // ������������ ������ �� ������ � � �����  ��������� � Sg1
  for i := 1 to Sg2.RowCount - 1 do      // ������ � ������
  Begin
      ss:=Sg2.Cells[0,i];
      for j := 1 to Sg1.RowCount - 1 do // stringgrid
      Begin
            if Sg2.Cells[0,i]=Sg1.Cells[0,j] then
            Begin
              Sg1.Cells[1,j]:=Sg2.Cells[1,i];  // ����
              if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
              Else Sg1.Cells[2,j]:='false';  // �������
            End
            Else   // ����������
            Begin
                if (Sg2.Cells[0,i]='1164') and (Sg1.Cells[0,j]='9509') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='6902') and (Sg1.Cells[0,j]='8052') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='3390') and (Sg1.Cells[0,j]='8048') then  //+
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B5113') and (Sg1.Cells[0,j]='B85113') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B5213') and (Sg1.Cells[0,j]='B85213') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B5313') and (Sg1.Cells[0,j]='B85313') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B5314') and (Sg1.Cells[0,j]='B85314') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B3012') and (Sg1.Cells[0,j]='B83012') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B3013') and (Sg1.Cells[0,j]='B83013') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B3014') and (Sg1.Cells[0,j]='B83014') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B2016') and (Sg1.Cells[0,j]='B82016') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B3123') and (Sg1.Cells[0,j]='B83123') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B3124') and (Sg1.Cells[0,j]='B83124') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B7310') and (Sg1.Cells[0,j]='B87310') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B1414') and (Sg1.Cells[0,j]='B81414') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B1423') and (Sg1.Cells[0,j]='B81423') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B1424') and (Sg1.Cells[0,j]='B81424') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B1433') and (Sg1.Cells[0,j]='B81433') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='B1434') and (Sg1.Cells[0,j]='B81434') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='7629') and (Sg1.Cells[0,j]='4065') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='2448') and (Sg1.Cells[0,j]='1193') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
                if (Sg2.Cells[0,i]='9500') and (Sg1.Cells[0,j]='1317') then
                Begin
                    Sg1.Cells[1,j]:=Sg2.Cells[1,i];
                    if Sg2.Cells[2,i]='����' then
                    begin
                      Sg1.Cells[2,j]:='true';
                      Sg1.Cells[3,j]:='���� � �������';   // ������� "nalichie"
                    end
                    Else Sg1.Cells[2,j]:='false';  // �������
                End;
            End;
      End;
    PB2.Position:=PB2.Position+1;
  End;


 // ����� ���������� ��������
  Label8.Visible:=true;
  Label8.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d);

 // ���� ��������
  if Sg1.RowCount>10 then
  Begin
    Button5.Enabled:=true;
    Image3.Visible:=true;
    Image7.Visible:=false;
  End
  Else
    Image7.Visible:=true;
  PB2.Visible:=false;
  Label7.Visible:=false;
end;

procedure TForm1.Button4Click(Sender: TObject);   // ���������� ��� ����� ���� ������
var
  s:string;
begin
  prise:='';
  s:='�����';
  Open_file(prise,s, Image2, Image6, Button3, OpenDialog1);
  prise:=OpenDialog1.FileName;
end;

procedure TForm1.Button5Click(Sender: TObject);  // ����������� ����� �������� � ���� � �����
begin
  Insert_d(site, Label7, Label9, Sg1, PB3, Image4,Image8);
end;

// ��������� ���� � ���������� ���� ������
// ��������� ���� � ���������� �������
procedure TForm1.Button6Click(Sender: TObject);
var
  w:string;
begin
  try
    kurs:= StrtoFloat(MaskEdit1.Text);
    proc:= StrtoFloat(MaskEdit2.Text);
    nd:= StrtoFloat(MaskEdit3.Text);
    w:=floattostr(kurs)+'||'+floattostr(proc)+'||'+floattostr(nd);
  Except
     kurs:=0;
  end;
//  Label2.Caption:=w;

  // ��������
  if (kurs<>0) then
  Begin
    Image9.Visible:=true;
    Button10.Enabled:=true;
    Image10.Visible:=false;
    MaskEdit1.Enabled:=false;
    MaskEdit2.Enabled:=false;
    MaskEdit3.Enabled:=false;
  End
  Else
    Begin
    Image10.Visible:=true;
    Button10.Enabled:=false;
    Image9.Visible:=false;
    End;

end;

procedure TForm1.Button7Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ExLApp, Sheet : OLEVariant;
  i, j, r, q:integer;
  s:string;
  cena:real;
  d: TDateTime;
begin
  d:=now;
  Label11.Visible:=true;
  PB1.Visible:=true;
 // ������������ ���� � �����
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(site); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(site)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  PB1.Max:=r*3; // ???
  q:=1;
  for i := 2 to r do      // ������
    Begin
          s:=sheet.cells[i,7];
          s:=Copy(s, 1, Pos('/', s)-1);
          Sg3.Cells[0,q]:=s;//sheet.cells[i,7];    // ��� ������
          Sg3.Cells[1,q]:=sheet.cells[i,19];   // ����
          Sg3.Cells[2,q]:=sheet.cells[i,30];   // ������� "true/false"
          Sg3.Cells[3,q]:=sheet.cells[i,5];    // ������� nalichie
          Sg3.RowCount:=q+1;
          PB1.Position:=q;
          q:=q+1;
    End;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.DisplayAlerts := False; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  // ��������� ����� � ������� ��� � ������ ����� 4
  Sg4.Visible:=true;
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise); // ��������� ���� � �����
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  r:=ExLApp.ActiveCell.Row;
  PB1.Max:=r*3;
  s:='';
  i:=0;
  q:=1;
  try
  for i := 2 to r do      // ������ � ������
  Begin
    s:=sheet.cells[i,4];
    if s<>'' then
    Begin
      if S[Length(s)]<>'�' then
         Begin
          s:=Copy(s, 1, Pos('/', s)-2);
          Sg4.Cells[0,q]:=s;   // ��� ������
          s:=sheet.cells[i,5];
          // ���� ������� ��������� �� 0
          if proc<>0 then
          cena:=strtofloat(s)*kurs+(strtofloat(s)*kurs)*(proc/100)+nd;
          // ���� ������� ��������� 0
          if proc=0 then cena:=strtofloat(s)*kurs+nd;
          Sg4.Cells[1,q]:=floattostr(cena);   // ����
          Sg4.Cells[2,q]:=sheet.cells[i,6];   // ������� 1
          Sg4.Cells[3,q]:=sheet.cells[i,7];   // ������� 2
          if (Sg4.Cells[2,q]<>'0')and(Sg4.Cells[3,q]<>'0') then
          Begin
            Sg4.Cells[4,q]:='���� � �������';   // ������� "nalichie"
            Sg4.Cells[5,q]:='true';   // ������� "true/false"
          End
          Else
          Begin
           if (Sg4.Cells[2,q]='0')and(Sg4.Cells[3,q]<>'0') then
           Begin
              Sg4.Cells[4,q]:='����p ��� �����.&lt;br&gt; �������� 1-3 ���.';   // ������� "nalichie"
              Sg4.Cells[5,q]:='true';   // ������� "true/false"
           End;
           if (Sg4.Cells[2,q]='0')and(Sg4.Cells[3,q]='0') then
           Begin
              Sg4.Cells[4,q]:='��� �����';   // ������� "nalichie"
              Sg4.Cells[5,q]:='false';   // ������� "true/false"
           End;
           if (Sg4.Cells[2,q]<>'0')and(Sg4.Cells[3,q]='0') then
           Begin
              Sg4.Cells[4,q]:='���� � �������';   // ������� "nalichie"
              Sg4.Cells[5,q]:='true';   // ������� "true/false"
           End;
          End;
          Sg4.Cells[6,q]:=sheet.cells[i,3];   // ������������
          Sg4.RowCount:=q+1;
          q:=q+1;
         End;
    End;
    PB1.Position:=PB1.Position+1;
  End;
  Except
    Image11.Visible:=true;
  end;
 if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.ActiveWorkbook.Close; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  // ������������ ���������� ������ �� ������ � ����� � �����
  for i := 1 to Sg4.RowCount-1  do      // ������ � ������
  Begin
      for j := 1 to Sg3.RowCount-1  do  // ������ � ����� � �����
      Begin
          if Sg4.Cells[0,i]=Sg3.Cells[0,j] then
          Begin
            Sg3.Cells[1,j]:=Sg4.Cells[1,i];   // ����
            Sg3.Cells[2,j]:=Sg4.Cells[5,i];   // ������� "true/false"
            Sg3.Cells[3,j]:=Sg4.Cells[4,i];   // ������� nalichie
          End;
      End;
      PB1.Position:=PB1.Position+1;
  End;

  PB1.Visible:=false;
  Label11.Visible:=false;
  Label12.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d);
  if Image11.Visible<>true then Image12.Visible:=true;
  Label12.Visible:=true;
  Button8.Enabled:=true;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
  Insert_d(site, Label11, Label12, Sg3, PB1, Image14,Image13);
end;

procedure TForm1.Button9Click(Sender: TObject);
var
  s,k:string;
begin
  prise:='';
  s:='���������� ���'; // �������� �� ������� ���� � �����
  Open_file(k,s, Image12, Image11, Button7, OpenDialog1);
  Image11.Visible:=false;
  Image12.Visible:=false;
  prise:=OpenDialog1.FileName;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
    // ����� ��� ����� � �����
    Sg1.Cells[0,0]:='���';
    Sg1.Cells[1,0]:='����';
    Sg1.Cells[2,0]:='�������';
    Sg1.Cells[3,0]:='������� nalichie';
    // ����� � ����� ������ LIQUI MOLY
    Sg2.Cells[0,0]:='���';
    Sg2.Cells[1,0]:='����';
    Sg2.Cells[2,0]:='�������';
    Sg2.Cells[3,0]:='������� nalichie';
    // ����� ��� ����� � �����
    Sg3.Cells[0,0]:='���';
    Sg3.Cells[1,0]:='����';
    Sg3.Cells[2,0]:='������� "true/false"';
    Sg3.Cells[3,0]:='������� "nalichie"';
    // ����� ��� ����� � ������ Motul
    Sg4.Cells[0,0]:='���';
    Sg4.Cells[1,0]:='����';
    Sg4.Cells[2,0]:='������� "nalichie"';
    Sg4.Cells[3,0]:='������� "true/false"';
    // ����� ��� ����� � �����
    Sg5.Cells[0,0]:='���';
    Sg5.Cells[1,0]:='����';
    Sg5.Cells[2,0]:='������� "true/false"';
    Sg5.Cells[3,0]:='������� "nalichie"';
    // ����� ��� ����� � ����� VATOIL � ELF
    Sg6.Cells[0,0]:='���';
    Sg6.Cells[1,0]:='����';
    Sg6.Cells[2,0]:='������� "true/false"';
    Sg6.Cells[3,0]:='������� "nalichie"';
    // ����� ��� ����� � �����
    Sg7.Cells[0,0]:='���';
    Sg7.Cells[1,0]:='����';
    Sg7.Cells[2,0]:='������� "true/false"';
    Sg7.Cells[3,0]:='������� "nalichie"';
    // ����� ��� ����� � ����� �����������
    Sg8.Cells[0,0]:='���';
    Sg8.Cells[1,0]:='����';
    Sg8.Cells[2,0]:='������� "true/false"';
    Sg8.Cells[3,0]:='������� "nalichie"';
    Image1.Visible:=false;
    Image2.Visible:=false;
    Image3.Visible:=false;
    Image4.Visible:=false;
    Image5.Visible:=false;
    Image6.Visible:=false;
    Image7.Visible:=false;
    Image8.Visible:=false;
    PB2.Visible:=false;
    PB3.Visible:=false;
    Sg1.Visible:=false;
    Sg2.Visible:=false;
    Label7.Visible:=false;
    Label8.Visible:=false;
    Label9.Visible:=false;
    //TabSheet3.TabVisible:=false;
end;

end.
