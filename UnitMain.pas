unit UnitMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, ComObj, Controls, Forms,
  Dialogs, StdCtrls;

type
  TFormMain = class(TForm)
    ButtonPrint: TButton;
    LabelStatus: TLabel;
    procedure ButtonPrintClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

procedure WordReplace(var WordDoc: OleVariant; bookmark: string; str: string);
begin
WordDoc.activedocument.range.Find.execute
         (bookmark,EmptyParam, EmptyParam,EmptyParam,
          EmptyParam,EmptyParam,EmptyParam,EmptyParam,
          EmptyParam,str,1,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
end;

procedure CertificatePrint(family: string; name: string; otch:string; passSeries: string;
  passNumber: string; ball_bio: string; ball_rus: string; ball_chem: string; sege_bio: string;
    sege_rus: string; sege_chem: string);
var
  WordDoc: OleVariant;
  FileNameDot, FileNameExport, CurrDir: String;
begin
FileNameDot := 'Certificate.dot';
  FileNameExport := family+' '+name+' '+otch+' '+passSeries+' '+passNumber+'.doc';
  CurrDir := ExtractFilePath(Application.ExeName)+'Справки\';
  //export to word
  WordDoc:=CreateOleObject('word.Application');
  WordDoc.Application.Documents.Add(CurrDir + FileNameDot, EmptyParam,EmptyParam,EmptyParam);
  WordDoc.Visible := false;//make word invisible
  //Family
  wordReplace(WordDoc,'@family', family);       
  //Name
  wordReplace(WordDoc,'@name', name);
  //Otch
  wordReplace(WordDoc,'@otch', otch);
  //Family
  wordReplace(WordDoc,'@passSeries', passSeries);
  //Family
  wordReplace(WordDoc,'@passNumber', passNumber);
  //ball biology
  wordReplace(WordDoc,'@ball_bio', ball_bio);    
  //sege biology
  wordReplace(WordDoc,'@sege_bio', sege_bio);
  //ball chemistry
  wordReplace(WordDoc,'@ball_chem', ball_chem);
  //sege chemistry
  wordReplace(WordDoc,'@sege_chem', sege_chem);
  //ball russian
  wordReplace(WordDoc,'@ball_rus', ball_rus);
  //sege russian
  wordReplace(WordDoc,'@sege_rus', sege_rus);
  //date and time
//  dateStr := DateToStr(date);
  wordReplace(WordDoc,'@datetime', DateToStr(date)+' '+TimeToStr(time));
  WordDoc.ActiveDocument.SaveAs(CurrDir + FileNameExport);
  WordDoc.Quit;
end;

procedure TFormMain.ButtonPrintClick(Sender: TObject);
var
  Excel: Variant;
  family, nextFamily, name, otch, passSeries, passNumber, subject, year, status, egeNom: string;
  sege_bio, sege_rus, sege_chem: string;
  ball_chem, ball_rus, ball_bio, stringNumber: integer;
begin
//  CertificatePrint('Иванов', 'Пётр', 'Олегович', '2345', '234567', '55', '66', '77');
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[ExtractFilePath(Application.ExeName)+'result.xlsx'];//, 0, True];
  stringNumber := 1;
  ball_chem := 0;
  ball_rus := 0;
  ball_bio := 0;
    sege_bio := '-';
    sege_rus := '-';
    sege_chem := '-';
  family := Excel.Cells[stringNumber, 1];
  while('' <> family) do
  begin
    family := Excel.Cells[stringNumber, 1];
    LabelStatus.Caption := 'Обработка '+family;
    name := Excel.Cells[stringNumber, 2];
    otch := Excel.Cells[stringNumber, 3];
    passSeries := Excel.Cells[stringNumber, 4];
    passNumber := Excel.Cells[stringNumber, 5];
    subject := Excel.Cells[stringNumber, 6];
    case subject[1] of//-
      'Х' :
            begin
              ball_chem := StrToInt(Excel.Cells[stringNumber, 7]);
              sege_chem := Excel.Cells[stringNumber, 12];
              if('' = sege_chem) then
               sege_chem := #0151;
            end;
      'Р' :
            begin
              ball_rus := StrToInt(Excel.Cells[stringNumber, 7]);
              sege_rus := Excel.Cells[stringNumber, 12];
              if('' = sege_rus) then
               sege_rus := #0151;
            end;
      'Б' :
            begin
              ball_bio := StrToInt(Excel.Cells[stringNumber, 7]);
              sege_bio := Excel.Cells[stringNumber, 12];
              if('' = sege_bio) then
               sege_bio := #0151;
            end;
//    ball := Excel.Cells[stringNumber, 7];
    end;
    year := Excel.Cells[stringNumber, 8];
    status := Excel.Cells[stringNumber, 10];
    egeNom := Excel.Cells[stringNumber, 12];
    nextFamily := Excel.Cells[stringNumber+1, 1];
    if(nextFamily <> family) then
    begin
      CertificatePrint(family, name, otch, passSeries, passNumber,
        IntToStr(ball_bio), IntToStr(ball_rus), IntToStr(ball_chem), sege_bio, sege_rus, sege_chem);
      sege_bio := '-';
      sege_rus := '-';
      sege_chem := '-';
      ball_chem := 0;
      ball_rus := 0;
      ball_bio := 0;
    end;
    stringNumber := stringNumber+1;
    Application.ProcessMessages;
  end;
  LabelStatus.Caption := 'Обработка завершена, ожидание';   //«акрытие Excel
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
end;

end.
