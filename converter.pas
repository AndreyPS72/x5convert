unit converter;

{$MODE DelphiUnicode}{$CODEPAGE UTF8}{$H+}

{$OPTIMIZATION OFF, NOREGVAR, UNCERTAIN, NOSTACKFRAME, NOPEEPHOLE, NOLOOPUNROLL, NOTAILREC, NOORDERFIELDS, NOFASTMATH, NOREMOVEEMPTYPROCS, NOCSE, NODFA} //debug Для отладки

interface


function ConvertX5File(const aFileName: string): integer;


implementation
uses SysUtils
     , Zipper
     , Laz2_DOM
     , Laz2_XMLRead
     , Laz2_XMLUtils
     , StrUtils
     , fpspreadsheet
//     , fpsopendocument
     , fpstypes
     , lazutf8
     , xlsxooxml
     , xlsxml
     ;





procedure UnZip(const aFileName, aOutDir: string);
var unzip: TUnZipper;
begin
unzip := TUnZipper.Create;
try

    unzip.Filename := aFileName;
    unzip.OutputPath := aOutDir;
    unzip.UnZipAllFiles();

except
end;
unzip.Free;
end;



const DataBufLen = 8*1024;
var DataBuf: array [0..DataBufLen-1] of double;
    samples_len: integer;
    sample_rate: double;

(*
procedure ExportWaveformToExcel(const aDir: string; var DataBuf: array of double; const DataBufLen: integer; sample_rate: double);
var ODS: TOds;
    i: integer;
    tbl: TOdfTable;
begin

ODS := nil;
try
  ODS := TOds.Create;

  if FileExists(TemplateFile) then begin
    // Открываю файл по шаблону
    if not ODS.LoadFromFile(TemplateFile) then
       Abort;
  end;

  tbl := ODS.GetTable('Waveform');
//  ODS.FindAndReplace('_ПолноеНаименование','CompanyLongName');

//  ODS.GenerateDocument(OutputFile, aDir);
//  ODS.GenerateDocument;
    ODS.ShowDocument();
except
end;

FreeAndNil(ODS);

end;
*)

(*
procedure ExportWaveformToExcel(const aDir: string; var DataBuf: array of double; const DataBufLen: integer; sample_rate: double);
var i: integer;
    Excel: OLEVariant;
    Doc: OLEVariant;
const OutputFile = 'data.xlsx';
begin

Excel:= Unassigned;
CoInitialize(nil);

try
   Excel:=GetActiveOLEObject('Excel.Application');
   Excel.Application.Visible := True;
except
  try
    Excel := CreateOLEObject('Excel.Application');
    Excel.Application.Visible := True;
  except
    // Excel Not Installed
    Abort;
  end;
end;

try
    if not FileExists(TemplateFile) then
       Abort;

    // Открываю файл по шаблону
    Doc:=Excel.Application.Workbooks.Add(UTF8Decode(TemplateFile));

except
end;

Excel.Application.ScreenUpdating := False;


Excel.Application.ScreenUpdating := True;
Doc:= Unassigned;
Excel:= Unassigned;

end;
*)



procedure ExportWaveformToExcel(const aDir: string);
var i: integer;
    wb: TsWorkbook;
    ws: TsWorksheet;
    c: PCell;
const TemplateFile = './x5.xlsx';
const OutputFile = 'data.xlsx';
begin

// Создание рабочей книги
wb := TsWorkbook.Create;
try
   if FileExists(TemplateFile) then begin
      wb.ReadFromFile(TemplateFile, sfOOXML);
      ws := wb.GetFirstWorksheet();
   end else begin
      ws := wb.AddWorksheet('Waveform');
   end;

   // Записываем ячейки
   for i:=0 to samples_len-1 do begin
       ws.WriteNumber(i, 0, double(i)/sample_rate);
       ws.WriteNumber(i, 1, DataBuf[i]);
   end;

   // Сохраняем электронную таблицу в файл
   wb.WriteToFile(ExpandFileName(OutputFile, aDir), sfOOXML, True);
except
end;
wb.Free;
end;



function StrToFloatUniversal(aStr: string): double;
begin
Result:=0.0;
try
   aStr:=ReplaceStr(DelSpace(aStr),',',DefaultFormatSettings.DecimalSeparator);
   aStr:=ReplaceStr(aStr,'.',DefaultFormatSettings.DecimalSeparator);
   Result:=StrToFloat(aStr);
except
end;
end;



procedure CheckDirForData(const aDir: string);
var Doc: TXMLDocument;
    Node: TDOMNode;
    F: longint;
begin

if not FileExists(ExpandFileName('procheck_data_signal.xml', aDir)) then
   Exit;
if not FileExists(ExpandFileName('procheck_data_signal_timestamp.bin', aDir)) then
   Exit;

sample_rate:=2560.0;

Doc := nil;
try
   ReadXMLFile(Doc, ExpandFileName('procheck_data_signal.xml', aDir));
   Node := Doc.DocumentElement.FindNode('tbl_signal_data_element');
   if Assigned(Node) and
      Node.HasChildNodes then begin
     sample_rate:=StrToFloatUniversal(TDOMElement(Node).GetAttribute('sample_rate'));
   end;

   F:=FileOpen(ExpandFileName('procheck_data_signal_timestamp.bin', aDir), fmOpenRead);
   samples_len:=FileRead (F, DataBuf, DataBufLen*sizeof(double)) div sizeof(double);
   FileClose(F);

   ExportWaveformToExcel(aDir);

except
end;
if Assigned(Doc) then
   FreeAndNil(Doc);

end;




procedure ScanFilesDir(aDir: string);
Var Info : TSearchRec;
begin

try

    If FindFirst (aDir+'/*',faDirectory, Info)=0 then begin

      Repeat
          if (Info.Name<>'.') and
             (Info.Name<>'..') then begin

             CheckDirForData(ExpandFileName(Info.Name, aDir));

             ScanFilesDir(ExpandFileName(Info.Name, aDir));
          end;

      Until FindNext(info)<>0;

      FindClose(Info);

    end;

except
end;

end;





function ConvertX5File(const aFileName: string): integer;
var FilesDir: string;
begin
if not FileExists(aFileName) then
   Exit(1);

FilesDir:=ExtractFileName(aFileName);
if Pos('.', FilesDir)>0 then
   FilesDir:=Copy(FilesDir, 1, Pos('.', FilesDir)-1);
FilesDir:='./'+ FilesDir;
UnZip(aFileName, FilesDir);

ScanFilesDir(FilesDir);

Result:=0;
end;


end.



