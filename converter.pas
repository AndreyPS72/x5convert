unit converter;

{$MODE DelphiUnicode}{$CODEPAGE UTF8}{$H+}

interface

function ConvertX5File(const aFileName: string): integer;


implementation
uses Forms
     , SysUtils
     , FileUtil
     , Zipper
     , Laz2_DOM
     , Laz2_XMLRead
     , Laz2_XMLUtils
     , StrUtils
     , fpspreadsheet
     , fpstypes
     , lazutf8
     , xlsxooxml
     , xlsxml
     , rFFT
     , formprogress
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



var DataBuf: TReal64ArrayZeroBased;
    ts_nr_of_samples: integer;
    sample_rate: double;
    sensor_id: string;
    signal_type: integer;
    units: integer;
    create_date,
    change_date,
    measure_time: string;

const TemplateFile = 'x5.xlsx';
//const OutputFile = 'data.xlsx';

var UnzippedFilesDir: string;


function GetExportFileName(const aDir: string; const DT: TDateTime; const sensor_id: UnicodeString): string;
var OutputFile: UnicodeString;
begin
OutputFile:=FormatDateTime('yyymmdd_hhnnss',DT)+'_'+sensor_id+'.xlsx';
result:=ExpandFileName(OutputFile, aDir);
end;


function GetExportSummaryFileName(const aDir: string; const DT: TDateTime; const sensor_id: UnicodeString): string;
var OutputFile: UnicodeString;
    OutputDir: UnicodeString;
begin
OutputDir:=ExpandFileName('id_'+sensor_id, UnzippedFilesDir);
if not DirectoryExists(OutputDir) then
   MkDir(OutputDir);
OutputFile:=GetExportFileName(aDir, DT, sensor_id);
OutputFile:=ReplaceStr(OutputFile, UnzippedFilesDir, '');
OutputFile:=ReplaceStr(OutputFile, '\', '_');
OutputFile:=ReplaceStr(OutputFile, '/', '_');
result:=ExpandFileName(OutputFile, OutputDir);
end;



procedure ExportWaveformToExcel(const aDir: string; const DT: TDateTime; const sensor_id: UnicodeString);
var i: integer;
    wb: TsWorkbook;
    ws: TsWorksheet;
    c: PCell;
    dF, v: double;
    name_summary: UnicodeString;
begin

// Создание рабочей книги
wb := TsWorkbook.Create;
wb.Options :=[boCalcBeforeSaving, boReadformulas];

try
   if FileExists(TemplateFile) then begin
      wb.ReadFromFile(TemplateFile, sfOOXML);
      ws := wb.GetFirstWorksheet();
   end else begin
      ws := wb.AddWorksheet('Waveform');
   end;

   // Записываем ячейки
   for i:=0 to ts_nr_of_samples-1 do begin
       ws.WriteNumber(i, 0, double(i)/sample_rate);
       ws.WriteNumber(i, 1, DataBuf[i] * 9.81);
   end;

   // Параметры сигнаал
   ws.WriteText(0, 10, 'sensor_id');        ws.WriteText(0, 11, sensor_id);
   ws.WriteText(1, 10, 'signal_type');      ws.WriteNumber(1, 11, signal_type);
   ws.WriteText(2, 10, 'unit');             ws.WriteNumber(2, 11, units);
   ws.WriteText(3, 10, 'ts_nr_of_samples'); ws.WriteNumber(3, 11, ts_nr_of_samples);
   ws.WriteText(4, 10, 'sample_rate');      ws.WriteNumber(4, 11, sample_rate);
   ws.WriteText(5, 10, 'create_date');      ws.WriteText(5, 11, create_date);
   ws.WriteText(6, 10, 'change_date');      ws.WriteText(6, 11, change_date);
   ws.WriteText(7, 10, 'measure_time');     ws.WriteText(7, 11, measure_time);

   RealFFT(@DataBuf, ts_nr_of_samples);
   dF := sample_rate / ((ts_nr_of_samples-1) * 2);

   if FileExists(TemplateFile) then begin
      ws := wb.GetNextWorksheet(ws);
   end else begin
      ws := wb.AddWorksheet('Spectrum');
   end;

   // Записываем ячейки
   ws.WriteNumber(0, 0, 0);
   ws.WriteNumber(0, 1, DataBuf[0] * 9.81); // Acc[0]
   ws.WriteNumber(0, 2, 0); // Vel[0]
   ws.WriteNumber(0, 3, 0); // Disp[0]

   for i:=1 to ts_nr_of_samples-1 do begin
       ws.WriteNumber(i, 0, dF * double(i));
       v:=DataBuf[i] * 9.81;
       ws.WriteNumber(i, 1, v); // Acc
       v:=v * 1000.0 / (2.0 * Pi * dF * double(i));
       ws.WriteNumber(i, 2, v); // Vel
       v:=v * 1000.0 / (2.0 * Pi * dF * double(i));
       ws.WriteNumber(i, 3, v); // Disp
   end;

   // Сохраняем электронную таблицу в файл
   wb.WriteToFile(GetExportFileName(aDir, DT, sensor_id), sfOOXML, True);

   CopyFile(GetExportFileName(aDir, DT, sensor_id), GetExportSummaryFileName(aDir, DT, sensor_id));

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
    DT: TDateTime;
    s:string;
begin

if not FileExists(ExpandFileName('procheck_data_signal.xml', aDir)) then
   Exit;
if not FileExists(ExpandFileName('procheck_data_signal_timestamp.bin', aDir)) then
   Exit;

sample_rate:=2560.0;
ts_nr_of_samples:=DataBufLen;
sensor_id:='';

Doc := nil;
try
   ReadXMLFile(Doc, ExpandFileName('procheck_data_signal.xml', aDir));
   Node := Doc.DocumentElement.FindNode('tbl_signal_data_element');
   if Assigned(Node) and
      Node.HasChildNodes then begin
     sample_rate:=StrToFloatUniversal(TDOMElement(Node).GetAttribute('sample_rate'));
     ts_nr_of_samples:=StrToInt(TDOMElement(Node).GetAttribute('ts_nr_of_samples'));
     sensor_id:=Trim(TDOMElement(Node).GetAttribute('sensor_id'));
     signal_type:=StrToInt(TDOMElement(Node).GetAttribute('signal_type'));
     units:=StrToInt(TDOMElement(Node).GetAttribute('unit'));
     DT:=StrToFloatUniversal(TDOMElement(Node).GetAttribute('create_date')); create_date:=DateTimeToStr(DT);
     DT:=StrToFloatUniversal(TDOMElement(Node).GetAttribute('change_date')); change_date:=DateTimeToStr(DT);
     DT:=StrToFloatUniversal(TDOMElement(Node).GetAttribute('measure_time')); measure_time:=DateTimeToStr(DT);
   end;

   F:=FileOpen(ExpandFileName('procheck_data_signal_timestamp.bin', aDir), fmOpenRead);
   ts_nr_of_samples:=FileRead (F, DataBuf, DataBufLen*sizeof(double)) div sizeof(double);
   FileClose(F);

   ExportWaveformToExcel(aDir, DT, sensor_id);

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

             FormMain.pbProgress.StepIt;
             Application.ProcessMessages;
             if StopProcess then
                break;
          end;

      Until FindNext(info)<>0;

      FindClose(Info);

    end;

except
end;

end;



var count_files: integer;

procedure CountFiles(aDir: string);
Var Info : TSearchRec;
begin

try

    If FindFirst (aDir+'/*',faDirectory, Info)=0 then begin

      Repeat
          if (Info.Name<>'.') and
             (Info.Name<>'..') then begin

             inc(count_files);

             CountFiles(ExpandFileName(Info.Name, aDir));
          end;

      Until FindNext(info)<>0;

      FindClose(Info);

    end;
except
end;

end;




procedure RealFFTTest();
var i: integer;
begin
ts_nr_of_samples:=8192;
for i:=0 to ts_nr_of_samples-1 do begin
      DataBuf[i]:=sin(i*0.01);
end;
RealFFT(@DataBuf, ts_nr_of_samples);
end;




function ConvertX5File(const aFileName: string): integer;
begin

if not FileExists(aFileName) then
   Exit(1);

FormMain.Show;
Application.ProcessMessages;

UnzippedFilesDir:=ExtractFileName(aFileName);
if Pos('.', UnzippedFilesDir)>0 then
   UnzippedFilesDir:=Copy(UnzippedFilesDir, 1, Pos('.', UnzippedFilesDir)-1);
UnzippedFilesDir:='./'+ UnzippedFilesDir;
UnZip(aFileName, UnzippedFilesDir);

count_files:=0;
CountFiles(UnzippedFilesDir);
FormMain.pbProgress.Max:=count_files;

ScanFilesDir(UnzippedFilesDir);

Result:=0;
end;




end.



