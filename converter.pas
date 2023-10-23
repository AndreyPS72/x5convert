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

procedure CheckDirForData(aDir: string);
var Doc: TXMLDocument;
    RootNode: TDOMNode;
    F: longint;
    len: integer;
begin

if not FileExists(ExpandFileName('procheck_data_signal.xml', aDir)) then
   Exit;
if not FileExists(ExpandFileName('procheck_data_signal_timestamp.bin', aDir)) then
   Exit;

Doc := nil;
try
   ReadXMLFile(Doc, ExpandFileName('procheck_data_signal.xml', aDir));
   RootNode := Doc.DocumentElement;

   F:=FileOpen(ExpandFileName('procheck_data_signal_timestamp.bin', aDir), fmOpenRead);
   len:=FileRead (F, DataBuf, DataBufLen*sizeof(double));
   FileClose(F);

   len:=len div sizeof(double);


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



