program x5convert;

{$MODE DelphiUnicode}{$CODEPAGE UTF8}{$H+}

uses
    {$IFDEF UNIX}
    cthreads,
    {$ENDIF}
    {$IFDEF HASAMIGA}
    athreads,
    {$ENDIF}
    Interfaces, // this includes the LCL widgetset
    Forms
    , formprogress, converter, rFFT
    ;

{$R *.res}

begin
    {$if declared(UseHeapTrace)}
      GlobalSkipIfNoLeaks := True; // Теперь окно с выводом heaprtc будет появляться только тогда, когда будет обнаружена утечка памяти
    {$endIf}

    RequireDerivedFormResource:=True;
    Application.Scaled:=True;
    Application.Initialize;
    Application.CreateForm(TFormMain, FormMain);
    Application.Run;
end.

