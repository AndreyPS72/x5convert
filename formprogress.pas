unit formprogress;

{$MODE DelphiUnicode}{$CODEPAGE UTF8}{$H+}

{$OPTIMIZATION OFF, NOREGVAR, UNCERTAIN, NOSTACKFRAME, NOPEEPHOLE, NOLOOPUNROLL, NOTAILREC, NOORDERFIELDS, NOFASTMATH, NOREMOVEEMPTYPROCS, NOCSE, NODFA} //debug Для отладки

interface

uses
    Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, Buttons,
    ExtCtrls;

type

    { TFormMain }

    TFormMain = class(TForm)
      BitBtnStop: TBitBtn;
      OpenDialogFile: TOpenDialog;
      pbProgress: TProgressBar;
      Timer1: TTimer;
      procedure BitBtnStopClick(Sender: TObject);
      procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
      procedure Timer1Timer(Sender: TObject);
    private

    public

    end;

var
    FormMain: TFormMain;

const StopProcess : boolean = false;

implementation
uses converter;

{$R *.lfm}

{ TFormMain }

procedure TFormMain.BitBtnStopClick(Sender: TObject);
begin
  StopProcess := true;
end;

procedure TFormMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  StopProcess := true;
  Application.ProcessMessages;
  CloseAction:=caFree;
end;


procedure TFormMain.Timer1Timer(Sender: TObject);
var res: integer;
begin

  Timer1.OnTimer:=nil;
  res:=1;

  if Application.ParamCount>0 then begin
     res:=ConvertX5File(Application.Params[1]);
  end else begin
    if OpenDialogFile.Execute then begin
       res:=ConvertX5File(OpenDialogFile.FileName);
    end
  end;

Application.Terminate;
while not Application.Terminated do
      Application.ProcessMessages;
Halt(res);
end;

end.

