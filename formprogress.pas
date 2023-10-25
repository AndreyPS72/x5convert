unit formprogress;

{$MODE DelphiUnicode}{$CODEPAGE UTF8}{$H+}

interface

uses
    Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, Buttons;

type

    { TFormMain }

    TFormMain = class(TForm)
      BitBtnStop: TBitBtn;
      pbProgress: TProgressBar;
      procedure BitBtnStopClick(Sender: TObject);
      procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
      procedure FormShow(Sender: TObject);
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

procedure TFormMain.FormShow(Sender: TObject);
begin

  if Application.ParamCount>0 then begin
//     Halt(ConvertX5File(Application.Params[1]));
     ConvertX5File(Application.Params[1])
  end;
end;

end.

