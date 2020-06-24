unit MainUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, VISSIMLIB_TLB,
  Vcl.ExtCtrls;

type
  TMainForm = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Panel1: TPanel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
    FVissim: IVissim;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses
  Vissim.ComProvider.Utilities;

{$R *.dfm}

procedure TMainForm.Button1Click(Sender: TObject);
begin
  FVissim := CoVissim.Create;
end;

procedure TMainForm.Button2Click(Sender: TObject);
begin
  ShowVissim(FVissim);
end;

procedure TMainForm.Button3Click(Sender: TObject);
begin
  HideVissim(FVissim);
end;

procedure TMainForm.Button4Click(Sender: TObject);
begin
  var vissimHWND := GetVissimMainWindowHandle(FVissim);

  if vissimHWND <> 0 then
    WinApi.Windows.SetParent(vissimHWND, Panel1.Handle);

end;

end.
