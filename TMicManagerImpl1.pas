unit TMicManagerImpl1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, TMicManagerProj1_TLB, ExtCtrls, MPlayer, StdCtrls;

type
  TTMicManagerX = class(TActiveForm, ITMicManagerX)
    Button1: TButton;
    Button3: TButton;
    Button4: TButton;
    MPInput: TMediaPlayer;
    TimChecker: TTimer;
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure TimCheckerTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure MPInputClick(Sender: TObject; Button: TMPBtnType;
      var DoDefault: Boolean);
  private
    { Private declarations }
    FEvents: ITMicManagerXEvents;
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_BiDiMode: TxBiDiMode; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: OLE_COLOR; safecall;
    function Get_Cursor: Smallint; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_Font: IFontDisp; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_Visible: WordBool; safecall;
    procedure _Set_Font(const Value: IFontDisp); safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_BiDiMode(Value: TxBiDiMode); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    procedure Set_Cursor(Value: Smallint); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_Font(var Value: IFontDisp); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
  public
    { Public declarations }
    Freq:longint;//=8000;
    rer:TFileStream;
    ElRead:longint;
    OneSec:array of longint;
    procedure Initialize; override;
  end;

implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TTMicManagerX }

procedure TTMicManagerX.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_TMicManagerXPage); }
end;

procedure TTMicManagerX.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as ITMicManagerXEvents;
end;

procedure TTMicManagerX.Initialize;
begin
  inherited Initialize;
  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
end;

function TTMicManagerX.Get_Active: WordBool;
begin
  Result := Active;
end;

function TTMicManagerX.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TTMicManagerX.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TTMicManagerX.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TTMicManagerX.Get_BiDiMode: TxBiDiMode;
begin
  Result := Ord(BiDiMode);
end;

function TTMicManagerX.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TTMicManagerX.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TTMicManagerX.Get_Cursor: Smallint;
begin
  Result := Smallint(Cursor);
end;

function TTMicManagerX.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TTMicManagerX.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TTMicManagerX.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TTMicManagerX.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TTMicManagerX.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TTMicManagerX.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TTMicManagerX.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TTMicManagerX.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TTMicManagerX.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TTMicManagerX.Get_Visible: WordBool;
begin
  Result := Visible;
end;

procedure TTMicManagerX._Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TTMicManagerX.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TTMicManagerX.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TTMicManagerX.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TTMicManagerX.Set_BiDiMode(Value: TxBiDiMode);
begin
  BiDiMode := TBiDiMode(Value);
end;

procedure TTMicManagerX.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TTMicManagerX.Set_Color(Value: OLE_COLOR);
begin
  Color := TColor(Value);
end;

procedure TTMicManagerX.Set_Cursor(Value: Smallint);
begin
  Cursor := TCursor(Value);
end;

procedure TTMicManagerX.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TTMicManagerX.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TTMicManagerX.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TTMicManagerX.Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TTMicManagerX.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TTMicManagerX.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TTMicManagerX.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TTMicManagerX.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TTMicManagerX.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TTMicManagerX.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

procedure TTMicManagerX.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TTMicManagerX.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TTMicManagerX.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TTMicManagerX.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TTMicManagerX.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TTMicManagerX.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TTMicManagerX.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TTMicManagerX.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TTMicManagerX.Button3Click(Sender: TObject);
begin
//MPInput.Save;
end;

procedure TTMicManagerX.Button4Click(Sender: TObject);
begin
MPInput.Close;
//MPInput.Destroy;
//DeleteFile('C:\windows\Personal\Yuriy\Work\MicState\*.tmp');
rer.Destroy;
Close;
end;

procedure TTMicManagerX.TimCheckerTimer(Sender: TObject);
begin
//MPInput.FileName
//MPInput.Stop;
MPInput.Save;
//MPInput.Update;
rer.Position:=rer.Size-Freq*2{sizeof(longint)};
rer.read(OneSec,Freq*2{sizeof(longint)});
MPInput.EndPos:=MPInput.StartPos;
//MPInput.Save;
MPInput.StartRecording;
//DeleteFile('C:\windows\Personal\Yuriy\Work\MicState\*.tmp');
end;

procedure TTMicManagerX.FormCreate(Sender: TObject);
begin
MPInput.DeviceType:=dtWaveAudio;
rer:=TFileStream.Create(MPInput.FileName,fmOpenRead or fmShareDenyNone);
end;

procedure TTMicManagerX.Button1Click(Sender: TObject);
//var
{   hWaveIn:integer;
   lpuDeviceID:PUINT;
   lpWaveInHdr:PWaveHdr;
   uSize:Cardinal;}
begin
//Edit1.Text:=IntToStr(waveInGetNumDevs());
//waveInReset(hWaveIn);
{waveInGetID(hWaveIn, lpuDeviceID);}
//waveInStart(hWaveIn);
//waveInOpen(
//waveInAddBuffer(hWaveIn,lpWaveInHdr,uSize);

end;

procedure TTMicManagerX.MPInputClick(Sender: TObject; Button: TMPBtnType;
  var DoDefault: Boolean);
var
   hWaveIn:integer;
{   lpuDeviceID:PUINT;
   lpWaveInHdr:PWaveHdr;
   uSize:Cardinal;}
begin
case Button of
btPlay:     hWaveIn:=1;
btRecord:   hWaveIn:=2;
btStop:     hWaveIn:=3;
end;

end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TTMicManagerX,
    Class_TMicManagerX,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
