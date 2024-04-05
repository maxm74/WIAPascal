unit wiademomain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Variants, WIA_TLB, WiaDef;

type

  { TForm1 }

  TForm1 = class(TForm)
    btIntCap: TButton;
    btIntCapture: TButton;
    btIntList: TButton;
    edDevTest: TEdit;
    Label1: TLabel;
    Memo2: TMemo;
    procedure btIntListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
    FWIA_DevMgr: WIA_TLB.DeviceManager;

  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.btIntListClick(Sender: TObject);
var
  i:integer;
  Scanner: Device;
  Picture: IItem;
  Image: OleVariant;
//  InVar: OLEVariant;
//  StringVar: OleVariant;
  OutVar: OLEVariant;
  AImage: IImageFile;
  ReturnString: string;
  astr:String;

begin
  // List of devices is a 1 based array
  Memo2.Lines.Add('Number of WIA Devices ='+IntToStr(FWIA_DevMgr.DeviceInfos.Count));

  for i:=1 to FWIA_DevMgr.DeviceInfos.Count do
  begin
    //InVar:=i;
    //StringVar:='Type';
    //astr :=utf8encode(FWIA_DevMgr.DeviceInfos[i].Properties['Type'].get_Value);
    astr :='['+IntToStr(i)+'] Type: '+IntToStr(FWIA_DevMgr.DeviceInfos[i].type_)+#13#10;

    astr :=astr+'    ID: '+FWIA_DevMgr.DeviceInfos[i].DeviceID+#13#10;
    //StringVar:='Name';
    astr :=astr+'    Name: '+utf8encode(FWIA_DevMgr.DeviceInfos[i].Properties['Name'].get_Value)+#13#10;

    Memo2.Lines.Add(astr);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  FWIA_DevMgr :=WIA_TLB.CoDeviceManager.Create;
end;

end.

