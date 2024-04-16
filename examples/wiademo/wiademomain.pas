unit wiademomain;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Variants, ComObj, WIA_LH, WiaDef, WIA_TLB;

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
    procedure FormDestroy(Sender: TObject);
  private
    { private declarations }
    FWIA_DevMgr: WIA_TLB.DeviceManager;
    //FWIA_DevMgr: WIA_LH.IWiaDevMgr2;

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
  curDev: IDeviceInfo;
  Picture: IItem;
  Image: OleVariant;
  InVar: OLEVariant;
  StringVar: OleVariant;
  OutVar: OLEVariant;
  AImage: IImageFile;
  ReturnString: string;
  astr:String;
//    lres :HResult;
//    ppIEnum: IEnumWIA_DEV_INFO;
//    devCount: ULONG;

begin
  // List of devices is a 1 based array
  Memo2.Lines.Add('Number of WIA Devices ='+IntToStr(FWIA_DevMgr.DeviceInfos.Count));

  for i:=1 to FWIA_DevMgr.DeviceInfos.Count do
  begin
    //InVar:=i;
    //StringVar:='Type';
    //astr :=utf8encode(FWIA_DevMgr.DeviceInfos[i].Properties['Type'].get_Value);

    curDev :=FWIA_DevMgr.DeviceInfos[i];

    astr :='['+IntToStr(i)+'] Type: '+IntToStr(curDev.type_)+#13#10;

    astr :=astr+'    ID: '+curDev.DeviceID+#13#10;
    //StringVar:='Name';
    astr :=astr+'    Name: '+utf8encode(curDev.Properties['Name'].get_Value)+#13#10;

    Memo2.Lines.Add(astr);
  end;

(*  if FWIA_DevMgr<>nil then
  begin
    lres :=FWIA_DevMgr.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL, ppIEnum);
    if (lres=S_OK) and (ppIEnum<>nil) then
    begin
      lres :=ppIEnum.GetCount(devCount);
      if (lres=S_OK) then
      begin
        Memo2.Lines.Add('Number of WIA Devices ='+IntToStr(devCount));
      end;

      ppIEnum._Release;
    end;
  end; *)
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  FWIA_DevMgr :=WIA_TLB.CoDeviceManager.Create;
  //FWIA_DevMgr :=CreateComObject(IID_IWiaDevMgr) as IWiaDevMgr;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
 // if FWIA_DevMgr<>nil then FWIA_DevMgr._Release;
end;

end.

