unit wiademomain;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Variants, ComObj, ActiveX, WIA_LH, WiaDef, WIA_TLB;

type

  { TForm1 }

  TForm1 = class(TForm)
    btIntCap: TButton;
    btIntCapture: TButton;
    btIntList: TButton;
    edDevTest: TEdit;
    Label1: TLabel;
    Memo2: TMemo;
    procedure btIntCapClick(Sender: TObject);
    procedure btIntListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { private declarations }
    //FWIA_DevMgr: WIA_TLB.DeviceManager;
    FWIA_DevMgr: WIA_LH.IWiaDevMgr2;
    lres :HResult;

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
  ppIEnum: IEnumWIA_DEV_INFO;
  devCount: ULONG;
  devFetched: ULONG;
  pWiaPropertyStorage: IWiaPropertyStorage;
  pPropIDS: array [0..2] of PROPID;
  pPropNames: array [0..2] of LPOLESTR;
  pPropSpec: array [0..2] of PROPSPEC;
  pPropVar: array [0..2] of PROPVARIANT;

begin
 (*
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
*)
  if FWIA_DevMgr<>nil then
  begin
    lres :=FWIA_DevMgr.EnumDeviceInfo(WIA_DEVINFO_ENUM_ALL, ppIEnum);
    if (lres=S_OK) and (ppIEnum<>nil) then
    begin
      lres :=ppIEnum.GetCount(devCount);
      if (lres=S_OK) then
      begin
        Memo2.Lines.Add('Number of WIA Devices ='+IntToStr(devCount));
      end;

      //
      // Define which properties you want to read:
      // Device ID.  This is what you would use to create
      // the device.
      //
      pPropSpec[0].ulKind := PRSPEC_PROPID;
      pPropSpec[0].propid := WIA_DIP_DEV_ID;
      pPropIDS[0] :=WIA_DIP_DEV_ID;

      //
      // Device Name
      //
      pPropSpec[1].ulKind := PRSPEC_PROPID;
      pPropSpec[1].propid := WIA_DIP_DEV_NAME;
      pPropIDS[1] :=WIA_DIP_DEV_NAME;

      //
      // Device description
      //
      pPropSpec[2].ulKind := PRSPEC_PROPID;
      pPropSpec[2].propid := WIA_DIP_DEV_TYPE;
      pPropIDS[2] :=WIA_DIP_DEV_TYPE;

      for i:=1 to devCount do
      begin
        FillChar(pPropVar, Sizeof(pPropVar), 0);
        FillChar(pPropNames, Sizeof(pPropNames), 0);

        pWiaPropertyStorage :=nil;

        lres :=ppIEnum.Next(1, pWiaPropertyStorage, devFetched);

        //
        // Ask for the property values
        //
        lres := pWiaPropertyStorage.ReadMultiple(3, @pPropSpec, @pPropVar);


        lres := pWiaPropertyStorage.ReadPropertyNames(3, @pPropIDS, @pPropNames);

        Memo2.Lines.Add('['+IntToStr(i)+']');

        if (VT_BSTR = pPropVar[0].vt)
        then Memo2.Lines.Add('   '+pPropNames[0]+': '+pPropVar[0].bstrVal);

        if (VT_BSTR = pPropVar[1].vt)
        then Memo2.Lines.Add('   '+pPropNames[1]+': '+pPropVar[1].bstrVal);

        if (VT_I4 = pPropVar[2].vt)
        then Memo2.Lines.Add('   '+pPropNames[2]+': '+IntToStr(pPropVar[2].iVal));

        CoTaskMemFree(pPropNames[0]);
        CoTaskMemFree(pPropNames[1]);
        CoTaskMemFree(pPropNames[2]);
      end;

      ppIEnum :=nil;
    end;
  end;
end;

procedure TForm1.btIntCapClick(Sender: TObject);
begin

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  //FWIA_DevMgr :=WIA_TLB.CoDeviceManager.Create;
  FWIA_DevMgr :=nil;
  lres :=CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, FWIA_DevMgr);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
 if FWIA_DevMgr<>nil then FWIA_DevMgr :=nil;
end;

end.

