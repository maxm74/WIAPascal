unit wiademomain;

{$ifdef fpc}
{$mode delphi}
{$endif}

{$H+}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Variants,
  //ComObj, ActiveX, WIA_LH, WiaDef, WIA_TLB
  ComObj, ActiveX, WIA_LH, WiaDef, WIA;

type

  { TForm1 }

  TForm1 = class(TForm)
    btIntCap: TButton;
    btListChilds: TButton;
    btIntList: TButton;
    btDownload: TButton;
    btSelect: TButton;
    edDevTest: TEdit;
    edSelItemName: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Memo2: TMemo;
    procedure btIntCapClick(Sender: TObject);
    procedure btListChildsClick(Sender: TObject);
    procedure btIntListClick(Sender: TObject);
    procedure btSelectClick(Sender: TObject);
    procedure btDownloadClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { private declarations }
    //FWIA_DevMgr: WIA_TLB.DeviceManager;
    //FWIA_DevMgr: WIA_LH.IWiaDevMgr2;
    lres :HResult;
    FWia: TWIAManager;

    function DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                         lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean;

  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$ifdef FPC}
  {$R *.lfm}
{$else}
  {$R *.dfm}
{$endif}

{ TForm1 }

procedure TForm1.btIntListClick(Sender: TObject);
var
  i:integer;
  curDev: TWIADevice;
(*  curDev: IDeviceInfo;
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
*)
begin
  Memo2.Lines.Add(#13#10+' List of WIA Devices');
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
  (*
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
  *)
  for i:=0 to FWia.DevicesCount-1 do
  begin
     curDev :=FWia.Devices[i];
     if (curDev <> nil) then
     begin
       Memo2.Lines.Add('['+IntToStr(i)+']');
       Memo2.Lines.Add('  ID : '+curDev.ID);
       Memo2.Lines.Add('  Manufacturer : '+curDev.Manufacturer);
       Memo2.Lines.Add('  Name : '+curDev.Name);
       Memo2.Lines.Add('  Type : '+WIADeviceTypeStr[curDev.Type_]);
     end
     else Memo2.Lines.Add('['+IntToStr(i)+'] = NIL');
  end;
end;

procedure TForm1.btSelectClick(Sender: TObject);
var
   newIndex: Integer;

begin
  newIndex:= FWIA.SelectDeviceDialog;
  if (newIndex > -1) then
  begin
    FWIA.SelectedDeviceIndex:= newIndex;
    edDevTest.Text:= IntToStr(newIndex);
  end;
end;

procedure TForm1.btDownloadClick(Sender: TObject);
var
   test :IWiaItem2;
   curDev: TWIADevice;
   OleStrID :POleStr;
   pWiaDevice: IWiaItem2;
   pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
   c: Integer;

begin
  try
     Memo2.Lines.Add(#13#10+' Download on Device : '+edDevTest.Text);
     curDev :=FWia.Devices[StrToInt(edDevTest.Text)];

     if edSelItemName.Text <> ''
     then begin
            if curDev.SelectItem(edSelItemName.Text)
            then c:= curDev.Download('', 'WiaTest.bmp')
            else Memo2.Lines.Add('Item '+edSelItemName.Text+' NOT FOUND');
          end
     else c:= curDev.Download('', 'WiaTest.bmp');

     Memo2.Lines.Add('Item Downloaded '+IntToStr(c)+' Files');
  finally
  end;
end;

procedure TForm1.btIntCapClick(Sender: TObject);
var
   i: Integer;
   test :IWiaItem2;
   curDev: TWIADevice;
   OleStrID :POleStr;
   pWiaDevice: IWiaItem2;
   pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
   pWiaPropertyStorage: IWiaPropertyStorage;
   pPropIDS: array [0..3] of PROPID;
   pPropNames: array [0..3] of POLESTR;
   pPropSpec: array [0..3] of PROPSPEC;
   pPropVar: array [0..3] of PROPVARIANT;

begin
  try
     Memo2.Lines.Add(#13#10+' List of Capabilities for Device : '+edDevTest.Text);
     curDev :=FWia.Devices[StrToInt(edDevTest.Text)];
     test :=curDev.RootItem;
     lres :=test.QueryInterface(IID_IWiaPropertyStorage, pWiaPropertyStorage);
     //WIA_IPS_PAGE_SIZE
      pPropSpec[0].ulKind := PRSPEC_PROPID;
      pPropSpec[0].propid := WIA_IPS_MAX_HORIZONTAL_SIZE;

      pPropSpec[1].ulKind := PRSPEC_PROPID;
      pPropSpec[1].propid := WIA_IPS_MAX_VERTICAL_SIZE;

      pPropSpec[2].ulKind := PRSPEC_PROPID;
      pPropSpec[2].propid := WIA_IPS_XRES; //WIA_IPS_OPTICAL_XRES;

      pPropSpec[3].ulKind := PRSPEC_PROPID;
      pPropSpec[3].propid := WIA_IPS_YRES; //WIA_IPS_OPTICAL_XRES;

      for i:=0 to Length(pPropSpec)-1 do pPropIDS[i] := pPropSpec[i].propid;

      FillChar(pPropVar, Sizeof(pPropVar), 0);
      FillChar(pPropNames, Sizeof(pPropNames), 0);

      lres := pWiaPropertyStorage.ReadMultiple(Length(pPropSpec), @pPropSpec, @pPropVar);

      if (lres = S_OK) then
      begin
        lres := pWiaPropertyStorage.ReadPropertyNames(Length(pPropSpec), @pPropIDS, @pPropNames);

        for i:=0 to Length(pPropSpec)-1 do
        begin
          if (VT_I4 = pPropVar[0].vt)
          then Memo2.Lines.Add('   '+pPropNames[i]+': '+IntToStr(pPropVar[i].iVal));

          CoTaskMemFree(pPropNames[i]);
        end;
      end;

     pWiaPropertyStorage :=nil;
  finally
  end;
end;

procedure TForm1.btListChildsClick(Sender: TObject);
var
   i: Integer;
   curDev: TWIADevice;
   curItem: PWIAItem;

begin
  Memo2.Lines.Add(#13#10+' List of Childs for Device : '+edDevTest.Text);
  curDev :=FWia.Devices[StrToInt(edDevTest.Text)];

  for i:=0 to curDev.ItemCount-1 do
  begin
    curItem:= curDev.Items[i];
    if curItem <> nil
    then  Memo2.Lines.Add('  Name='+curItem^.Name+'  Kind='+IntToHex(curItem^.ItemType));
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  //FWIA_DevMgr :=WIA_TLB.CoDeviceManager.Create;
 // FWIA_DevMgr :=nil;
 // lres :=CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, FWIA_DevMgr);

  FWia :=TWIAManager.Create;
  FWia.OnAfterDeviceTransfer:= DeviceTransferEvent;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
 //if FWIA_DevMgr<>nil then FWIA_DevMgr :=nil;

  FWia.Free;
end;

function TForm1.DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice; lFlags: LONG;
  pWiaTransferParams: PWiaTransferParams): Boolean;
begin
  Result:= True;

  if (pWiaTransferParams <> nil) then
  Case pWiaTransferParams^.lMessage of
    WIA_TRANSFER_MSG_STATUS: begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_STATUS : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
    end;
    WIA_TRANSFER_MSG_END_OF_STREAM: begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_END_OF_STREAM : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
    end;
    WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_END_OF_TRANSFER : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
      Memo2.Lines.Add('   Downloaded='+BoolToStr(AWiaDevice.Downloaded, True));
    end;
    WIA_TRANSFER_MSG_DEVICE_STATUS: begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_DEVICE_STATUS : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
    end;
    WIA_TRANSFER_MSG_NEW_PAGE: begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_NEW_PAGE : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
    end
    else begin
      Memo2.Lines.Add('WIA_TRANSFER_MSG_'+IntToHex(pWiaTransferParams^.lMessage)+' : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
    end;
  end;
end;

end.

