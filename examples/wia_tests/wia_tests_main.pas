(******************************************************************************
*                FreePascal \ Delphi WIA Implementation                       *
*                                                                             *
*  FILE: wia_tests_main.pas                                                   *
*                                                                             *
*  VERSION:     1.0.1                                                         *
*                                                                             *
*  DESCRIPTION:                                                               *
*    WIA Test Form.                                                           *
*    This is a form to do internal tests on the scanner's response            *
*    not a full scanning application, see the DigIt project on my             *
*    GitHub repository for that.                                              *
*                                                                             *
*******************************************************************************
*                                                                             *
*  (c) 2025 Massimo Magnano                                                   *
*                                                                             *
*  See changelog.txt for Change Log                                           *
*                                                                             *
*******************************************************************************)
unit wia_tests_main;

{$ifdef fpc}
{$mode delphi}
{$endif}

{$H+}
{$POINTERMATH ON}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Spin, Variants,
  //WIA_TLB,
  ComObj, ActiveX, WIA_LH, WiaDef, WIA;

type

  { TFormWIATests }

  TFormWIATests = class(TForm)
    btIntCap: TButton;
    btIntList: TButton;
    btDownload: TButton;
    btListChilds: TButton;
    btSelect: TButton;
    edDevTest: TEdit;
    edDPI: TSpinEdit;
    edSelItemName: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
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
  FormWIATests: TFormWIATests;

implementation

{$ifdef FPC}
  {$R *.lfm}
{$else}
  {$R *.dfm}
{$endif}

{ TFormWIATests }

procedure TFormWIATests.btIntListClick(Sender: TObject);
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
       Memo2.Lines.Add('  Type : '+WIADeviceType(curDev.Type_));
     end
     else Memo2.Lines.Add('['+IntToStr(i)+'] = NIL');
  end;
end;

procedure TFormWIATests.btSelectClick(Sender: TObject);
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

procedure TFormWIATests.btDownloadClick(Sender: TObject);
var
   curDev: TWIADevice;
   c: Integer;
   path: String;

begin
  try
     Memo2.Lines.Add(#13#10+' Download on Device : '+edDevTest.Text);
     curDev :=FWia.Devices[StrToInt(edDevTest.Text)];

      if edSelItemName.Text <> ''
      then if not(curDev.SelectItem(edSelItemName.Text)) then
           begin
             Memo2.Lines.Add('Item '+edSelItemName.Text+' NOT FOUND');
             exit;
           end;

      path:= ExtractFilePath(ParamStr(0))+'download';
      c:= edDPI.Value;
      curDev.SetProperty(WIA_IPS_XRES, VT_INT, c);
      curDev.SetProperty(WIA_IPS_YRES, VT_INT, c);
      c:= curDev.Download(path, 'WiaTest', '.bmp');
      Memo2.Lines.Add('Item Downloaded '+IntToStr(c)+' Files'+#13#10+'  on '+path);
  finally
  end;
end;

procedure TFormWIATests.btIntCapClick(Sender: TObject);
var
   curDev: TWIADevice;
   i, k: Integer;
   pWiaPropertyStorage: IWiaPropertyStorage;
   (*
   test :IWiaItem2;
   OleStrID :POleStr;
   pWiaDevice: IWiaItem2;
   pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
   pPropIDS: array [0..3] of PROPID;
   pPropNames: array [0..3] of POLESTR;
   pPropSpec: array [0..3] of PROPSPEC;
   pPropVar: array [0..3] of PROPVARIANT;
   pFlags: array[0..3] of ULONG;
   sPropNames: array [0..3] of String;
*)
   fStr: String;
   propFlags: TWIAPropertyFlags;
   propType: TVarType;
   propValue,
   propDefaultValue: Integer;
   propListValues: TArrayInteger;

   procedure WriteMemo;
   var
      i: Integer;

   begin
     if not(propFlags = []) then
     begin
      fStr:='  Flags: ';
      if (WIAProp_READ in propFlags) then fStr:= fStr+'R,';
      if (WIAProp_WRITE in propFlags) then fStr:= fStr+'W,';
      if (WIAProp_SYNC_REQUIRED in propFlags) then fStr:= fStr+'S,';
      if (WIAProp_NONE in propFlags) then fStr:= fStr+'N,';
      if (WIAProp_RANGE in propFlags) then fStr:= fStr+'RANGE,';
      if (WIAProp_LIST in propFlags) then fStr:= fStr+'LIST,';
      if (WIAProp_FLAG in propFlags) then fStr:= fStr+'FLAG,';
      if (WIAProp_CACHEABLE in propFlags) then fStr:= fStr+'C';

      Memo2.Lines.Add(fStr+#13#10'  DefaultValue: '+IntToStr(propDefaultValue));
      Memo2.Lines.Add('  CurrentValue: '+IntToStr(propValue));
      fStr:='  Values: ';
      for i:=0 to Length(propListValues)-1 do fStr:= fStr+IntToStr(propListValues[i])+', ';
     end
     else fStr:='  Error';

     Memo2.Lines.Add(fStr);
   end;

begin
  try
     Memo2.Lines.Add(#13#10+' List of Capabilities for Device : '+edDevTest.Text);
     curDev :=FWia.Devices[StrToInt(edDevTest.Text)];

     if edSelItemName.Text <> ''
     then curDev.SelectItem(edSelItemName.Text);

     pWiaPropertyStorage:= curDev.SelectedPropertiesIntf;

     if pWiaPropertyStorage=nil then
     begin
      Memo2.Lines.Add('Item '+edSelItemName.Text+' NOT FOUND');
      exit;
     end;

(* Direct Code
//     pWiaPropertyStorage:= curDev.RootProperties;

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

      Memo2.Lines.Add(#13#10'SizeOf(TPROPVARIANT) = '+IntToStr(SizeOf(TPROPVARIANT)));

      lres := pWiaPropertyStorage.ReadPropertyNames(Length(pPropSpec), @pPropIDS, @pPropNames);
      if (lres = S_OK) then
      begin
        for i:=0 to Length(pPropSpec)-1 do
        begin
          sPropNames[i]:= pPropNames[i];

          CoTaskMemFree(pPropNames[i]);
        end;
      end;

      lres:= pWiaPropertyStorage.GetPropertyAttributes(Length(pPropSpec), @pPropSpec, @pFlags, @pPropVar);
      if (lres = S_OK) then
      begin
        Memo2.Lines.Add(#13#10'PropertyAttributes:');

        for i:=0 to Length(pPropSpec)-1 do
        begin
          fStr:=' '+sPropNames[i]+' ';

          if (pFlags[i] and WIA_PROP_READ <> 0) then fStr:= fStr+'R,';
          if (pFlags[i] and WIA_PROP_WRITE <> 0) then fStr:= fStr+'W,';
          if (pFlags[i] and WIA_PROP_SYNC_REQUIRED <> 0) then fStr:= fStr+'S,';
          if (pFlags[i] and WIA_PROP_NONE <> 0) then fStr:= fStr+'N,';
          if (pFlags[i] and WIA_PROP_RANGE <> 0) then fStr:= fStr+'RANGE,';
          if (pFlags[i] and WIA_PROP_LIST <> 0) then fStr:= fStr+'LIST,';
          if (pFlags[i] and WIA_PROP_FLAG <> 0) then fStr:= fStr+'F,';
          if (pFlags[i] and WIA_PROP_CACHEABLE <> 0) then fStr:= fStr+'C,';

          fStr:=fStr+': ';

          if (pPropVar[i].vt and VT_VECTOR <> 0)
          then begin
               fStr:= fStr+' (Vector of VT_I4) = (';
               Case (pPropVar[i].vt and not(VT_VECTOR)) of
               VT_I4: begin
                        for k:=0 to  pPropVar[i].cal.cElems-1 do
                         fStr:= fStr+IntToStr(longint(pPropVar[i].cal.pElems[k]))+',';
                      end;
               end;
               fStr:= fStr+')';
               end
          else begin

          end;

          Memo2.Lines.Add(fStr);
        end;
      end;

      lres := pWiaPropertyStorage.ReadMultiple(Length(pPropSpec), @pPropSpec, @pPropVar);
      if (lres = S_OK) then
      begin
        Memo2.Lines.Add(#13#10'Property Values:');

        for i:=0 to Length(pPropSpec)-1 do
        begin
          fStr:=' '+sPropNames[i]+' ';

          Case pPropVar[i].vt of
          VT_I4: fStr:= fStr+' (VT_I4) = '+IntToStr(pPropVar[i].iVal);
          end;

          Memo2.Lines.Add(fStr);
        end;
      end;

//     pWiaPropertyStorage :=nil;
*)
  //Class Code
(*  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_DPS_PAGE_SIZE):');
  propFlags:= curDev.GetProperty(WIA_DPS_PAGE_SIZE, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;
*)
  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_IPS_XRES):');
  propFlags:= curDev.GetProperty(WIA_IPS_XRES, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;

  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_IPA_DEPTH):');
  propFlags:= curDev.GetProperty(WIA_IPA_DEPTH, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;

  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_IPA_DATATYPE):');
  propFlags:= curDev.GetProperty(WIA_IPA_DATATYPE, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;

  //WIA_IPS_BRIGHTNESS for Range Test
  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_IPS_BRIGHTNESS):');
  propFlags:= curDev.GetProperty(WIA_IPS_BRIGHTNESS, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;

  //WIA_IPS_DOCUMENT_HANDLING_SELECT for Flag/Error Test
  Memo2.Lines.Add(#13#10'PropertyAttributes (WIA_IPA_ACCESS_RIGHTS):');
  propFlags:= curDev.GetProperty(WIA_IPA_ACCESS_RIGHTS, propType, propValue, propDefaultValue, propListValues);
  WriteMemo;

  finally
  end;
end;

procedure TFormWIATests.btListChildsClick(Sender: TObject);
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
    then  Memo2.Lines.Add('  Name='+curItem^.Name+'  Kind='+IntToHex(Integer(curItem^.ItemType)));
  end;
end;

procedure TFormWIATests.FormCreate(Sender: TObject);
begin
  //FWIA_DevMgr :=WIA_TLB.CoDeviceManager.Create;
 // FWIA_DevMgr :=nil;
 // lres :=CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, FWIA_DevMgr);

  FWia :=TWIAManager.Create;
  FWia.OnAfterDeviceTransfer:= DeviceTransferEvent;
end;

procedure TFormWIATests.FormDestroy(Sender: TObject);
begin
 //if FWIA_DevMgr<>nil then FWIA_DevMgr :=nil;

  if (FWIA<>nil) then FWia.Free;
end;

function TFormWIATests.DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice; lFlags: LONG;
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

