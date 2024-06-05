unit WIA;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, WiaDef, WIA_LH;

type
  TWIAManager = class;

  TWIADeviceType = (
      WIADeviceTypeDefault          = StiDeviceTypeDefault,
      WIADeviceTypeScanner          = StiDeviceTypeScanner,
      WIADeviceTypeDigitalCamera    = StiDeviceTypeDigitalCamera,
      WIADeviceTypeStreamingVideo   = StiDeviceTypeStreamingVideo
  );

  { TWIADevice }

  TWIADevice = class(TObject)
  protected
    rOwner: TWIAManager;
    rIndex: Integer;
    rID,
    rManufacturer,
    rName: String;
    rType: TWIADeviceType;
    rSubType: Word;
    lres: HResult;
    pWiaDevice: IWiaItem2;

    function GetWiaDevice: IWiaItem2;

  public
    constructor Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
    destructor Destroy; override;

    property ID: String read rID;
    property Manufacturer: String read rManufacturer;
    property Name: String read rName;
    property Type_: TWIADeviceType read rType;
    property SubType: Word read rSubType;

    property WiaDevice: IWiaItem2 read GetWiaDevice;
  end;

  { TWIAManager }

  TWIAManager = class(TObject)
  protected
    pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
    lres: HResult;
    HasEnumerated: Boolean;
    rSelectedDeviceIndex: Integer;
    rDeviceList : array of TWIADevice;

    function GetSelectedDevice: TWIADevice;
    function GetSelectedDeviceIndex: Integer;
    procedure SetSelectedDeviceIndex(AValue: Integer);
    function GetDevice(Index: Integer): TWIADevice;
    function GetDevicesCount: Integer;

    function CreateDevManager: IUnknown; virtual;

    procedure EmptyDeviceList(setZeroLength:Boolean);

    //Enumerate the avaliable devices
    function EnumerateDevices: Boolean;

  public
    constructor Create;
    destructor Destroy; override;

    {Finds a matching Device index}
    function FindDevice(AID: String): Integer; overload;
    function FindDevice(Value: TWIADevice): Integer; overload;
    function FindDevice(AName: String; AManufacturer: String=''): Integer; overload;

    {Returns a Device}
    property Devices[Index: Integer]: TWIADevice read GetDevice;

    {Returns the number of Devices}
    property DevicesCount: Integer read GetDevicesCount;

    //Selected Device in a dialog
    property SelectedDeviceIndex: Integer read GetSelectedDeviceIndex write SetSelectedDeviceIndex;

    //Selected Device in a dialog
    property SelectedDevice: TWIADevice read GetSelectedDevice;

  end;

const
  WIADeviceTypeStr : array  [TWIADeviceType] of String = (
    'Default', 'Scanner', 'Digital Camera', 'Streaming Video'
  );



implementation

uses ComObj, ActiveX;

{ TWIADevice }

function TWIADevice.GetWiaDevice: IWiaItem2;
//var
//   OleStrID :BSTR;

begin
  Result :=nil;

  if (rOwner <> nil) then
  begin
    if (pWiaDevice = nil)
    then try
           //OleStrID :=SysAllocString(StringToOleStr(Self.rID));

           lres :=rOwner.pWIA_DevMgr.CreateDevice(0, StringToOleStr(Self.rID), pWiaDevice);
           if (lres = S_OK) then Result :=pWiaDevice;
         finally
           //if (OleStrID<>nil) then SysFreeString(OleStrID);
         end
    else Result :=pWiaDevice;
  end;
end;

constructor TWIADevice.Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
begin
  inherited Create;

  rOwner :=AOwner;
  rIndex :=AIndex;
  rID :=ADeviceID;
  pWiaDevice :=nil;
end;

destructor TWIADevice.Destroy;
begin
  if (pWiaDevice<>nil) then pWiaDevice:=nil; //Free the Interface

  inherited Destroy;
end;

{ TWIAManager }

function TWIAManager.GetSelectedDevice: TWIADevice;
begin
  if (rSelectedDeviceIndex>=0) and (rSelectedDeviceIndex<Length(rDeviceList))
  then Result :=rDeviceList[rSelectedDeviceIndex]
  else Result :=nil;
end;

function TWIAManager.GetSelectedDeviceIndex: Integer;
begin
   Result :=rSelectedDeviceIndex;
end;

procedure TWIAManager.SetSelectedDeviceIndex(AValue: Integer);
begin
  if (AValue<>rSelectedDeviceIndex) and (AValue>=0) and (AValue<Length(rDeviceList)) then
  begin
    if (rDeviceList[AValue]<>nil)
    then rSelectedDeviceIndex :=AValue
    else rSelectedDeviceIndex :=-1;
  end;
end;

function TWIAManager.GetDevice(Index: Integer): TWIADevice;
begin
  if (Index>=0) and (Index<Length(rDeviceList))
  then Result :=rDeviceList[Index]
  else Result :=nil;
end;

function TWIAManager.GetDevicesCount: Integer;
begin
  if (pWIA_DevMgr=nil)
  then pWIA_DevMgr :=WIA_LH.IWiaDevMgr2(CreateDevManager);

  //Enumerate devices if needed
  if not(HasEnumerated)
  then HasEnumerated :=EnumerateDevices;

  Result := Length(rDeviceList);
end;

function TWIAManager.CreateDevManager: IUnknown;
begin
  lres :=CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, Result);
end;

procedure TWIAManager.EmptyDeviceList(setZeroLength:Boolean);
var
   i:Integer;

begin
  for i:=Low(rDeviceList) to High(rDeviceList) do
    if (rDeviceList[i]<>nil) then FreeAndNil(rDeviceList[i]);
  if setZeroLength then SetLength(rDeviceList, 0);
end;

function TWIAManager.EnumerateDevices: Boolean;
var
  i:integer;
  ppIEnum: IEnumWIA_DEV_INFO;
  devCount: ULONG;
  devFetched: ULONG;
  pWiaPropertyStorage: IWiaPropertyStorage;
  //pPropIDS: array [0..3] of PROPID;
  //pPropNames: array [0..3] of LPOLESTR;
  pPropSpec: array [0..3] of PROPSPEC;
  pPropVar: array [0..3] of PROPVARIANT;

begin
  Result :=False;
  SetLength(rDeviceList, 0);

  try
    if pWIA_DevMgr<>nil then
    begin
      lres :=pWIA_DevMgr.EnumDeviceInfo(WIA_DEVINFO_ENUM_ALL, ppIEnum);
      if (lres=S_OK) and (ppIEnum<>nil) then
      begin
        lres :=ppIEnum.GetCount(devCount);

        if (lres<>S_OK)
        then Exception.Create('Number of WIA Devices not available');

        //
        // Define which properties you want to read:
        // Device ID.  This is what you would use to create
        // the device.
        //
        pPropSpec[0].ulKind := PRSPEC_PROPID;
        pPropSpec[0].propid := WIA_DIP_DEV_ID;
        //pPropIDS[0] :=WIA_DIP_DEV_ID;

        //
        // Device Manufacturer
        //
        pPropSpec[1].ulKind := PRSPEC_PROPID;
        pPropSpec[1].propid := WIA_DIP_VEND_DESC;
        //pPropIDS[1] :=WIA_DIP_VEND_DESC;

        //
        // Device Name
        //
        pPropSpec[2].ulKind := PRSPEC_PROPID;
        pPropSpec[2].propid := WIA_DIP_DEV_NAME;
        //pPropIDS[2] :=WIA_DIP_DEV_NAME;

        //
        // Device Type
        //
        pPropSpec[3].ulKind := PRSPEC_PROPID;
        pPropSpec[3].propid := WIA_DIP_DEV_TYPE;
        //pPropIDS[3] :=WIA_DIP_DEV_TYPE;

        SetLength(rDeviceList, devCount);
        EmptyDeviceList(False);

        pWiaPropertyStorage :=nil;

        for i:=0 to devCount-1 do
        begin
          FillChar(pPropVar, Sizeof(pPropVar), 0);
          //FillChar(pPropNames, Sizeof(pPropNames), 0);

          lres :=ppIEnum.Next(1, pWiaPropertyStorage, devFetched);

          if (lres<>S_OK)
          then Exception.Create('pWiaPropertyStorage for Device '+IntToStr(i)+' not available');

          //
          // Ask for the property values
          //
          lres := pWiaPropertyStorage.ReadMultiple(4, @pPropSpec, @pPropVar);


         // lres := pWiaPropertyStorage.ReadPropertyNames(4, @pPropIDS, @pPropNames);

          if (VT_BSTR = pPropVar[0].vt)
          then rDeviceList[i] :=TWIADevice.Create(Self, i, pPropVar[0].bstrVal)
          else Exception.Create('ID of Device '+IntToStr(i)+' not String');

          if (VT_BSTR = pPropVar[1].vt)
          then rDeviceList[i].rManufacturer :=pPropVar[1].bstrVal
          else Exception.Create('Manufacturer of Device '+IntToStr(i)+' not String');

          if (VT_BSTR = pPropVar[2].vt)
          then rDeviceList[i].rName :=pPropVar[2].bstrVal
          else Exception.Create('Name of Device '+IntToStr(i)+' not String');

          if (VT_I4 = pPropVar[3].vt)
          then rDeviceList[i].rType :=TWiaDeviceType(pPropVar[3].iVal)
          else Exception.Create('DeviceType of Device '+IntToStr(i)+' not Integer');

          pWiaPropertyStorage :=nil;

         (*CoTaskMemFree(pPropNames[0]);
          CoTaskMemFree(pPropNames[1]);
          CoTaskMemFree(pPropNames[2]);
          CoTaskMemFree(pPropNames[3]);*)
        end;

        ppIEnum :=nil;
      end;
    end;

    Result :=True;

  except
    EmptyDeviceList(True);
    Result :=False;
  end;
end;

constructor TWIAManager.Create;
begin
  inherited Create;

  HasEnumerated :=False;
  pWIA_DevMgr :=nil;
  rSelectedDeviceIndex :=-1;
end;

destructor TWIAManager.Destroy;
begin
  EmptyDeviceList(True);
  if (pWIA_DevMgr<>nil) then pWIA_DevMgr :=nil; //Free the Interface

  inherited Destroy;
end;

function TWIAManager.FindDevice(AID: String): Integer;
var
   i      :Integer;
   curDev :TWIADevice;

begin
    Result :=-1;
    for i:=0 to DevicesCount-1 do
    begin
      curDev :=Devices[i];
      if (curDev <> nil) and
         (curDev.ID = AID)
      then begin Result:=i; break; end;
    end;
end;

function TWIAManager.FindDevice(Value: TWIADevice): Integer;
var
   i      :Integer;

begin
  Result :=-1;
  if (Value <> nil)
  then for i:=0 to DevicesCount-1 do
         if (Devices[i] = Value) then begin Result:=i; break; end;
end;

function TWIAManager.FindDevice(AName: String; AManufacturer: String): Integer;
var
   i      :Integer;
   curDev :TWIADevice;

begin
    Result :=-1;
    for i:=0 to DevicesCount-1 do
    begin
      curDev:=Devices[i];
      { #todo -oMaxM : if there is more identical device? }
      if (curDev <> nil) and
         (curDev.Name = AName) and
         ((AManufacturer <> '') and (curDev.Manufacturer = AManufacturer))
      then begin Result:=i; break; end;
    end;
end;

end.

