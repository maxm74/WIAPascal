unit WIA;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, ActiveX, WiaDef, WIA_LH;

type
  TWIAComponent = class;

  TWIADeviceType = (
      wiaDeviceTypeDefault          = StiDeviceTypeDefault,
      wiaDeviceTypeScanner          = StiDeviceTypeScanner,
      wiaDeviceTypeDigitalCamera    = StiDeviceTypeDigitalCamera,
      StiDeviceTypeStreamingVideo   = StiDeviceTypeStreamingVideo
  );

  { TWIASource }

  TWIASource = class(TObject)
  protected
    rOwner: TWIAComponent;
    rDeviceID: String;
    rDeviceName: String;
    rDeviceSubType: Word;
    rDeviceType: TWIADeviceType;

    pWiaDevice: IWiaItem2;

  public
    {Object being created/destroyed}
    constructor Create(AOwner: TWIAComponent; ADeviceID: String);
    destructor Destroy; override;

    property DeviceName: String read rDeviceName;
    property DeviceID: String read rDeviceID;
    property DeviceType: TWIADeviceType read rDeviceType;
    property DeviceSubType: Word read rDeviceSubType;
  end;

  {Component kinds}

  { TWIAComponent }

  TWIAComponent = class(TObject)
  protected
    pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
    lres: HResult;
    HasEnumerated: Boolean;
    rDeviceList : array of TWIASource;

    function GetSelectedSource: TWIASource;
    function GetSelectedSourceIndex: Integer;
    function GetSource(Index: Integer): TWIASource;
    function GetSourceCount: Integer;
    procedure SetSelectedSourceIndex(AValue: Integer);

    function CreateDevManager: IUnknown; virtual;

    procedure EmptyDeviceList(setZeroLength:Boolean);

    //Enumerate the avaliable devices
    function EnumerateDevices: Boolean;

  public
    constructor Create;
    destructor Destroy; override;

    {Returns a source}
    property Source[Index: Integer]: TWIASource read GetSource;

    {Returns the number of sources, after Library and Source Manager}
    {has being loaded}
    property SourceCount: Integer read GetSourceCount;

    //Selected source in a dialog
    property SelectedSourceIndex: Integer read GetSelectedSourceIndex write SetSelectedSourceIndex;

    //Selected source in a dialog
    property SelectedSource: TWIASource read GetSelectedSource;

  end;


implementation

{ TWIASource }

constructor TWIASource.Create(AOwner: TWIAComponent; ADeviceID: String);
begin
  inherited Create;

  rOwner :=AOwner;
  rDeviceID :=ADeviceID;
end;

destructor TWIASource.Destroy;
begin
  if (pWiaDevice<>nil) then pWiaDevice:=nil; //Free the Interface

  inherited Destroy;
end;

{ TWIAComponent }

function TWIAComponent.GetSelectedSource: TWIASource;
begin

end;

function TWIAComponent.GetSelectedSourceIndex: Integer;
begin

end;

function TWIAComponent.GetSource(Index: Integer): TWIASource;
begin

end;

function TWIAComponent.GetSourceCount: Integer;
begin
  if (pWIA_DevMgr=nil)
  then pWIA_DevMgr :=WIA_LH.IWiaDevMgr2(CreateDevManager);

  //Enumerate devices if needed
  if not(HasEnumerated)
  then HasEnumerated :=EnumerateDevices;

  Result := Length(rDeviceList);
end;

procedure TWIAComponent.SetSelectedSourceIndex(AValue: Integer);
begin

end;

function TWIAComponent.CreateDevManager: IUnknown;
begin
  lres :=CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, Result);
end;

procedure TWIAComponent.EmptyDeviceList(setZeroLength:Boolean);
var
   i:Integer;

begin
  for i:=Low(rDeviceList) to High(rDeviceList) do
    if (rDeviceList[i]<>nil) then FreeAndNil(rDeviceList[i]);
  if setZeroLength then SetLength(rDeviceList, 0);
end;

function TWIAComponent.EnumerateDevices: Boolean;
var
  i:integer;
  ppIEnum: IEnumWIA_DEV_INFO;
  devCount: ULONG;
  devFetched: ULONG;
  pWiaPropertyStorage: IWiaPropertyStorage;
  pPropIDS: array [0..2] of PROPID;
  pPropNames: array [0..2] of LPOLESTR;
  pPropSpec: array [0..2] of PROPSPEC;
  pPropVar: array [0..2] of PROPVARIANT;

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

        SetLength(rDeviceList, devCount);
        EmptyDeviceList(False);

        for i:=0 to devCount-1 do
        begin
          FillChar(pPropVar, Sizeof(pPropVar), 0);
          FillChar(pPropNames, Sizeof(pPropNames), 0);

          pWiaPropertyStorage :=nil;

          lres :=ppIEnum.Next(1, pWiaPropertyStorage, devFetched);

          if (lres<>S_OK)
          then Exception.Create('pWiaPropertyStorage for Device '+IntToStr(i)+' not available');

          //
          // Ask for the property values
          //
          lres := pWiaPropertyStorage.ReadMultiple(3, @pPropSpec, @pPropVar);


          lres := pWiaPropertyStorage.ReadPropertyNames(3, @pPropIDS, @pPropNames);

          if (VT_BSTR = pPropVar[0].vt)
          then rDeviceList[i] :=TWIASource.Create(Self, pPropVar[0].bstrVal)
          else Exception.Create('ID of Device '+IntToStr(i)+' not String');

          if (VT_BSTR = pPropVar[1].vt)
          then rDeviceList[i].rDeviceName :=pPropVar[1].bstrVal
          else Exception.Create('DeviceName of Device '+IntToStr(i)+' not String');

          if (VT_I4 = pPropVar[2].vt)
          then rDeviceList[i].rDeviceType :=TWiaDeviceType(pPropVar[2].iVal)
          else Exception.Create('DeviceType of Device '+IntToStr(i)+' not Integer');

          CoTaskMemFree(pPropNames[0]);
          CoTaskMemFree(pPropNames[1]);
          CoTaskMemFree(pPropNames[2]);
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

constructor TWIAComponent.Create;
begin
  inherited Create;

  HasEnumerated :=False;
  pWIA_DevMgr :=nil;
end;

destructor TWIAComponent.Destroy;
begin
  if (pWIA_DevMgr<>nil) then pWIA_DevMgr :=nil; //Free the Interface



  inherited Destroy;
end;

end.

