unit WIA;

{$H+}

interface

uses
  Windows, Classes, SysUtils, ComObj, ActiveX, WiaDef, WIA_LH;

type
  TWIAManager = class;

  TWIADeviceType = (
      WIADeviceTypeDefault          = StiDeviceTypeDefault,
      WIADeviceTypeScanner          = StiDeviceTypeScanner,
      WIADeviceTypeDigitalCamera    = StiDeviceTypeDigitalCamera,
      WIADeviceTypeStreamingVideo   = StiDeviceTypeStreamingVideo
  );

  { TWIADevice }

  TWIADevice = class(TNoRefCountObject, IWiaTransferCallback)
  protected
    rOwner: TWIAManager;
    rIndex: Integer;
    rID,
    rManufacturer,
    rName: String;
    rType: TWIADeviceType;
    rSubType: Word;
    rVersion,
    rVersionSub: Integer;
    lres: HResult;
    pWiaRootItem: IWiaItem2;
    StreamDestination: TFileStream;
    StreamAdapter: TStreamAdapter;
    DownloadComplete: Boolean;

    function GetWiaRootItem: IWiaItem2;
    function CreateDestinationStream(bstrItemName: String; var ppDestination: IStream): HRESULT;

  public
    function TransferCallback(lFlags: LONG;
                              pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;

    function GetNextStream(lFlags: LONG;
                           bstrItemName,
                           bstrFullItemName: BSTR;
                           out ppDestination: IStream): HRESULT; stdcall;

    constructor Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
    destructor Destroy; override;

    function Download(const childIndex: Integer = 0): Boolean; virtual;

    property ID: String read rID;
    property Manufacturer: String read rManufacturer;
    property Name: String read rName;
    property Type_: TWIADeviceType read rType;
    property SubType: Word read rSubType;

    property WiaRootItem: IWiaItem2 read GetWiaRootItem;
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

    //Clears the list of sources
    procedure ClearDeviceList;

    //Display a dialog to let the user choose a Device and returns it's index
    function SelectDeviceDialog: Integer; virtual;

    //Finds a matching Device index
    function FindDevice(AID: String): Integer; overload;
    function FindDevice(Value: TWIADevice): Integer; overload;
    function FindDevice(AName: String; AManufacturer: String=''): Integer; overload;

    //Returns a Device
    property Devices[Index: Integer]: TWIADevice read GetDevice;

    //Returns the number of Devices
    property DevicesCount: Integer read GetDevicesCount;

    //Selected Device index
    property SelectedDeviceIndex: Integer read GetSelectedDeviceIndex write SetSelectedDeviceIndex;

    //Selected Device in a dialog
    property SelectedDevice: TWIADevice read GetSelectedDevice;

  end;

const
  WIADeviceTypeStr : array  [TWIADeviceType] of String = (
    'Default', 'Scanner', 'Digital Camera', 'Streaming Video'
  );



implementation

uses WIA_SelectForm;

procedure VersionStrToInt(const s: String; var Ver, VerSub: Integer);
var
   pPos, ppPos: Integer;

begin
  Ver:= 0;
  VerSub:= 0;

  try
     pPos:= Pos('.', s);
     if (pPos > 0) then
     begin
       Ver:= StrToInt(Copy(s, 0, pPos-1));
       ppPos:= Pos('.', s, pPos+1);
       if (ppPos > 0)
       then VerSub:= StrToInt(Copy(s, pPos+1, ppPos-1))
       else VerSub:= StrToInt(Copy(s, pPos+1, 255));
     end;
  except

  end;
end;

{ TWIADevice }

function TWIADevice.GetWiaRootItem: IWiaItem2;
//var
//   OleStrID :BSTR;

begin
  Result :=nil;

  if (rOwner <> nil) then
  begin
    if (pWiaRootItem = nil)
    then try
           lres :=rOwner.pWIA_DevMgr.CreateDevice(0, StringToOleStr(Self.rID), pWiaRootItem);
           if (lres = S_OK) then Result :=pWiaRootItem;
         finally
         end
    else Result :=pWiaRootItem;
  end;
end;

function TWIADevice.CreateDestinationStream(bstrItemName: String; var ppDestination: IStream): HRESULT;
begin
  { #todo 10 -oMaxM : Gestione degli Stream in ingresso }

  if (StreamDestination = nil)
  then StreamDestination:= TFileStream.Create('testWia.dat', fmCreate);

  if (StreamAdapter = nil)
  then StreamAdapter:= TStreamAdapter.Create(StreamDestination, soReference);
end;

function TWIADevice.TransferCallback(lFlags: LONG; pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;
begin
  Result:= S_OK;

  if (pWiaTransferParams <> nil) then
  Case pWiaTransferParams^.lMessage of
    WIA_TRANSFER_MSG_STATUS: begin
    end;
    WIA_TRANSFER_MSG_END_OF_STREAM: begin
      if (StreamDestination <> nil) then
      begin
        StreamDestination.Flush;
        //FreeAndNil(StreamDestination);
      end;
    end;
    WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
      DownloadComplete:= True;
      if (StreamDestination <> nil) then
      begin
        //StreamDestination.Flush;
        FreeAndNil(StreamDestination);
      end;
    end;
    WIA_TRANSFER_MSG_DEVICE_STATUS: begin
    end;
    WIA_TRANSFER_MSG_NEW_PAGE: begin
    end
    else begin

    end;
  end;
end;

function TWIADevice.GetNextStream(lFlags: LONG; bstrItemName, bstrFullItemName: BSTR; out ppDestination: IStream): HRESULT; stdcall;
begin
  Result:= S_OK;

  //  Return a new stream for this item's data.
  //
  Result:= CreateDestinationStream(bstrItemName, ppDestination);

end;

constructor TWIADevice.Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
begin
  inherited Create;

  rOwner :=AOwner;
  rIndex :=AIndex;
  rID :=ADeviceID;
  pWiaRootItem :=nil;
  StreamAdapter:= nil;
  StreamDestination:= nil;
  DownloadComplete:= False;
end;

destructor TWIADevice.Destroy;
begin
  if (pWiaRootItem <> nil) then pWiaRootItem:= nil; //Free the Interface
  if (StreamAdapter <> nil) then StreamAdapter.Free;
  if (StreamDestination <> nil) then StreamDestination.Free;

  inherited Destroy;
end;

(*
            if (device.Version == WiaVersion.Wia10)
            {
                // In WIA 1.0, the root device only has a single child, "Scan"
                // https://docs.microsoft.com/en-us/windows-hardware/drivers/image/wia-scanner-tree
                return device.GetSubItems().First();
            }
            else
            {
                // In WIA 2.0, the root device may have multiple children, i.e. "Flatbed" and "Feeder"
                // https://docs.microsoft.com/en-us/windows-hardware/drivers/image/non-duplex-capable-document-feeder
                // The "Feeder" child may also have a pair of children (for front/back sides with duplex)
                // https://docs.microsoft.com/en-us/windows-hardware/drivers/image/simple-duplex-capable-document-feeder
                var items = device.GetSubItems();
                var preferredItemName = _options.PaperSource == PaperSource.Flatbed ? "Flatbed" : "Feeder";
                return items.FirstOrDefault(x => x.Name() == preferredItemName) ?? items.First();
            }

*)

function TWIADevice.Download(const childIndex: Integer): Boolean;
var
   pWiaTransfer: IWiaTransfer;
   myTickStart, curTick:QWord;

begin
  Result:= False;

  if (pWiaRootItem = nil) then GetWiaRootItem;
  if (pWiaRootItem <> nil) then
  begin
    //Vedi Nota Sopra...
    Case rVersion of
      1: begin

      end;
      2: begin
      end;
    end;

    lres:= pWiaRootItem.QueryInterface(IID_IWiaTransfer, pWiaTransfer);
    if (lres = S_OK) and (pWiaTransfer <> nil) then
    begin
      lres:= pWiaTransfer.Download(0, Self); //WIA_TRANSFER_ACQUIRE_CHILDREN);

      DownloadComplete:= False;
      myTickStart:= GetTickCount64; curTick:= myTickStart;
      repeat
        CheckSynchronize;

        curTick:= GetTickCount64;

      until DownloadComplete or ((curTick-myTickStart)>30000);

      // Release the IWiaTransfer
      pWiaTransfer:= nil;
    end;
  end;
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
  pPropSpec: array [0..4] of PROPSPEC;
  pPropVar: array [0..4] of PROPVARIANT;

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

        //
        // Device Wia Version
        //
        pPropSpec[4].ulKind := PRSPEC_PROPID;
        pPropSpec[4].propid := WIA_DIP_WIA_VERSION;
        //pPropIDS[4] :=WIA_DIP_WIA_VERSION;


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
          lres := pWiaPropertyStorage.ReadMultiple(Length(pPropSpec), @pPropSpec, @pPropVar);


         // lres := pWiaPropertyStorage.ReadPropertyNames(Length(pPropIDS), @pPropIDS, @pPropNames);

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

          if (VT_BSTR = pPropVar[4].vt)
          then VersionStrToInt(pPropVar[4].bstrVal, rDeviceList[i].rVersion, rDeviceList[i].rVersionSub)
          else Exception.Create('WiaVersion of Device '+IntToStr(i)+' not String');

          pWiaPropertyStorage:= nil;

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

function TWIAManager.SelectDeviceDialog: Integer;
begin
  Result:= -1;
  try
    Result:= TWIASelectForm.Execute(Self);

  finally
    FreeAndNil(WIASelectForm);
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

procedure TWIAManager.ClearDeviceList;
begin
  EmptyDeviceList(True);
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

