(****************************************************************************
*                FreePascal \ Delphi WIA Implementation
*
*  FILE: WIA.pas
*
*  VERSION:     0.0.1
*
*  DESCRIPTION:
*    WIA Base classes.
*
*****************************************************************************
*
*  (c) 2024 Massimo Magnano
*
*  See changelog.txt for Change Log
*
*****************************************************************************)

unit WIA;

{$H+}

interface

uses
  Windows, Classes, SysUtils, ComObj, ActiveX, WiaDef, WIA_LH;

type
  TWIAManager = class;

  TWIAItem = record
    Name: String;
    ItemType: LONG;
  end;
  PWIAItem = ^TWIAItem;

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
    pRootItem,
    pSelectedItem: IWiaItem2;
    StreamDestination: TFileStream;
    StreamAdapter: TStreamAdapter;
    rSelectedItemIndex: Integer;
    HasEnumerated: Boolean;
    rItemList : array of TWIAItem;

    rDownloaded: Boolean;
    Download_Count: Integer;
    Download_Path,
    Download_BaseFileName: String;

    function GeItem(Index: Integer): PWIAItem;
    function GetItemCount: Integer;
    procedure SetSelectedItemIndex(AValue: Integer);
    function GetRootItem: IWiaItem2;
    function GetSelectedItem: IWiaItem2;

    //Enumerate the avaliable items
    function EnumerateItems: Boolean;

    function CreateDestinationStream(FileName: String; var ppDestination: IStream): HRESULT;

  public
    function TransferCallback(lFlags: LONG;
                              pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;

    function GetNextStream(lFlags: LONG;
                           bstrItemName,
                           bstrFullItemName: BSTR;
                           out ppDestination: IStream): HRESULT; stdcall;

    constructor Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
    destructor Destroy; override;

    function SelectItem(AName: String): Boolean;

    //Download the Selected Item and return the number of files transfered
    function Download(APath, ABaseFileName: String): Integer; virtual;

    property ID: String read rID;
    property Manufacturer: String read rManufacturer;
    property Name: String read rName;
    property Type_: TWIADeviceType read rType;
    property SubType: Word read rSubType;

    property ItemCount: Integer read GetItemCount;

    //Returns a Item
    property Items[Index: Integer]: PWIAItem read GeItem;

    property RootItem: IWiaItem2 read GetRootItem;
    property SelectedItemIndex: Integer read rSelectedItemIndex write SetSelectedItemIndex;
    property SelectedItem: IWiaItem2 read GetSelectedItem;

    property Downloaded: Boolean read rDownloaded;
  end;

  { TWIAManager }

  TOnDeviceTransfer = function (AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                         lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean of object;

  TWIAManager = class(TObject)
  protected
    pWIA_DevMgr: WIA_LH.IWiaDevMgr2;
    lres: HResult;
    HasEnumerated: Boolean;
    rSelectedDeviceIndex: Integer;
    rDeviceList : array of TWIADevice;
    rOnAfterDeviceTransfer,
    rOnBeforeDeviceTransfer: TOnDeviceTransfer;

    function GetSelectedDevice: TWIADevice;
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
    property SelectedDeviceIndex: Integer read rSelectedDeviceIndex write SetSelectedDeviceIndex;

    //Selected Device in a dialog
    property SelectedDevice: TWIADevice read GetSelectedDevice;

    //Events
    property OnBeforeDeviceTransfer:TOnDeviceTransfer read rOnBeforeDeviceTransfer write rOnBeforeDeviceTransfer;
    property OnAfterDeviceTransfer:TOnDeviceTransfer read rOnAfterDeviceTransfer write rOnAfterDeviceTransfer;
  end;

const
  WIADeviceTypeStr : array  [TWIADeviceType] of String = (
    'Default', 'Scanner', 'Digital Camera', 'Streaming Video'
  );



implementation

uses WIA_SelectForm;

{$ifndef fpc}
const
  AllowDirectorySeparators : set of AnsiChar = ['\','/'];
  DirectorySeparator = '\';
{$endif}

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

function TWIADevice.GetSelectedItem: IWiaItem2;
begin
  //Enumerate Items if needed
  if not(HasEnumerated)
  then HasEnumerated:= EnumerateItems;

  if (rSelectedItemIndex >= 0) and (rSelectedItemIndex < GetItemCount)
  then Result:= pSelectedItem
  else Result:= nil;
end;

function TWIADevice.GeItem(Index: Integer): PWIAItem;
begin
  if (Index >= 0) and (Index < Length(rItemList))
  then Result:= @rItemList[Index]
  else Result:= nil;
end;

function TWIADevice.GetItemCount: Integer;
begin
  //Enumerate Items if needed
  if not(HasEnumerated)
  then HasEnumerated:= EnumerateItems;

  Result:= Length(rItemList);
end;

procedure TWIADevice.SetSelectedItemIndex(AValue: Integer);
begin
  if (rSelectedItemIndex <> AValue) and
     (AValue >= 0) and (AValue < GetItemCount)
  then rSelectedItemIndex:= AValue;
end;

function TWIADevice.GetRootItem: IWiaItem2;
//var
//   OleStrID :BSTR;

begin
  Result :=nil;

  if (rOwner <> nil) then
  begin
    if (pRootItem = nil)
    then try
           lres :=rOwner.pWIA_DevMgr.CreateDevice(0, StringToOleStr(Self.rID), pRootItem);
           if (lres = S_OK) then Result :=pRootItem;
         finally
         end
    else Result :=pRootItem;
  end;
end;

function TWIADevice.EnumerateItems: Boolean;
var
   pIEnumItem: IEnumWiaItem2;
   pItem: IWiaItem2;
   iCount,
   itemFetched: ULONG;
   i: Integer;
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   pWiaPropertyStorage: IWiaPropertyStorage;

begin
  Result :=False;
  SetLength(rItemList, 0);

  lres:= GetRootItem.EnumChildItems(nil, pIEnumItem);
  if (lres = S_OK) then
  begin
    lres:= pIEnumItem.GetCount(iCount);
    if (lres = S_OK) then
    begin
      SetLength(rItemList, iCount);

      //Select the First item by default
      if (rSelectedItemIndex < 0) or (rSelectedItemIndex > iCount-1)
      then rSelectedItemIndex:= 0;

      //Free old Selected Item if any
      if (pSelectedItem <> nil)
      then pSelectedItem:= nil;

      Result :=True;

      for i:=0 to iCount-1 do
      begin
        lres:= pIEnumItem.Next(1, pItem, itemFetched);

        Result := (lres = S_OK);
        if Result then
        begin
          rItemList[i].ItemType:= 0;

          lres:= pItem.GetItemType(rItemList[i].ItemType);
          lres:= pItem.QueryInterface(IID_IWiaPropertyStorage, pWiaPropertyStorage);
          if (pWiaPropertyStorage <> nil) then
          begin
            pPropSpec.ulKind := PRSPEC_PROPID;
            pPropSpec.propid := WIA_IPA_ITEM_NAME; //WIA_IPA_FULL_ITEM_NAME

            lres := pWiaPropertyStorage.ReadMultiple(1, @pPropSpec, @pPropVar);

            if (VT_BSTR = pPropVar.vt)
            then rItemList[i].Name :=pPropVar.bstrVal
            else rItemList[i].Name :='?';

            pWiaPropertyStorage:= nil;
          end;

          if (i = rSelectedItemIndex)
          then pSelectedItem:= pItem
          else pItem:= nil;
        end
        else break;
      end;
    end;

    pIEnumItem:= nil;
  end;
end;

function TWIADevice.CreateDestinationStream(FileName: String; var ppDestination: IStream): HRESULT;
begin
  StreamDestination:= TFileStream.Create(FileName, fmCreate);
  ppDestination:= TStreamAdapter.Create(StreamDestination, soOwned);
  Result:= S_OK;
end;

function TWIADevice.TransferCallback(lFlags: LONG; pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;
begin
  if not(Assigned(rOwner.OnBeforeDeviceTransfer))
  then Result:= S_OK
  else if rOwner.OnBeforeDeviceTransfer(rOwner, Self, lFlags, pWiaTransferParams)
       then Result:= S_OK
       else Result:= S_FALSE;   { #todo 2 -oMaxM : Test if this value cancel the Download }

  if (Result = S_OK) and (pWiaTransferParams <> nil) then
  Case pWiaTransferParams^.lMessage of
    WIA_TRANSFER_MSG_STATUS: begin
    end;
    WIA_TRANSFER_MSG_END_OF_STREAM: begin
    end;
    WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
      rDownloaded:= True;
      pWiaTransferParams^.lPercentComplete:=100;
    end;
    WIA_TRANSFER_MSG_DEVICE_STATUS: begin
    end;
    WIA_TRANSFER_MSG_NEW_PAGE: begin
    end
    else begin

    end;
  end;

  if Assigned(rOwner.OnAfterDeviceTransfer)
  then if rOwner.OnAfterDeviceTransfer(rOwner, Self, lFlags, pWiaTransferParams)
       then Result:= S_OK
       else Result:= S_FALSE;   { #todo 2 -oMaxM : Test if this value cancel the Download }
end;

function TWIADevice.GetNextStream(lFlags: LONG; bstrItemName, bstrFullItemName: BSTR; out ppDestination: IStream): HRESULT; stdcall;
var
   FileName: String;
   i: Integer;

begin
  Result:= S_OK;

  //  Return a new stream for this item's data.
  //
  if (Download_Count = 0)
  then Result:= CreateDestinationStream(Download_Path+Download_BaseFileName, ppDestination)
  else begin
         FileName:= Download_BaseFileName;
         i:= LastDelimiter('.', FileName);
         if (i <= 0) then i:=MaxInt;
         Insert('-'+IntToStr(Download_Count), FileName, i);
         Result:= CreateDestinationStream(Download_Path+FileName, ppDestination);
       end;

  Inc(Download_Count);
end;

constructor TWIADevice.Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
begin
  inherited Create;

  rOwner :=AOwner;
  HasEnumerated :=False;
  rSelectedItemIndex :=-1;
  rIndex :=AIndex;
  rID :=ADeviceID;
  pRootItem :=nil;
  pSelectedItem :=nil;
  StreamAdapter:= nil;
  StreamDestination:= nil;
  Download_Path:= '';
  Download_BaseFileName:= '';
  Download_Count:= 0;
  rDownloaded:= False;
end;

destructor TWIADevice.Destroy;
begin
  if (pRootItem <> nil) then pRootItem:= nil; //Free the Interface
  if (pSelectedItem <> nil) then pSelectedItem:= nil;
  if (StreamAdapter <> nil) then StreamAdapter:= nil;
  SetLength(rItemList, 0);

  inherited Destroy;
end;

function TWIADevice.SelectItem(AName: String): Boolean;
var
   pFindItem: IWiaItem2;
   i: Integer;

begin
  Result:= False;
(* { #note -oMaxM : does not work with ITEM_NAME only with FULL_ITEM_NAME  }

  lres:= GetRootItem.FindItemByName(0, StringToOleStr(AName), pFindItem);
  if (lres = S_OK) then
  begin
    pSelectedItem:= nil; //Free Selected Item
    pSelectedItem:= pFindItem;

    Result:= True;
  end;
  *)
  for i:=0 to ItemCount-1 do
  begin
    if (rItemList[i].Name = AName) then
    begin
      Result:= True;
      rSelectedItemIndex:= i;
      break;
    end;
  end;

  //Re Enumerate Items so the correct IWiaItem2 interface is assigned in pSelectedItem
  if Result
  then Result:= EnumerateItems;
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

function TWIADevice.Download(APath, ABaseFileName: String): Integer;
var
   pWiaTransfer: IWiaTransfer;
   myTickStart, curTick: UInt64;
   selItemType: LONG;

begin
  Result:= 0;

  if (pSelectedItem = nil) then GetSelectedItem;
  if (pSelectedItem <> nil) then
  begin
(*    //Vedi Nota Sopra...
    Case rVersion of
      1: begin

      end;
      2: begin
      end;
    end;
*)
    lres:= pSelectedItem.QueryInterface(IID_IWiaTransfer, pWiaTransfer);
    if (lres = S_OK) and (pWiaTransfer <> nil) then
    begin
      selItemType:= rItemList[rSelectedItemIndex].ItemType;

      if (APath = '') or CharInSet(APath[Length(APath)], AllowDirectorySeparators)
      then Download_Path:= APath
      else Download_Path:= APath+DirectorySeparator;

      Download_BaseFileName:= ABaseFileName;
      Download_Count:= 0;
      rDownloaded:= False;

      if (selItemType and WiaItemTypeTransfer = WiaItemTypeTransfer) then
      begin
        if (selItemType and WiaItemTypeFolder = WiaItemTypeFolder)
        then begin
               lres:= pWiaTransfer.Download(WIA_TRANSFER_ACQUIRE_CHILDREN, Self);
             end
        else
        if (selItemType and WiaItemTypeFile = WiaItemTypeFile)
        then begin
               lres:= pWiaTransfer.Download(0, Self);
             end;
      end;

      { #todo 2 -oMaxM : Test if all Scanner is Synch }
      (*

      myTickStart:= GetTickCount64; curTick:= myTickStart;
      repeat
        CheckSynchronize(100);

        curTick:= GetTickCount64;

      until (rDownloaded) or ((curTick-myTickStart) > 27666);
      *)

      // Release the IWiaTransfer
      pWiaTransfer:= nil;

      if rDownloaded
      then Result:= Download_Count
      else Result:= 0;
    end;
  end;
end;

{ TWIAManager }

function TWIAManager.GetSelectedDevice: TWIADevice;
begin
  if (rSelectedDeviceIndex >= 0) and (rSelectedDeviceIndex < Length(rDeviceList))
  then Result:= rDeviceList[rSelectedDeviceIndex]
  else Result:= nil;
end;

procedure TWIAManager.SetSelectedDeviceIndex(AValue: Integer);
begin
  if (AValue <> rSelectedDeviceIndex) and
     (AValue >= 0) and (AValue < Length(rDeviceList)) then
  begin
    if (rDeviceList[AValue] <> nil)
    then rSelectedDeviceIndex:= AValue
    else rSelectedDeviceIndex:= -1;
  end;
end;

function TWIAManager.GetDevice(Index: Integer): TWIADevice;
begin
  if (Index >= 0) and (Index < Length(rDeviceList))
  then Result:= rDeviceList[Index]
  else Result:= nil;
end;

function TWIAManager.GetDevicesCount: Integer;
begin
  if (pWIA_DevMgr = nil)
  then pWIA_DevMgr:= WIA_LH.IWiaDevMgr2(CreateDevManager);

  //Enumerate devices if needed
  if not(HasEnumerated)
  then HasEnumerated:= EnumerateDevices;

  Result:= Length(rDeviceList);
end;

function TWIAManager.CreateDevManager: IUnknown;
begin
  lres:= CoCreateInstance(CLSID_WiaDevMgr2, nil, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, Result);
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

  HasEnumerated:= False;
  pWIA_DevMgr:= nil;
  rSelectedDeviceIndex:= -1;
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

