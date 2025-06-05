(****************************************************************************
*                FreePascal \ Delphi WIA Implementation
*
*  FILE: WIA.pas
*
*  VERSION:     0.0.2
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
{$R-}
{$POINTERMATH ON}

interface

uses
  Windows, DelphiCompatibility, Classes, SysUtils,
  {$ifdef fpc} testutils, {$endif}
  ComObj, ActiveX, WiaDef, WIA_LH, Wia_PaperSizes;

type
  //Dinamic Array types
  TArraySingle = array of Single;
  TArrayDouble = array of Double;
  TArrayInteger = array of Integer;
  TArraySmallint = array of Smallint;
  TArrayByte = array of Byte;
  TArrayWord = array of Word;
  TArrayLongWord = array of LongWord;
  TStringArray = array of String;
  TArrayGUID = array of TGUID;

  TWIAManager = class;

  TWIAItemType = (
    witFree,
    witImage,
    witFile,
    witFolder,
    witRoot,
    witAnalyze,
    witAudio,
    witDevice,
    witDeleted,
    witDisconnected,
    witHPanorama,
    witVPanorama,
    witBurst,
    witStorage,
    witTransfer,
    witGenerated,
    witHasAttachments,
    witVideo,
    witTwainCompatibility,
    witRemoved,
    witDocument,
    witProgrammableDataSource
  );
  TWIAItemTypes = set of TWIAItemType;

  TWIAItemCategory = (
    wicNULL,
    wicFINISHED_FILE,
    wicFLATBED,
    wicFEEDER,
    wicFILM,
    wicROOT,
    wicFOLDER,
    wicFEEDER_FRONT,
    wicFEEDER_BACK,
    wicAUTO,
    wicIMPRINTER,
    wicENDORSER,
    wicBARCODE_READER,
    wicPATCH_CODE_READER,
    wicMICR_READER
 );

  TWIAItem = record
    Name: String;
    ItemType: TWIAItemTypes;
    ItemCategory: TWIAItemCategory;
  end;
  PWIAItem = ^TWIAItem;
  TArrayWIAItem = array of TWIAItem;

  TWIADeviceType = (
    WIADeviceTypeDefault          = StiDeviceTypeDefault,
    WIADeviceTypeScanner          = StiDeviceTypeScanner,
    WIADeviceTypeDigitalCamera    = StiDeviceTypeDigitalCamera,
    WIADeviceTypeStreamingVideo   = StiDeviceTypeStreamingVideo
  );

  TWIAPropertyFlag = (
    WIAProp_READ, WIAProp_WRITE, WIAProp_SYNC_REQUIRED, WIAProp_NONE,
    WIAProp_RANGE, WIAProp_LIST, WIAProp_FLAG, WIAProp_CACHEABLE
  );
  TWIAPropertyFlags = set of TWIAPropertyFlag;

  TWIARotation = (
    wrPortrait = WiaDef.PORTRAIT,
    wrLandscape = WiaDef.LANSCAPE,
    wrRot180 = WiaDef.ROT180,
    wrRot270 = WiaDef.ROT270
  );
  TWIARotationSet = set of TWIARotation;

  TWIADocumentHandling = (
    wdhFeeder,
    wdhFlatbed,
    wdhDuplex,
    wdhFront_First,
    wdhBack_First,
    wdhFront_Only,
    wdhBack_Only,
    wdhNext_Page,
    wdhPreFeed,
    wdhAuto_Advance,
    wdhAdvanced_Duplex
  );
  TWIADocumentHandlingSet = set of TWIADocumentHandling;

  TWIAAlignVertical = (waVTop, waVCenter, waVBottom);
  TWIAAlignHorizontal = (waHLeft, waHCenter, waHRight);

  TWIAImageFormat = (
    wifUNDEFINED,
    wifRAWRGB,
    wifMEMORYBMP,
    wifBMP,
    wifEMF,
    wifWMF,
    wifJPEG,
    wifPNG,
    wifGIF,
    wifTIFF,
    wifEXIF,
    wifPHOTOCD,
    wifFLASHPIX,
    wifICO,
    wifCIFF,
    wifPICT,
    wifJPEG2K,
    wifJPEG2KX,
    wifRAW,
    wifJBIG,
    wifJBIG2
  );
  TWIAImageFormatSet = set of TWIAImageFormat;

  TWIADataType = (
    wdtBN = WIA_DATA_THRESHOLD,
    wdtDITHER = WIA_DATA_DITHER,
    wdtGRAYSCALE = WIA_DATA_GRAYSCALE,
    wdtCOLOR = WIA_DATA_COLOR,
    wdtCOLOR_THRESHOLD = WIA_DATA_COLOR_THRESHOLD,
    wdtCOLOR_DITHER = WIA_DATA_COLOR_DITHER,
    wdtRAW_RGB = WIA_DATA_RAW_RGB,
    wdtRAW_BGR = WIA_DATA_RAW_BGR,
    wdtRAW_YUV = WIA_DATA_RAW_YUV,
    wdtRAW_YUVK = WIA_DATA_RAW_YUVK,
    wdtRAW_CMY = WIA_DATA_RAW_CMY,
    wdtRAW_CMYK = WIA_DATA_RAW_CMYK,
    wdtAUTO = WIA_DATA_AUTO
  );
  TWIADataTypeSet = set of TWIADataType;

  TWIAParams = packed record
    NativeUI: Boolean;
    PaperType: TWIAPaperType;

    //Used only if PaperType=wptCUSTOM
    PaperW,
    PaperH: Integer;

    Rotation: TWIARotation;
    HAlign: TWIAAlignHorizontal;
    VAlign: TWIAAlignVertical;
    Resolution,
    Contrast,
    //BitDepth,
    Brightness: Integer;
    DataType: TWIADataType;
    DocHandling: TWIADocumentHandlingSet; { #todo 5 -oMaxM : Must be tested in a Duplex Scanner first }
  end;
  TArrayWIAParams = array of TWIAParams;

  TWIAParamsCapabilities = packed record
    PaperSizeMaxWidth,
    PaperSizeMaxHeight: Integer;
    PaperTypeSet: TWIAPaperTypeSet;
    PaperTypeCurrent,
    PaperTypeDefault: TWIAPaperType;
    RotationCurrent,
    RotationDefault: TWIARotation;
    RotationSet: TWIARotationSet;
    ResolutionArray: TArrayInteger;
    ResolutionRange: Boolean;
    ResolutionCurrent,
    ResolutionDefault,
    BrightnessCurrent,
    BrightnessDefault,
    BrightnessMin,
    BrightnessMax,
    BrightnessStep,
    ContrastCurrent,
    ContrastDefault,
    ContrastMin,
    ContrastMax,
    ContrastStep (*,
    BitDepthCurrent,
    BitDepthDefault*): Integer;
    //BitDepthArray: TArrayInteger;
    DataTypeCurrent,
    DataTypeDefault: TWIADataType;
    DataTypeSet: TWIADataTypeSet;
    DocHandlingCurrent,
    DocHandlingDefault,
    DocHandlingSet: TWIADocumentHandlingSet; { #todo 5 -oMaxM : Must be tested in a Duplex Scanner }
  end;
  TArrayWIAParamsCapabilities = array of TWIAParamsCapabilities;

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
    pRootProperties,
    pSelectedProperties: IWiaPropertyStorage;

    StreamDestination: TFileStream;
    StreamAdapter: TStreamAdapter;
    rSelectedItemIndex: Integer;
    HasEnumerated: Boolean;
    rItemList : TArrayWIAItem;

    rXRes, rYRes: Integer; //Used with PaperSizes_Calculated, if -1 then i need to Get Values from Device
    rPaperLandscape: Boolean;
    rHAlign: TWIAAlignHorizontal;
    rVAlign: TWIAAlignVertical;

    rDownloaded: Boolean;
    rDownload_Count: Integer;
    rDownload_Path,
    rDownload_Ext,
    rDownload_FileName: String;

    function GetItem(Index: Integer): PWIAItem;
    function GetItemCount: Integer;
    procedure SetSelectedItemIndex(AValue: Integer);
    function GetSelectedItem: PWIAItem;

    function GetRootItemIntf: IWiaItem2;
    function GetSelectedItemIntf: IWiaItem2;
    function GetRootPropertiesIntf: IWiaPropertyStorage;
    function GetSelectedPropertiesIntf: IWiaPropertyStorage;

    //Enumerate the avaliable items
    function EnumerateItems: Boolean;

    function CreateDestinationStream(FileName: String; var ppDestination: IStream): HRESULT;

  public
    //Behaviour Variables
    PaperSizes_Calculated: Boolean;

    //IWiaTransferCallback implementation
    function TransferCallback(lFlags: LONG;
                              pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;
    function GetNextStream(lFlags: LONG;
                           bstrItemName,
                           bstrFullItemName: BSTR;
                           out ppDestination: IStream): HRESULT; stdcall;

    constructor Create(AOwner: TWIAManager; AIndex: Integer; ADeviceID: String);
    destructor Destroy; override;

    function SelectItem(AName: String): Boolean;

    //Download the Selected Item and return the number of files transfered.
    // if multiple pages is downloaded then the file names are
    // APath\AFileName-n.AExt where n is then Index (when 0 n is not present)
    function Download(APath, AFileName, AExt: String): Integer; overload;
    function Download(APath, AFileName, AExt: String; AFormat: TWIAImageFormat): Integer; overload;
    function Download(APath, AFileName, AExt: String; AFormat: TWIAImageFormat;
                      var DownloadedFiles: TStringArray; UseRelativePath: Boolean=False): Integer; overload;
    (*    //  In Wia 2 the user must specify what to download from the feeder using the ADocHandling parameter,
        //  the SettingsForm store DocHandling in Params so the user can use here.
        function Download(APath, AFileName, AExt: String; ADocHandling: TWIADocumentHandlingSet=[]): Integer; overload;
        function Download(APath, AFileName, AExt: String;
                          AFormat: TWIAImageFormat;
                          ADocHandling: TWIADocumentHandlingSet=[]): Integer; overload;
    *)

    //Download using Native UI and return the number of files transfered in DownloadedFiles array.
    //  The system dialog works at Device level, so the selected item is ignored
    function DownloadNativeUI(hwndParent: HWND; useSystemUI: Boolean;
                              APath, AFileName: String;
                              var DownloadedFiles: TStringArray; UseRelativePath: Boolean=False): Integer;

    //Get Current Property Value and it's type given the ID
    function GetProperty(APropId: PROPID; var propType: TVarType;
                         var APropValue; useRoot: Boolean=False): Boolean; overload;

    //Get Current, Default and Possible Values of a Property given the ID,
    //  Depending on the type returned in propType
    //  APropListValues can be a Dynamic Array of Integers, Real, etc... user must free it
    //  if Result contain the Flag WIAProp_RANGE then use WIA_RANGE_XXX Indexes to get MIN/MAX/STEP Values
    function GetProperty(APropId: PROPID; var propType: TVarType;
                         var APropValue, APropDefaultValue;
                         var APropListValues;
                         useRoot: Boolean=False): TWIAPropertyFlags; overload;

    //Get a Range Property with Current, Default, Min, Max, Step Values  { #note 5 -oMaxM : do I keep it? }
    function GetProperty(APropId: PROPID; var propType: TVarType;
                         var APropValue, APropDefault, APropMin, APropMax, APropStep;
                         useRoot: Boolean=False): Boolean; overload;

    //Set the Property Value given the ID, the user must know the correct type to use
    function SetProperty(APropId: PROPID; propType: TVarType; const APropValue; useRoot: Boolean=False): Boolean;

{ #note -oMaxM : Build overloaded functions for Set/Get Property? }
(*
    function SetProperty(APropId: PROPID; APropValue: Smallint; useRoot: Boolean=False): Boolean; overload;  //VT_I2
    function SetProperty(APropId: PROPID; APropValue: Integer; useRoot: Boolean=False): Boolean; overload;   //VT_I4, VT_INT
    function SetProperty(APropId: PROPID; APropValue: Single; useRoot: Boolean=False): Boolean; overload;    //VT_R4
    function SetProperty(APropId: PROPID; APropValue: Double; useRoot: Boolean=False): Boolean; overload;    //VT_R8
    function SetProperty(APropId: PROPID; APropValue: Currency; useRoot: Boolean=False): Boolean; overload;  //VT_R8
    function SetProperty(APropId: PROPID; APropValue: TDateTime; useRoot: Boolean=False): Boolean; overload; //VT_DATE
    function SetProperty(APropId: PROPID; APropValue: BSTR; useRoot: Boolean=False): Boolean; overload;      //VT_BSTR
    function SetProperty(APropId: PROPID; APropValue: Boolean; useRoot: Boolean=False): Boolean; overload;   //VT_BOOL
    function SetProperty(APropId: PROPID; APropValue: Word; useRoot: Boolean=False): Boolean; overload;      //VT_UI2
    function SetProperty(APropId: PROPID; APropValue: DWord; useRoot: Boolean=False): Boolean; overload;     //VT_UI4, VT_UINT
    function SetProperty(APropId: PROPID; APropValue: Int64; useRoot: Boolean=False): Boolean; overload;     //VT_I8
    function SetProperty(APropId: PROPID; APropValue: UInt64; useRoot: Boolean=False): Boolean; overload;    //VT_UI8
    function SetProperty(APropId: PROPID; APropValue: LPSTR; useRoot: Boolean=False): Boolean; overload;     //VT_LPSTR
//    procedure SetProperty(APropId: PROPID; APropValue: LPWSTR; useRoot: Boolean=False): Boolean; overload;  //VT_LPWSTR
*)

    //Get Available Values for XResolution,
    //  if Result contain the Flag WIAProp_RANGE then use WIA_RANGE_XXX Indexes to get MIN/MAX/STEP Values
    function GetResolutionsX(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean=False): TWIAPropertyFlags;
    //Get Available Values for YResolutions
    function GetResolutionsY(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean=False): TWIAPropertyFlags;

    //Get the Minimun and Maximum Resolutions Values
    function GetResolutionsLimit(var AMin, AMax: Integer; useRoot: Boolean=False): Boolean; overload;
    function GetResolutionsLimit(var AMinX, AMaxX, AMinY, AMaxY: Integer; useRoot: Boolean=False): Boolean; overload;

    //Get Current Resolutions
    function GetResolution(var AXRes, AYRes: Integer; useRoot: Boolean=False): Boolean;

    //Set Current Resolutions, The user is responsible for checking the validity of the values
    function SetResolution(const AXRes, AYRes: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Paper Align
    function GetPaperAlign(var ALandscape:Boolean; var HAlign: TWIAAlignHorizontal; var VAlign: TWIAAlignVertical;
                           useRoot: Boolean=False): Boolean;
    //Set Current Paper Align
    function SetPaperAlign(const ALandscape:Boolean; const HAlign: TWIAAlignHorizontal; const VAlign: TWIAAlignVertical;
                           useRoot: Boolean=False): Boolean;

    //Get Max Paper Width, Height
    function GetPaperSizeMax(var AMaxWidth, AMaxHeight: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Paper Type
    //  if PaperSizes_Calculated then the PaperSize is calculated using real Paper Size
    //  else this functions uses WIA_IPS_PAGE_SIZE
    function GetPaperType(var Current: TWIAPaperType; useRoot: Boolean=False): Boolean; overload;
    //Get Available Paper Sizes
    function GetPaperType(var Current, Default: TWIAPaperType; var Values: TWIAPaperTypeSet; useRoot: Boolean=False): Boolean; overload;

    //Set Current Paper Type,
    //  if PaperSizes_Calculated then the Area is calculated using real Paper Size,Orientation and Align Values
    //  else this function use WIA_IPS_PAGE_SIZE and the user is responsible for checking the validity of the value
    function SetPaperType(const Value: TWIAPaperType; useRoot: Boolean=False): Boolean;

    //Set Current Paper Size, the Area is calculated using Width, Height (in TH Inchs),Orientation and Align Values
    function SetPaperSize(Width, Height: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Paper Landscape,
    //  if PaperSizes_Calculated=False then this function use WIA_IPS_ROTATION else return internal value
    function GetPaperLandscape(var Value: Boolean; useRoot: Boolean=False): Boolean;

    //Set Current Paper Landscape,
    //  if PaperSizes_Calculated=False then this function use WIA_IPS_ROTATION
    //  else set internal value and rotation is done when papersize is setted
    function SetPaperLandscape(const Value: Boolean; useRoot: Boolean=False): Boolean;

    //Get Current Rotation,
    //  Not to be confused with PaperLandscape, this function only uses WIA_IPS_ROTATION
    function GetRotation(var Value: TWIARotation; useRoot: Boolean=False): Boolean; overload;
    //Get Available Rotations
    function GetRotation(var Current, Default: TWIARotation; var Values: TWIARotationSet; useRoot: Boolean=False): Boolean; overload;

    //Set Current Rotation,
    //  Not to be confused with PaperLandscape, this function only uses WIA_IPS_ROTATION
    //  which rotates the image after capturing it
    function SetRotation(const Value: TWIARotation; useRoot: Boolean=False): Boolean;

    //Get Current DocumentHandling,
    function GetDocumentHandling(var Value: TWIADocumentHandlingSet; useRoot: Boolean=False): Boolean; overload;
    //Get Available DocumentHandling
    function GetDocumentHandling(var Current, Default, Values: TWIADocumentHandlingSet; useRoot: Boolean=False): Boolean; overload;

    //Set Current Rotation,
    //  Not to be confused with PaperLandscape, this function only uses WIA_IPS_ROTATION
    //  which rotates the image after capturing it
    function SetDocumentHandling(const Value: TWIADocumentHandlingSet; useRoot: Boolean=False): Boolean;

    //Get Current Pages (0 = All)
    function GetPages(var Current: Integer; useRoot: Boolean=False): Boolean; overload;
    //Get Current, Default and Range Values for Pages
    function GetPages(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean=False): Boolean; overload;

    //Set Current Pages (0 = All)
    //  If a Feeder Scanner is unable to scan only one side of a page while in Duplex you must use an even value
    function SetPages(const Value: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Brightness
    function GetBrightness(var Current: Integer; useRoot: Boolean=False): Boolean; overload;
    //Get Current, Default and Range Values for Brightness
    function GetBrightness(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean=False): Boolean; overload;

    //Set Current Brightness, The user is responsible for checking the validity of the value
    function SetBrightness(const Value: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Contrast
    function GetContrast(var Current: Integer; useRoot: Boolean=False): Boolean; overload;
    //Get Current, Default and Range Values for Contrast
    function GetContrast(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean=False): Boolean; overload;

    //Set Current Contrast, The user is responsible for checking the validity of the value
    function SetContrast(const Value: Integer; useRoot: Boolean=False): Boolean;

    //Get Current Image Format
    function GetImageFormat(var Current: TWIAImageFormat; useRoot: Boolean=False): Boolean; overload;
    //Get Available Image Formats
    function GetImageFormat(var Current, Default: TWIAImageFormat; var Values: TWIAImageFormatSet; useRoot: Boolean=False): Boolean; overload;

    //Set Current Image Format
    function SetImageFormat(const Value: TWIAImageFormat; useRoot: Boolean=False): Boolean;

     //Get Current Image DataType
    function GetDataType(var Current: TWIADataType; useRoot: Boolean=False): Boolean; overload;
    //Get Available Image DataTypes
    function GetDataType(var Current, Default: TWIADataType; var Values: TWIADataTypeSet; useRoot: Boolean=False): Boolean; overload;

    //Set Current Image DataType
    function SetDataType(const Value: TWIADataType; useRoot: Boolean=False): Boolean;

    //Get Available Values for BitDepth
    function GetBitDepth(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean=False): TWIAPropertyFlags; overload;
    //Get Current BitDepth
    function GetBitDepth(var Current: Integer; useRoot: Boolean=False): Boolean; overload;

    //Set Current BitDepth, The user is responsible for checking the validity of the value
    function SetBitDepth(const Value: Integer; useRoot: Boolean=False): Boolean;

    //Get Capabilities for Current Selected Item
    function GetParamsCapabilities(var Value: TWIAParamsCapabilities): Boolean;

    //Set Params to Current Selected Item
    function SetParams(const AParams: TWIAParams): Boolean;

    //Get Sub Items of Selected Item if any
    function GetSelectedItemSubItems(var AItemArray: TArrayWIAItem): Boolean;

    property ID: String read rID;
    property Manufacturer: String read rManufacturer;
    property Name: String read rName;
    property Type_: TWIADeviceType read rType;
    property SubType: Word read rSubType;

    property ItemCount: Integer read GetItemCount;

    //Returns an Item by it's Index
    // Note:
    //   In WIA 1.0, the root device only has a single child, "Scan"
    //     https://docs.microsoft.com/en-us/windows-hardware/drivers/image/wia-scanner-tree
    //   In WIA 2.0, the root device may have multiple children, i.e. "Flatbed" and "Feeder"
    //     https://docs.microsoft.com/en-us/windows-hardware/drivers/image/non-duplex-capable-document-feeder
    //     The "Feeder" child may also have a pair of children (for front/back sides with duplex)
    //     https://docs.microsoft.com/en-us/windows-hardware/drivers/image/simple-duplex-capable-document-feeder

    { #todo 10 -oMaxM : It would be better to keep an array of Item Classes where you can move the various Get/Set Properties methods}
    property Items[Index: Integer]: PWIAItem read GetItem;

    property RootItemIntf: IWiaItem2 read GetRootItemIntf;
    property RootPropertiesIntf: IWiaPropertyStorage read GetRootPropertiesIntf;

    property SelectedItemIntf: IWiaItem2 read GetSelectedItemIntf;
    property SelectedPropertiesIntf: IWiaPropertyStorage read GetSelectedPropertiesIntf;

    property SelectedItem: PWIAItem read GetSelectedItem;
    property SelectedItemIndex: Integer read rSelectedItemIndex write SetSelectedItemIndex;

    property Downloaded: Boolean read rDownloaded;
    property Download_Count: Integer read rDownload_Count;
    property Download_Path: String read rDownload_Path;
    property Download_Ext: String read rDownload_Ext;
    property Download_FileName: String read rDownload_FileName;
  end;

  { TWIAManager }

  TOnDeviceTransfer = function (AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                         lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean of object;

  TWIAManager = class(TObject)
  protected
    rEnumAll: Boolean;
    pDevMgr: WIA_LH.IWiaDevMgr2;
    lres: HResult;
    HasEnumerated: Boolean;
    rSelectedDeviceIndex: Integer;
    rDeviceList : array of TWIADevice;
    rOnAfterDeviceTransfer,
    rOnBeforeDeviceTransfer: TOnDeviceTransfer;

    function GetDevMgrIntf: WIA_LH.IWiaDevMgr2;
    procedure SetEnumAll(AValue: Boolean);
    function GetSelectedDevice: TWIADevice;
    procedure SetSelectedDeviceIndex(AValue: Integer);
    function GetDevice(Index: Integer): TWIADevice;
    function GetDevicesCount: Integer;

    function CreateDevManager: IUnknown; virtual;

    procedure EmptyDeviceList(setZeroLength:Boolean);

    //Enumerate the avaliable devices
    function EnumerateDevices: Boolean;

  public
    constructor Create(AEnumAll: Boolean = True);
    destructor Destroy; override;

    //Clears the list of sources
    procedure ClearDeviceList;

    //Refresh the list of sources
    procedure RefreshDeviceList;

    //Display a dialog to let the user choose a Device and returns it's index
    function SelectDeviceDialog: Integer; virtual;

    //Finds a matching Device index
    function FindDevice(AID: String): Integer; overload;
    function FindDevice(Value: TWIADevice): Integer; overload;
    function FindDevice(AName: String; AManufacturer: String=''): Integer; overload;

    //Find a Device by it's ID and Select the given Item
    procedure SelectDeviceItem(ADeviceID, ADeviceItem: String; var ADevice: TWIADevice; var AIndex: Integer);

    property DevMgrIntf: WIA_LH.IWiaDevMgr2 read GetDevMgrIntf;

    //Kind of Enum, if True Enum even disconnected Devices
    property EnumAll: Boolean read rEnumAll write SetEnumAll;

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
  WIADeviceTypeDescr : array  [TWIADeviceType] of String = (
    'Default', 'Scanner', 'Digital Camera', 'Streaming Video'
  );

  WiaImageFormatGUID : array [TWIAImageFormat] of TGUID = (
    '{b96b3ca9-0728-11d3-9d7b-0000f81ef32e}',
    '{bca48b55-f272-4371-b0f1-4a150d057bb4}',
    '{b96b3caa-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cab-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cac-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cad-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cae-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3caf-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb0-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb1-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb2-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb3-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb4-0728-11d3-9d7b-0000f81ef32e}',
    '{b96b3cb5-0728-11d3-9d7b-0000f81ef32e}',
    '{9821a8ab-3a7e-4215-94e0-d27a460c03b2}',
    '{a6bc85d8-6b3e-40ee-a95c-25d482e41adc}',
    '{344ee2b2-39db-4dde-8173-c4b75f8f1e49}',
    '{43e14614-c80a-4850-baf3-4b152dc8da27}',
    '{6f120719-f1a8-4e07-9ade-9b64c63a3dcc}',
    '{41e8dd92-2f0a-43d4-8636-f1614ba11e46}',
    '{bb8e7e67-283c-4235-9e59-0b9bf94ca687}'
  );
  WiaImageFormatDescr : array [TWIAImageFormat] of String = (
    'Undefined',
    'Raw RGB format',
    'Windows bitmap without a header',
    'Windows Device Independent Bitmap',
    'Extended Windows metafile',
    'Windows metafile',
    'JPEG compressed format',
    'W3C PNG format',
    'GIF image format',
    'Tagged Image File format',
    'Exchangeable File Format',
    'Eastman Kodak file format',
    'FlashPix format',
    'Windows icon file format',
    'Camera Image File format',
    'Apple file format',
    'JPEG 2000 compressed format',
    'JPEG 2000X compressed format',
    'WIA Raw image file format',
    'Joint Bi-level Image experts Group format',
    'Joint Bi-level Image experts Group format (ver 2)'
  );

  WiaItemCategoryGUID : array[TWIAItemCategory] of TGUID = (
    '',
    '{ff2b77ca-cf84-432b-a735-3a130dde2a88}',
    '{fb607b1f-43f3-488b-855b-fb703ec342a6}',
    '{fe131934-f84c-42ad-8da4-6129cddd7288}',
    '{fcf65be7-3ce3-4473-af85-f5d37d21b68a}',
    '{f193526f-59b8-4a26-9888-e16e4f97ce10}',
    '{c692a446-6f5a-481d-85bb-92e2e86fd30a}',
    '{4823175c-3b28-487b-a7e6-eebc17614fd1}',
    '{61ca74d4-39db-42aa-89b1-8c19c9cd4c23}',
    '{defe5fd8-6c97-4dde-b11e-cb509b270e11}',
    '{fc65016d-9202-43dd-91a7-64c2954cfb8b}',
    '{47102cc3-127f-4771-adfc-991ab8ee1e97}',
    '{36e178a0-473f-494b-af8f-6c3f6d7486fc}',
    '{8faa1a6d-9c8a-42cd-98b3-ee9700cbc74f}',
    '{3b86c1ec-71bc-4645-b4d5-1b19da2be978}'
  );

  WIADataTypeDescr: array [wdtBN..wdtRAW_CMYK] of String = (
    'Black & White',
    'Gray scale (Dither)',
    'Gray scale',
    'Color (RGB)',
    'Color (BN)',
    'Color (Dither)',
    'Color (RAW RGB)',
    'Color (RAW BGR)',
    'Color (RAW YUV)',
    'Color (RAW YUVK)',
    'Color (RAW CMY)',
    'Color (RAW CMYK)'
  );


function WIAItemTypes(pItemType: LONG): TWIAItemTypes;
function WIAItemCategory(const AGUID: TGUID): TWIAItemCategory;

function WIAPropertyFlags(pFlags: ULONG): TWIAPropertyFlags;

function WIACopyCurrentValues(const WIACap: TWIAParamsCapabilities;
                              aHAlign: TWIAAlignHorizontal=waHLeft;
                              aVAlign: TWIAAlignVertical=waVTop): TWIAParams;
function WIACopyDefaultValues(const WIACap: TWIAParamsCapabilities;
                              aHAlign: TWIAAlignHorizontal=waHLeft;
                              aVAlign: TWIAAlignVertical=waVTop): TWIAParams;

function WIAImageFormat(const AGUID: TGUID; var Value: TWIAImageFormat): Boolean;

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

function FullPathToRelativePath(const ABasePath: String; var APath: String): Boolean;
begin
  Result:= (Pos(ABasePath, APath) = 1);
  if Result
  then APath:= '.'+DirectorySeparator+Copy(APath, Length(ABasePath)+1, MaxInt);
end;

function WIAItemTypes(pItemType: LONG): TWIAItemTypes;
begin
  Result :=[];

  if (pItemType and WiaItemTypeFree <> 0) then Result:= Result+[witFree];
  if (pItemType and WiaItemTypeImage <> 0) then Result:= Result+[witImage];
  if (pItemType and WiaItemTypeFile <> 0) then Result:= Result+[witFile];
  if (pItemType and WiaItemTypeFolder <> 0) then Result:= Result+[witFolder];
  if (pItemType and WiaItemTypeRoot <> 0) then Result:= Result+[witRoot];
  if (pItemType and WiaItemTypeAnalyze <> 0) then Result:= Result+[witAnalyze];
  if (pItemType and WiaItemTypeAudio <> 0) then Result:= Result+[witAudio];
  if (pItemType and WiaItemTypeDevice <> 0) then Result:= Result+[witDevice];
  if (pItemType and WiaItemTypeDeleted <> 0) then Result:= Result+[witDeleted];
  if (pItemType and WiaItemTypeDisconnected <> 0) then Result:= Result+[witDisconnected];
  if (pItemType and WiaItemTypeHPanorama <> 0) then Result:= Result+[witHPanorama];
  if (pItemType and WiaItemTypeVPanorama <> 0) then Result:= Result+[witVPanorama];
  if (pItemType and WiaItemTypeBurst <> 0) then Result:= Result+[witBurst];
  if (pItemType and WiaItemTypeStorage <> 0) then Result:= Result+[witStorage];
  if (pItemType and WiaItemTypeTransfer	<> 0) then Result:= Result+[witTransfer];
  if (pItemType and WiaItemTypeGenerated <> 0) then Result:= Result+[witGenerated];
  if (pItemType and WiaItemTypeHasAttachments <> 0) then Result:= Result+[witHasAttachments];
  if (pItemType and WiaItemTypeVideo <> 0) then Result:= Result+[witVideo];
  if (pItemType and WiaItemTypeTwainCompatibility <> 0) then Result:= Result+[witTwainCompatibility];
  if (pItemType and WiaItemTypeRemoved <> 0) then Result:= Result+[witRemoved];
  if (pItemType and WiaItemTypeDocument	<> 0) then Result:= Result+[witDocument];
  if (pItemType and WiaItemTypeProgrammableDataSource <> 0) then Result:= Result+[witProgrammableDataSource];
end;

function WIAItemCategory(const AGUID: TGUID): TWIAItemCategory;
var
   i: TWIAItemCategory;

begin
  Result:= wicNULL;

  for i:=Low(TWIAItemCategory) to High(TWIAItemCategory) do
    if IsEqualGUID(WiaItemCategoryGUID[i], AGUID) then
    begin
      Result:= i;
      break;
    end;
end;

function WIAPropertyFlags(pFlags: ULONG): TWIAPropertyFlags;
begin
  Result :=[];

  if (pFlags and WIA_PROP_READ <> 0) then Result:= Result+[WIAProp_READ];
  if (pFlags and WIA_PROP_WRITE <> 0) then Result:= Result+[WIAProp_WRITE];
  if (pFlags and WIA_PROP_SYNC_REQUIRED <> 0) then Result:= Result+[WIAProp_SYNC_REQUIRED];
  if (pFlags and WIA_PROP_NONE <> 0) then Result:= Result+[WIAProp_NONE];
  if (pFlags and WIA_PROP_RANGE <> 0) then Result:= Result+[WIAProp_RANGE];
  if (pFlags and WIA_PROP_LIST <> 0) then Result:= Result+[WIAProp_LIST];
  if (pFlags and WIA_PROP_FLAG <> 0) then Result:= Result+[WIAProp_FLAG];
  if (pFlags and WIA_PROP_CACHEABLE <> 0) then Result:= Result+[WIAProp_CACHEABLE];
end;

function WIACopyCurrentValues(const WIACap: TWIAParamsCapabilities;
                              aHAlign: TWIAAlignHorizontal=waHLeft;
                              aVAlign: TWIAAlignVertical=waVTop): TWIAParams;
begin
  with Result do
  begin
    PaperType:= WIACap.PaperTypeCurrent;
    Resolution:= WIACap.ResolutionCurrent;
    Contrast:= WiaCap.ContrastCurrent;
    Brightness:= WIACap.BrightnessCurrent;
    DocHandling:= WIACap.DocHandlingCurrent;
    //BitDepth:= WIACap.BitDepthCurrent;
    DataType:= WIACap.DataTypeCurrent;
    HAlign:= aHAlign;
    VAlign:= aVAlign;
  end;
end;

function WIACopyDefaultValues(const WIACap: TWIAParamsCapabilities;
                              aHAlign: TWIAAlignHorizontal=waHLeft;
                              aVAlign: TWIAAlignVertical=waVTop): TWIAParams;
begin
  FillChar(Result, Sizeof(Result), 0);
  with Result do
  begin
    PaperType:= WIACap.PaperTypeDefault;
    Resolution:= WIACap.ResolutionDefault;
    Contrast:= WiaCap.ContrastDefault;
    Brightness:= WIACap.BrightnessDefault;
    DocHandling:= WIACap.DocHandlingDefault;
    //BitDepth:= WIACap.BitDepthDefault;
    DataType:= WIACap.DataTypeDefault;
    HAlign:= aHAlign;
    VAlign:= aVAlign;
  end;
end;

function WIAImageFormat(const AGUID: TGUID; var Value: TWIAImageFormat): Boolean;
var
   i: TWIAImageFormat;

begin
  Result:= False;

  for i:=Low(TWIAImageFormat) to High(TWIAImageFormat) do
    if IsEqualGUID(WiaImageFormatGUID[i], AGUID) then
    begin
      Value:= i;
      Result:= True;
      break;
    end;
end;

function WIADocumentHandling(iDocumentHandling: Integer): TWIADocumentHandlingSet;
begin
  Result :=[];

  if (iDocumentHandling and FEEDER <> 0) then Result:= Result+[wdhFeeder];
  if (iDocumentHandling and FLATBED <> 0) then Result:= Result+[wdhFlatbed];
  if (iDocumentHandling and DUPLEX <> 0) then Result:= Result+[wdhDuplex];
  if (iDocumentHandling and FRONT_FIRST <> 0) then Result:= Result+[wdhFront_First];
  if (iDocumentHandling and BACK_FIRST <> 0) then Result:= Result+[wdhBack_First];
  if (iDocumentHandling and FRONT_ONLY <> 0) then Result:= Result+[wdhFront_Only];
  if (iDocumentHandling and BACK_ONLY <> 0) then Result:= Result+[wdhBack_Only];
  if (iDocumentHandling and NEXT_PAGE <> 0) then Result:= Result+[wdhNext_Page];
  if (iDocumentHandling and PREFEED <> 0) then Result:= Result+[wdhPreFeed];
  if (iDocumentHandling and AUTO_ADVANCE <> 0) then Result:= Result+[wdhAuto_Advance];
  if (iDocumentHandling and ADVANCED_DUPLEX <> 0) then Result:= Result+[wdhAdvanced_Duplex];
end;

function WIADocumentHandlingInt(sDocumentHandling: TWIADocumentHandlingSet): Integer;
begin
  Result :=0;

  if (wdhFeeder in sDocumentHandling) then Result:= Result or FEEDER;
  if (wdhFlatbed in sDocumentHandling) then Result:= Result or FLATBED;
  if (wdhDuplex in sDocumentHandling) then Result:= Result or DUPLEX;
  if (wdhFront_First in sDocumentHandling) then Result:= Result or FRONT_FIRST;
  if (wdhBack_First in sDocumentHandling) then Result:= Result or BACK_FIRST;
  if (wdhFront_Only in sDocumentHandling) then Result:= Result or FRONT_ONLY;
  if (wdhBack_Only in sDocumentHandling) then Result:= Result or BACK_ONLY;
  if (wdhNext_Page in sDocumentHandling) then Result:= Result or NEXT_PAGE;
  if (wdhPreFeed in sDocumentHandling) then Result:= Result or PREFEED;
  if (wdhAuto_Advance in sDocumentHandling) then Result:= Result or AUTO_ADVANCE;
  if (wdhAdvanced_Duplex in sDocumentHandling) then Result:= Result or ADVANCED_DUPLEX;
end;

{ TWIADevice }

function TWIADevice.GetRootItemIntf: IWiaItem2;
begin
  Result :=nil;

  if (rOwner <> nil) then
  begin
    if (pRootItem = nil)
    then try
           lres :=rOwner.pDevMgr.CreateDevice(0, StringToOleStr(Self.rID), pRootItem);
           if (lres = S_OK) then Result :=pRootItem;
         finally
         end
    else Result :=pRootItem;
  end;
end;

function TWIADevice.GetSelectedItemIntf: IWiaItem2;
begin
  //Enumerate Items if needed
  if not(HasEnumerated)
  then HasEnumerated:= EnumerateItems;

  if HasEnumerated and
     (rSelectedItemIndex >= 0) and (rSelectedItemIndex < GetItemCount)
  then Result:= pSelectedItem
  else Result:= nil;
end;

function TWIADevice.GetRootPropertiesIntf: IWiaPropertyStorage;
begin
  Result:= nil;

  if (pRootItem = nil) then GetRootItemIntf;
  if (pRootItem <> nil) then
  begin
    if (pRootProperties = nil)
    then lres:= pRootItem.QueryInterface(IID_IWiaPropertyStorage, pRootProperties);

    Result:= pRootProperties;
  end;
end;

function TWIADevice.GetSelectedPropertiesIntf: IWiaPropertyStorage;
begin
  Result:= nil;

  if (pSelectedItem = nil) then GetSelectedItemIntf;
  if (pSelectedItem <> nil)
  then Result:= pSelectedProperties;
end;

function TWIADevice.GetItem(Index: Integer): PWIAItem;
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

  if HasEnumerated
  then Result:= Length(rItemList)
  else Result:= 0;
end;

procedure TWIADevice.SetSelectedItemIndex(AValue: Integer);
begin
  if (rSelectedItemIndex <> AValue) and
     (AValue >= 0) and (AValue < GetItemCount)
  then begin
         rSelectedItemIndex:= AValue;

         //Re Enumerate Items so the correct IWiaItem2 interface is assigned in pSelectedItem
         EnumerateItems;
       end;
end;

function TWIADevice.GetSelectedItem: PWIAItem;
begin
  Result:= GetItem(rSelectedItemIndex);
end;

function TWIADevice.EnumerateItems: Boolean;
var
   pIEnumItem: IEnumWiaItem2;
   pItem: IWiaItem2;
   iCount,
   itemFetched: ULONG;
   itemType: LONG;
   itemCategory: TGUID;
   i: Integer;
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   pWiaPropertyStorage: IWiaPropertyStorage;

begin
  Result :=False;
  try
  rItemList:= nil;

  if (GetRootItemIntf = nil) then exit;

  lres:= pRootItem.EnumChildItems(nil, pIEnumItem);
  if (lres = S_OK) then
  begin
    lres:= pIEnumItem.GetCount(iCount);
    if (lres = S_OK) then
    begin
      SetLength(rItemList, iCount);

      //Select the First item by default
      if (rSelectedItemIndex < 0) or (rSelectedItemIndex > iCount-1)
      then rSelectedItemIndex:= 0;

      //If there is an Item Selected free Interfaces pointers
      if (pSelectedItem <> nil) then
      begin
        pSelectedItem:= nil;
        pSelectedProperties:= nil;
      end;

      for i:=0 to iCount-1 do
      begin
        lres:= pIEnumItem.Next(1, pItem, itemFetched);

        Result := (lres = S_OK);
        if Result then
        begin
          lres:= pItem.GetItemType(itemType);
          if (lres = S_OK)
          then rItemList[i].ItemType :=WIAItemTypes(itemType)
          else rItemList[i].ItemType :=[];

          lres:= pItem.GetItemCategory(itemCategory);
          if (lres = S_OK)
          then rItemList[i].ItemCategory :=WIAItemCategory(itemCategory)
          else rItemList[i].ItemCategory :=wicNULL;

          lres:= pItem.QueryInterface(IID_IWiaPropertyStorage, pWiaPropertyStorage);
          if (lres = S_OK) and (pWiaPropertyStorage <> nil) then
          begin
            pPropSpec.ulKind := PRSPEC_PROPID;
            pPropSpec.propid := WIA_IPA_ITEM_NAME; //WIA_IPA_FULL_ITEM_NAME

            lres := pWiaPropertyStorage.ReadMultiple(1, @pPropSpec, @pPropVar);

            if (VT_BSTR = pPropVar.vt)
            then rItemList[i].Name :=pPropVar.bstrVal
            else rItemList[i].Name :='?';
          end;

          if (i = rSelectedItemIndex)
          then begin
                 //if it is the Selected Item keep the Interfaces pointers
                 pSelectedItem:= pItem;
                 pSelectedProperties:= pWiaPropertyStorage;
               end
          else begin
                 //else Release it
                 pItem:= nil;
                 pWiaPropertyStorage:= nil;
               end;
        end
        else break;
      end;

      Result :=True;
    end;

    pIEnumItem:= nil;
  end;

  except
    rItemList:= nil;
    Result:= False;
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
      if (pWiaTransferParams^.ulTransferredBytes > 0)
      then Inc(rDownload_Count);

      rDownloaded:= (rDownload_Count > 0);
    end;
    WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
      //rDownloaded:= True;
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
begin
  Result:= S_OK;

  //  Return a new stream for this item's data.
  if (rDownload_Count = 0)
  then Result:= CreateDestinationStream(rDownload_Path+rDownload_FileName+rDownload_Ext, ppDestination)
  else Result:= CreateDestinationStream(rDownload_Path+rDownload_FileName+
                                        '-'+IntToStr(rDownload_Count)+rDownload_Ext, ppDestination);

  //Inc(rDownload_Count);
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
  pSelectedProperties:= nil;
  StreamAdapter:= nil;
  StreamDestination:= nil;
  rDownload_Path:= '';
  rDownload_Ext:= '';
  rDownload_FileName:= '';
  rDownload_Count:= 0;
  rDownloaded:= False;

  //By Default is True because Microsoft Documentation says:
  //"Note   Flatbed and Film child items must support only the WIA_IPS_X* Properties"
  //So the WIA_IPS_PAGE_SIZE it's useless
  PaperSizes_Calculated:= True;
  rXRes:= -1; rYRes:= -1;
  rPaperLandscape:= False;
  rVAlign:= waVTop;
  rHAlign:= waHLeft;
end;

destructor TWIADevice.Destroy;
begin
  //Free the Interfaces
  if (pRootItem <> nil) then pRootItem:= nil;
  if (pSelectedItem <> nil) then pSelectedItem:= nil;
  if (pSelectedProperties <> nil) then pSelectedProperties:= nil;
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

function TWIADevice.Download(APath, AFileName, AExt: String): Integer;
var
   pWiaTransfer: IWiaTransfer;
   myTickStart, curTick: UInt64;
   selItem: TWIAItem;

   procedure DownloadSingleItem;
   begin
     lres:= pSelectedItem.QueryInterface(IID_IWiaTransfer, pWiaTransfer);
     if (lres = S_OK) and (pWiaTransfer <> nil) then
     try
       { #todo 10 -oMaxM : Check this in Various Scanner / Camera }
       // in My Samsung 00082007 =  [witImage,witFile,witFolder,witProgrammableDataSource] WHY witFolder?
       // in Kyocera via LAN = [witFolder, witStorage]
       (*
       if (witTransfer in selItem.ItemType) then
       begin
         if (witProgrammableDataSource in selItem.ItemType)
         then lres:= pWiaTransfer.Download(0, Self)
         else
         if (witDocument in selItem.ItemType) then
         begin
           if (witFolder in selItem.ItemType)
           then begin
                  lres:= pWiaTransfer.Download(WIA_TRANSFER_ACQUIRE_CHILDREN, Self);
                end
           else
           if (witFile in selItem.ItemType)
           then begin
                  lres:= pWiaTransfer.Download(0, Self);
                end;
         end;
       end;
       *)

       lres:= pWiaTransfer.Download(0, Self);

     finally
       // Release the IWiaTransfer
       pWiaTransfer:= nil;
     end;
   end;

   procedure DownloadAdvDuplex;
   var
      AItemArray: TArrayWIAItem;
      curSubItem: IWiaItem2;

   begin
     try
        if GetSelectedItemSubItems(AItemArray) then
        begin
          { #todo 10 -oMaxM : Implement me (If i find a Duplex Scanner for Free) }
        end;

     finally
       AItemArray:= nil;
     end;
   end;

begin
  Result:= 0;

  if (pSelectedItem = nil) then GetSelectedItemIntf;
  if (pSelectedItem <> nil) then
  begin
    selItem:= rItemList[rSelectedItemIndex];

    if (APath = '') or CharInSet(APath[Length(APath)], AllowDirectorySeparators)
    then rDownload_Path:= APath
    else rDownload_Path:= APath+DirectorySeparator;

    if not(ForceDirectories(rDownload_Path)) then exit;

    rDownload_FileName:= AFileName;
    rDownload_Ext:= AExt;
    rDownload_Count:= 0;
    rDownloaded:= False;

    (*if (selItem.ItemCategory = wicFEEDER)
    then begin
           if (rVersion = 2) and (wdhAdvanced_Duplex in ADocHandling) //or have SubItems?
           then DownloadAdvDuplex
           else DownloadSingleItem;
         end
     else*) DownloadSingleItem;

      { #todo 2 -oMaxM : Test if all Scanner is Synch }
      (*

      myTickStart:= GetTickCount64; curTick:= myTickStart;
      repeat
        CheckSynchronize(100);

        curTick:= GetTickCount64;

      until (rDownloaded) or ((curTick-myTickStart) > 27666);
      *)

      if (lres = S_OK) and rDownloaded
      then Result:= rDownload_Count
      else Result:= 0;
  end;
end;

function TWIADevice.Download(APath, AFileName, AExt: String; AFormat: TWIAImageFormat): Integer;
begin
  Result:= 0;

  if SetImageFormat(AFormat) then Result:= Download(APath, AFileName, AExt);
end;

function TWIADevice.Download(APath, AFileName, AExt: String; AFormat: TWIAImageFormat;
                             var DownloadedFiles: TStringArray; UseRelativePath: Boolean): Integer;
var
   i: Integer;

begin
  Result:= 0;
  DownloadedFiles:= nil;

  if SetImageFormat(AFormat) then
  begin
    Result:= Download(APath, AFileName, AExt);
    if (Result > 0 ) then
    begin
      SetLength(DownloadedFiles, Result);

      if UseRelativePath
      then begin
             DownloadedFiles[0]:= rDownload_FileName+rDownload_Ext;
             for i:=1 to Result-1 do
               DownloadedFiles[i]:= rDownload_FileName+'-'+IntToStr(i)+rDownload_Ext;
           end
      else begin
             DownloadedFiles[0]:= rDownload_Path+rDownload_FileName+rDownload_Ext;
             for i:=1 to Result-1 do
               DownloadedFiles[i]:= rDownload_Path+rDownload_FileName+'-'+IntToStr(i)+rDownload_Ext;
           end;
    end;
  end;
end;

function TWIADevice.DownloadNativeUI(hwndParent: HWND; useSystemUI: Boolean; APath, AFileName: String;
  var DownloadedFiles: TStringArray; UseRelativePath: Boolean): Integer;
var
   dlgFlags: LONG;
   i: Integer;
   filePaths: PBSTR;
   itemArray: PIWiaItem2;

begin
  Result:= 0;
  DownloadedFiles:= nil;

  if (pRootItem = nil) then GetRootItemIntf;
  if (pRootItem <> nil) then
  begin
    if (APath = '') or CharInSet(APath[Length(APath)], AllowDirectorySeparators)
    then rDownload_Path:= APath
    else rDownload_Path:= APath+DirectorySeparator;

    if not(ForceDirectories(rDownload_Path)) then exit;

    rDownload_FileName:= AFileName;
    rDownload_Ext:= '';
    rDownload_Count:= 0;
    rDownloaded:= False;

    if useSystemUI
    then dlgFlags:= WIA_DEVICE_DIALOG_USE_COMMON_UI
    else dlgFlags:= 0;

    filePaths:= nil;
    itemArray:= nil;

           //Maybe pSelectedItem. but does not works
    lres:= pRootItem.DeviceDlg(dlgFlags, hwndParent,
                               StringToOleStr(rDownload_Path), StringToOleStr(rDownload_FileName),
                               rDownload_Count, filePaths, itemArray);
(*  Alternative
    lres:= rOwner.pDevMgr.GetImageDlg(0, StringToOleStr(Self.ID), hwndParent,
                                StringToOleStr(rDownload_Path), StringToOleStr(rDownload_FileName),
                                rDownload_Count, filePaths, itemArray);
*)
     if (lres = S_OK) then
     begin
       //Copy filePaths to DownloadedFiles and Free elements
       SetLength(DownloadedFiles, rDownload_Count);
       for i:=0 to rDownload_Count-1 do
       begin
         DownloadedFiles[i]:= filePaths^[i];

         if UseRelativePath
         then FullPathToRelativePath(rDownload_Path, DownloadedFiles[i]);

         SysFreeString(filePaths^[i]);
       end;

       Result:= rDownload_Count;
     end;

     if (filePaths <> nil) then CoTaskMemFree(filePaths);
  end;
end;

function TWIADevice.GetProperty(APropId: PROPID; var propType: TVarType;
                                var APropValue; useRoot: Boolean): Boolean;
var
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   curProp: IWiaPropertyStorage;

begin
  Result:= False;

  if useRoot
  then curProp:= GetRootPropertiesIntf
  else curProp:= GetSelectedPropertiesIntf;

  if (curProp <> nil) then
  begin
       pPropSpec.ulKind:= PRSPEC_PROPID;
       pPropSpec.propid:= APropId;

       { #note -oMaxM : The Overloaded Version also call  }
       //lres:= GetPropertyAttributes(1, @pPropSpec, @pFlags, @pPropVar);

       lres:= curProp.ReadMultiple(1, @pPropSpec, @pPropVar);

       Result:= (lres = S_OK);

       if Result then
       begin
         propType:= pPropVar.vt;

         // Convert ONLY the Types Used in WIA to APropValue,
         Case propType of
         VT_I2: SmallInt(APropValue):= pPropVar.iVal; //2 byte signed int
         VT_I4, VT_INT: Integer(APropValue):= pPropVar.lVal; //4 byte signed int, signed machine int
         VT_R4: Single(APropValue):= pPropVar.fltVal; //4 byte real
         VT_R8: Double(APropValue):= pPropVar.dblVal; //8 byte real
         VT_BSTR: String(APropValue):= pPropVar.bstrVal; //OLE Automation string,
         //VT_LPSTR: String(APropValue):= pPropVar.pszVal; //null terminated string
         //VT_LPWSTR: String(APropValue):= pPropVar.pwszVal; //wide null terminated string
         VT_UI1: Byte(APropValue):= pPropVar.bVal; //unsigned AnsiChar
         VT_UI2: Word(APropValue):= pPropVar.uiVal; //unsigned short
         VT_UI4, VT_UINT : LongWord(APropValue):= pPropVar.ulVal; //unsigned long
         VT_CLSID : TGUID(APropValue):= pPropVar.puuid^; //A Class ID
         else Result:= False;
         end;
       end;
   end;
end;

function TWIADevice.GetProperty(APropId: PROPID; var propType: TVarType;
                                var APropValue, APropDefaultValue;
                                var APropListValues; useRoot: Boolean): TWIAPropertyFlags;
var
   pPropSpec: PROPSPEC;
   pPropVar,
   pPropInfo: PROPVARIANT;
   pFlags: ULONG;
   curProp: IWiaPropertyStorage;
   i: Integer;
   numElems,
   firstElem: DWord;

begin
  Result:= [];

  if useRoot
  then curProp:= GetRootPropertiesIntf
  else curProp:= GetSelectedPropertiesIntf;

  if (curProp <> nil) then
  begin
       pPropSpec.ulKind:= PRSPEC_PROPID;
       pPropSpec.propid:= APropId;

       lres:= curProp.GetPropertyAttributes(1, @pPropSpec, @pFlags, @pPropInfo);

       if (lres = S_OK) then
       begin
         propType:= pPropInfo.vt;

         Result:= WIAPropertyFlags(pFlags);

         if not(Result = []) then
         begin
           lres:= curProp.ReadMultiple(1, @pPropSpec, @pPropVar);
           propType:= pPropVar.vt;
           { #todo -oMaxM : What to do if the two types (pPropVar and pPropInfo) are different? }

           { #todo 10 -oMaxM : Convert ONLY the Types Used in WIA }

           if (lres = S_OK) then
           Case propType of
             VT_I2: begin //2 byte signed int
               SmallInt(APropValue):= pPropVar.iVal;
               numElems:= 0;
               firstElem:= 0;
               SmallInt(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.cai.cElems-WIA_LIST_VALUES; //pElems[WIA_LIST_COUNT]
                      firstElem:= WIA_LIST_VALUES;
                      SmallInt(APropDefaultValue):= pPropInfo.cai.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      SmallInt(APropDefaultValue):= pPropInfo.cai.pElems[WIA_RANGE_NOM];
                    end;
               if (WIAProp_FLAG in Result)
               then begin
                      numElems:= pPropInfo.cai.cElems-WIA_FLAG_VALUES; //WIA_FLAG_NUM_ELEMS;
                      firstElem:= WIA_FLAG_VALUES; //0;
                      SmallInt(APropDefaultValue):= pPropInfo.cai.pElems[WIA_FLAG_NOM];
                    end;

               SetLength(TArraySmallInt(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArraySmallInt(APropListValues)[i-firstElem]:= Integer(pPropInfo.cai.pElems[i]);
             end;
             VT_I4, VT_INT: begin //4 byte signed int, signed machine int
               Integer(APropValue):= pPropVar.lVal;
               numElems:= 0;
               firstElem:= 0;
               Integer(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.cal.pElems[WIA_LIST_COUNT]; //pPropInfo.cal.cElems-WIA_LIST_VALUES
                      firstElem:= WIA_LIST_VALUES;
                      Integer(APropDefaultValue):= pPropInfo.cal.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      Integer(APropDefaultValue):= pPropInfo.cal.pElems[WIA_RANGE_NOM];
                    end;
               if (WIAProp_FLAG in Result)
               then begin
                      numElems:= pPropInfo.cal.cElems-WIA_FLAG_VALUES; //WIA_FLAG_NUM_ELEMS;
                      firstElem:= WIA_FLAG_VALUES; //0;
                      Integer(APropDefaultValue):= pPropInfo.cal.pElems[WIA_FLAG_NOM];
                    end;

               SetLength(TArrayInteger(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArrayInteger(APropListValues)[i-firstElem]:= Integer(pPropInfo.cal.pElems[i]);
             end;
             VT_R4: begin //4 byte real
               Single(APropValue):= pPropVar.fltVal;
               numElems:= 0;
               firstElem:= 0;
               Single(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                 { #note -oMaxM : documentaion says to use pElems[WIA_LIST_COUNT] but is a nonsense when the type is not an Integer}
                      numElems:= pPropInfo.caflt.cElems-WIA_LIST_VALUES;
                      firstElem:= WIA_LIST_VALUES;
                      Single(APropDefaultValue):= pPropInfo.caflt.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      Single(APropDefaultValue):= pPropInfo.caflt.pElems[WIA_RANGE_NOM];
                    end;
               //if (WIAProp_FLAG in Result) ?? nonsense

               SetLength(TArraySingle(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArraySingle(APropListValues)[i-firstElem]:= Single(pPropInfo.caflt.pElems[i]);
             end;
             VT_R8: begin //8 byte real
               Double(APropValue):= pPropVar.dblVal;
               numElems:= 0;
               firstElem:= 0;
               Double(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                 { #note -oMaxM : documentaion says to use pElems[WIA_LIST_COUNT] but is a nonsense when the type is not an Integer}
                      numElems:= pPropInfo.cadbl.cElems-WIA_LIST_VALUES;
                      firstElem:= WIA_LIST_VALUES;
                      Double(APropDefaultValue):= pPropInfo.cadbl.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      Double(APropDefaultValue):= pPropInfo.cadbl.pElems[WIA_RANGE_NOM];
                    end;
               //if (WIAProp_FLAG in Result) ?? nonsense

               SetLength(TArrayDouble(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArrayDouble(APropListValues)[i-firstElem]:= Double(pPropInfo.cadbl.pElems[i]);
             end;
             VT_BSTR: begin //OLE Automation string
               String(APropValue):= pPropVar.bstrVal;
               numElems:= 0;
               firstElem:= 0;
               String(APropDefaultValue):= '';

               if (WIAProp_LIST in Result)
               then begin
                 { #note -oMaxM : documentaion says to use pElems[WIA_LIST_COUNT] but is a nonsense when the type is not an Integer}
                      numElems:= pPropInfo.cabstr.cElems-WIA_LIST_VALUES;
                      firstElem:= WIA_LIST_VALUES;
                      String(APropDefaultValue):= pPropInfo.cabstr.pElems[WIA_LIST_NOM];
                    end
               else
               //if (WIAProp_RANGE in Result) ?? nonsense
               //if (WIAProp_FLAG in Result) ?? nonsense

               SetLength(TStringArray(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TStringArray(APropListValues)[i-firstElem]:= String(pPropInfo.cabstr.pElems[i]);
             end;
             VT_UI1: begin //unsigned AnsiChar
               Byte(APropValue):= pPropVar.bVal;
               numElems:= 0;
               firstElem:= 0;
               Byte(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.caub.cElems-WIA_LIST_VALUES; //pElems[WIA_LIST_COUNT]
                      firstElem:= WIA_LIST_VALUES;
                      Byte(APropDefaultValue):= pPropInfo.caub.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      Byte(APropDefaultValue):= pPropInfo.caub.pElems[WIA_RANGE_NOM];
                    end;
               if (WIAProp_FLAG in Result)
               then begin
                      numElems:= pPropInfo.caub.cElems-WIA_FLAG_VALUES; //WIA_FLAG_NUM_ELEMS;
                      firstElem:= WIA_FLAG_VALUES;
                      Byte(APropDefaultValue):= pPropInfo.caub.pElems[WIA_FLAG_NOM];
                    end;

               SetLength(TArrayByte(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArrayByte(APropListValues)[i-firstElem]:= Byte(pPropInfo.caub.pElems[i]);
             end;
             VT_UI2: begin //unsigned short
               Word(APropValue):= pPropVar.uiVal;
               numElems:= 0;
               firstElem:= 0;
               Word(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.caui.cElems-WIA_LIST_VALUES; //pElems[WIA_LIST_COUNT]
                      firstElem:= WIA_LIST_VALUES;
                      Word(APropDefaultValue):= pPropInfo.caui.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      Word(APropDefaultValue):= pPropInfo.caui.pElems[WIA_RANGE_NOM];
                    end;
               if (WIAProp_FLAG in Result)
               then begin
                      numElems:= pPropInfo.caui.cElems-WIA_FLAG_VALUES; //WIA_FLAG_NUM_ELEMS;
                      firstElem:= WIA_FLAG_VALUES; //0;
                      Word(APropDefaultValue):= pPropInfo.caui.pElems[WIA_FLAG_NOM];
                    end;

               SetLength(TArrayWord(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArrayWord(APropListValues)[i-firstElem]:= Word(pPropInfo.caui.pElems[i]);
             end;
             VT_UI4, VT_UINT : begin //unsigned long
               LongWord(APropValue):= pPropVar.ulVal;
               numElems:= 0;
               firstElem:= 0;
               LongWord(APropDefaultValue):= 0;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.caul.cElems-WIA_LIST_VALUES; //pElems[WIA_LIST_COUNT]
                      firstElem:= WIA_LIST_VALUES;
                      LongWord(APropDefaultValue):= pPropInfo.caul.pElems[WIA_LIST_NOM];
                    end
               else
               if (WIAProp_RANGE in Result)
               then begin
                      numElems:= WIA_RANGE_NUM_ELEMS;
                      firstElem:= 0;
                      LongWord(APropDefaultValue):= pPropInfo.caul.pElems[WIA_RANGE_NOM];
                    end;
               if (WIAProp_FLAG in Result)
               then begin
                      numElems:= pPropInfo.caul.cElems-WIA_FLAG_VALUES; //WIA_FLAG_NUM_ELEMS;
                      firstElem:= WIA_FLAG_VALUES; //0;
                      LongWord(APropDefaultValue):= pPropInfo.caul.pElems[WIA_FLAG_NOM];
                    end;

               SetLength(TArrayLongWord(APropListValues), numElems);
               for i:=firstElem to firstElem+numElems-1 do
                 TArrayLongWord(APropListValues)[i-firstElem]:= LongWord(pPropInfo.caul.pElems[i]);
             end;
             VT_CLSID : begin //A Class ID
               //TGUID(APropValue):= pPropVar.puuid^; //it should be this assign but it isn't
               TGUID(APropValue):= GUID_NULL;

               numElems:= 0;
               firstElem:= 0;
               TGUID(APropDefaultValue):= GUID_NULL;

               if (WIAProp_LIST in Result)
               then begin
                      numElems:= pPropInfo.cauuid.cElems;
                      firstElem:= 0;
                      { #note -oMaxM : I don't understand the logic but in this case the WIA_LIST_XXX indexes are not valid}
                      //TGUID(APropDefaultValue):= pPropInfo.cauuid.pElems[WIA_LIST_NOM];

                      SetLength(TArrayGUID(APropListValues), numElems);
                      for i:=firstElem to firstElem+numElems-1 do
                        TArrayGUID(APropListValues)[i-firstElem]:= TGUID(pPropInfo.cauuid.pElems[i]);
                    end;
               //if (WIAProp_RANGE in Result) ?? nonsense
               //if (WIAProp_FLAG in Result) ?? nonsense
            end;
            else Result:= [];
           end;
         end;
       end;
  end;
end;

function TWIADevice.GetProperty(APropId: PROPID; var propType: TVarType;
                                var APropValue, APropDefault, APropMin, APropMax, APropStep;
                                useRoot: Boolean): Boolean;
var
   pFlags: TWIAPropertyFlags;
   iValues: Pointer;

begin
  Result:= False;
  try
     pFlags:= GetProperty(APropId, propType, APropValue, APropDefault, iValues, useRoot);

     Result:= (WIAProp_RANGE in pFlags);

     if Result then
     begin
       Case propType of
         VT_I4, VT_INT: begin
            Result:= (Length(TArrayInteger(iValues)) = WIA_RANGE_NUM_ELEMS);
            if Result then
            begin
              Integer(APropMin):= TArrayInteger(iValues)[WIA_RANGE_MIN];
              Integer(APropMax):= TArrayInteger(iValues)[WIA_RANGE_MAX];
              Integer(APropStep):= TArrayInteger(iValues)[WIA_RANGE_STEP];
            end;
         end;
       end;
     end;

  finally
    iValues:= nil;
  end;
end;

function TWIADevice.SetProperty(APropId: PROPID; propType: TVarType; const APropValue; useRoot: Boolean): Boolean;
var
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   curProp: IWiaPropertyStorage;

begin
  Result:= False;

  if useRoot
  then curProp:= GetRootPropertiesIntf
  else curProp:= GetSelectedPropertiesIntf;

  if (curProp <> nil) then
  begin
       pPropSpec.ulKind:= PRSPEC_PROPID;
       pPropSpec.propid:= APropId;
       pPropVar.vt:= propType;

       { #todo 10 -oMaxM : Convert ONLY the Types Used in WIA }
       Case propType of
         VT_I2: begin //2 byte signed int
           pPropVar.iVal:= SmallInt(APropValue);
         end;
         VT_I4, VT_INT: begin //4 byte signed int, signed machine int
           pPropVar.lVal:= Integer(APropValue);
         end;
         VT_R4: begin //4 byte real
           pPropVar.fltVal:= Single(APropValue);
         end;
         VT_R8, VT_DATE: begin //8 byte real , date
           if (propType = VT_R8)
           then pPropVar.dblVal:= Double(APropValue)
           else pPropVar.date:= Double(APropValue);
         end;
         VT_CY: begin //currency
           pPropVar.cyVal:= CURRENCY(APropValue);
         end;
         VT_BSTR, VT_LPSTR, VT_LPWSTR: begin //OLE Automation string, null terminated string, wide null terminated string
            { #note 5 -oMaxM : Test this Casts }
           case propType of
           VT_BSTR: pPropVar.bstrVal:= PWideChar(String(APropValue));
           VT_LPSTR: pPropVar.pszVal:= PAnsiChar(String(APropValue));
           VT_LPWSTR: pPropVar.pwszVal:= PWideChar(String(APropValue));
           end;
         end;
         VT_BOOL: begin //True=-1, False=0
           pPropVar.boolVal:= Boolean(APropValue);
         end;
         VT_I1: begin //signed AnsiChar
           { #note -oMaxM : Delphi has wrong declaration of cVal as ShortInt, correct one is Char }
           {$ifdef fpc}
           pPropVar.cVal:= AnsiChar(APropValue);
           {$else}
           pPropVar.cVal:= ShortInt(APropValue);
           {$endif}
         end;
         VT_UI1: begin //unsigned AnsiChar
           pPropVar.bVal:= Byte(APropValue);
         end;
         VT_UI2: begin //unsigned short
           pPropVar.uiVal:= Word(APropValue);
         end;
         VT_UI4, VT_UINT : begin //unsigned long
           pPropVar.ulVal:= LongWord(APropValue);
         end;
         VT_I8 : begin //signed 64-bit int
           pPropVar.hVal:= LARGE_INTEGER(APropValue);
         end;
         VT_UI8 : begin //unsigned 64-bit int
           pPropVar.uhVal:= ULARGE_INTEGER(APropValue);
         end;
         VT_CLSID : begin //A Class ID
           pPropVar.puuid:= @TGUID(APropValue); { #note 5 -oMaxM : Test this Cast }
         end;
     end;

     lres:= curProp.WriteMultiple(1, @pPropSpec, @pPropVar, 2);

     Result:= (lres = S_OK);
  end;
end;

function TWIADevice.GetResolutionsX(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean): TWIAPropertyFlags;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_XRES, propType, Current, Default, Values, useRoot);
  { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}
end;

function TWIADevice.GetResolutionsY(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean): TWIAPropertyFlags;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_YRES, propType, Current, Default, Values, useRoot);
  { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}
end;

function TWIADevice.GetResolutionsLimit(var AMin, AMax: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   Current,
   Default: Integer;
   pFlags: TWIAPropertyFlags;
   Values: TArrayInteger;

begin
  Result:= False;
  try
     pFlags:= GetProperty(WIA_IPS_XRES, propType, Current, Default, Values, useRoot);
     if not(WIAProp_READ in pFlags) then exit;

     if (WIAProp_RANGE in pFlags)
     then begin
            AMin:= Values[WIA_RANGE_MIN];
            AMax:= Values[WIA_RANGE_MAX];
            Result:= True;
        end
     else
     if (WIAProp_LIST in pFlags)
     then begin
            //In theory the minimum is the first value and the maximum is the last,
            //  but you never know a little paranoia doesn't hurt
            AMin:= MaxInt;
            AMax:= 0;
            for Current:=0 to Length(Values)-1 do
            begin
              if (Values[Current] < AMin) then AMin:= Values[Current];
              if (Values[Current] > AMax) then AMax:= Values[Current];
            end;
            Result:= True;
          end;

  finally
    Values:= nil;
  end;
end;

function TWIADevice.GetResolutionsLimit(var AMinX, AMaxX, AMinY, AMaxY: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   Current,
   Default: Integer;
   pFlags: TWIAPropertyFlags;
   ValuesX,
   ValuesY: TArrayInteger;

begin
  Result:= False;
  try
     pFlags:= GetProperty(WIA_IPS_XRES, propType, Current, Default, ValuesX, useRoot);
     if not(WIAProp_READ in pFlags) then exit;

     pFlags:= GetProperty(WIA_IPS_YRES, propType, Current, Default, ValuesY, useRoot);
     if not(WIAProp_READ in pFlags) then exit;

     if (WIAProp_RANGE in pFlags)
     then begin
            AMinX:= ValuesX[WIA_RANGE_MIN];
            AMaxX:= ValuesX[WIA_RANGE_MAX];
            AMinY:= ValuesY[WIA_RANGE_MIN];
            AMaxY:= ValuesY[WIA_RANGE_MAX];
            Result:= True;
        end
     else
     if (WIAProp_LIST in pFlags)
     then begin
            //In theory the minimum is the first value and the maximum is the last,
            //  but you never know a little paranoia doesn't hurt
            AMinX:= MaxInt;
            AMaxX:= 0;
            AMinY:= MaxInt;
            AMaxY:= 0;
            for Current:=0 to Length(ValuesX)-1 do
            begin
              if (ValuesX[Current] < AMinX) then AMinX:= ValuesX[Current];
              if (ValuesX[Current] > AMaxX) then AMaxX:= ValuesX[Current];
              try
                 //Y may have a different size than X, use try/except block
                 if (ValuesY[Current] < AMinY) then AMinY:= ValuesY[Current];
                 if (ValuesY[Current] > AMaxY) then AMaxY:= ValuesY[Current];
              except
              end;
            end;
            Result:= True;
          end;

  finally
    ValuesX:= nil;
    ValuesY:= nil;
  end;
end;

function TWIADevice.GetResolution(var AXRes, AYRes: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_XRES, propType, AXRes, useRoot) and
           GetProperty(WIA_IPS_YRES, propType, AYRes, useRoot);
  { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}
end;

function TWIADevice.SetResolution(const AXRes, AYRes: Integer; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPS_XRES, VT_I4, AXRes, useRoot);
  if Result
  then rXRes:= AXRes
  else Exit;

  Result:= SetProperty(WIA_IPS_YRES, VT_I4, AYRes, useRoot);
  if Result then rYRes:= AYRes;
end;

function TWIADevice.GetPaperAlign(var ALandscape: Boolean; var HAlign: TWIAAlignHorizontal; var VAlign: TWIAAlignVertical;
                                  useRoot: Boolean=False): Boolean;
begin
  HAlign:= rHAlign;
  VAlign:= rVAlign;
  Result:= GetPaperLandscape(ALandscape, useRoot);
end;

function TWIADevice.SetPaperAlign(const ALandscape: Boolean; const HAlign: TWIAAlignHorizontal; const VAlign: TWIAAlignVertical;
                                  useRoot: Boolean=False): Boolean;
begin
  rHAlign:= HAlign;
  rVAlign:= VAlign;
  Result:= SetPaperLandscape(ALandscape, useRoot);
end;

function TWIADevice.GetPaperSizeMax(var AMaxWidth, AMaxHeight: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   curSource: Integer;

begin
  if (rVersion = 1)
  then begin
         Result:= GetProperty(WIA_DPS_DOCUMENT_HANDLING_SELECT, propType, curSource, True);
         if Result then
         begin
           if (curSource and FEEDER <> 0)
           then Result:= GetProperty(WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE, propType, AMaxWidth, True) and
                         GetProperty(WIA_DPS_VERTICAL_SHEET_FEED_SIZE, propType, AMaxHeight, True)
           else Result:= GetProperty(WIA_DPS_HORIZONTAL_BED_SIZE, propType, AMaxWidth, True) and
                         GetProperty(WIA_DPS_VERTICAL_BED_SIZE, propType, AMaxHeight, True);
         end;
       end
  else Result:= GetProperty(WIA_IPS_MAX_HORIZONTAL_SIZE, propType, AMaxWidth, useRoot) and
                GetProperty(WIA_IPS_MAX_VERTICAL_SIZE, propType, AMaxHeight, useRoot);
end;

function TWIADevice.GetPaperType(var Current: TWIAPaperType; useRoot: Boolean): Boolean;
var
   iMaxWidth,
   iMaxHeight,
   iWidth,
   iHeight: Integer;
   propType: TVarType;

begin
  Result:= False;

  //WIA_IPS_PAGE_SIZE does NOT work, so we calculate the Page Size using
  // WIA_IPS_PAGE_WIDTH/WIA_IPS_PAGE_HEIGHT and PaperSizesWIA Array
  //if the user still wants to use it he can set PaperSizes_Calculated to False.

  if PaperSizes_Calculated
  then begin
        Result:= GetProperty(WIA_IPS_PAGE_WIDTH, propType, iWidth, useRoot) and
                 GetProperty(WIA_IPS_PAGE_HEIGHT, propType, iHeight, useRoot); { #note 5 -oMaxM : Always return False in HP ?}
        if Result
        then begin
               Result:= GetPaperSizeMax(iMaxWidth, iMaxHeight);
               if (iWidth >= iMaxWidth) or (iHeight >= iMaxHeight)
               then Current:= wptMAX
               else Current:= CalculatePaperSize(iWidth, iHeight);
             end
        else Current:= wptCUSTOM;

        Result:= True;
       end
  else Result:= GetProperty(WIA_IPS_PAGE_SIZE, propType, Current, useRoot);
end;

function TWIADevice.GetPaperType(var Current, Default: TWIAPaperType; var Values: TWIAPaperTypeSet; useRoot: Boolean): Boolean;
var
   iMaxWidth,
   iMaxHeight,
   i: Integer;
   intValues: TArrayInteger;
   propType: TVarType;
   pFlags: TWIAPropertyFlags;

begin
  Result:= False;
  try
     Values:= [];

     //WIA_IPS_PAGE_SIZE does NOT work, so we calculate this set using
     // WIA_IPS_MAX_HORIZONTAL/VERTICAL_SIZE and PaperSizesWIA Array
     //if the user still wants to use it he can set PaperSizes_Calculated to False.

     if PaperSizes_Calculated
     then begin
            Result:= GetPaperSizeMax(iMaxWidth, iMaxHeight);
            if not(Result) then Exit;

            Values:= CalculatePaperSizeSet(iMaxWidth, iMaxHeight);

            //Current (Reuse the Max Variables)
            Result:= GetProperty(WIA_IPS_PAGE_WIDTH, propType, iMaxWidth, useRoot) and
                     GetProperty(WIA_IPS_PAGE_HEIGHT, propType, iMaxHeight, useRoot);{ #note 5 -oMaxM : Always return False in HP ?}
            if Result
            then Current:= CalculatePaperSize(iMaxWidth, iMaxHeight)
            else Current:= wptCUSTOM;

            { #todo -oMaxM : How to Calculate Default Value??? }
            Default:= wptA4;

            Result:= True;
          end
     else begin
            pFlags:= GetProperty(WIA_IPS_PAGE_SIZE, propType, Current, Default, intValues, useRoot);
            if not(WIAProp_READ in pFlags) then Exit;

            if (WIAProp_LIST in pFlags) then
            begin
              for i:=0 to Length(intValues)-1 do Values:= Values+[TWIAPaperType(intValues[i])];

              Result:= True;
            end;
          end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetPaperType(const Value: TWIAPaperType; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   pFlags: TWIAPropertyFlags;
   MaxWidth,
   MaxHeight,
   Width,
   Height,
   X, Y,
   iPixels: Integer;
   XMux,
   YMux: Single;

begin
  Result:= False;

  if PaperSizes_Calculated
  then begin
         Result:= GetPaperSizeMax(MaxWidth, MaxHeight, useRoot);
         if not(Result) then Exit;

         Result:= False;

         if (rXRes = -1) or (rYRes = -1)
         then if not(GetResolution(rXRes, rYRes)) then  Exit;

         XMux:= rXRes / 1000;
         YMux:= rYRes / 1000;

         //if wpsMAX then assigns the entire area,
         //otherwise check if the real paper size fits within the entire area
         if (Value = wptMAX)
         then begin
                Width:= MaxWidth;
                Height:= MaxHeight;
                //we deliberately ignore Aligments because we cannot rotate/align the maximum size
                X:= 0;
                Y:= 0;
              end
         else begin
                if (rPaperLandscape)
                then begin
                       //Swap Width with Height
                       Width:= PaperSizesWIA[Value].h;
                       Height:= PaperSizesWIA[Value].w;
                     end
                else begin
                       Width:= PaperSizesWIA[Value].w;
                       Height:= PaperSizesWIA[Value].h;
                     end;

                Case rHAlign of
                  waHLeft: X:= 0;
                  waHCenter: X:= Round((MaxWidth-Width) / 2);
                  waHRight: X:= MaxWidth-Width;
                end;

                Case rVAlign of
                  waVTop: Y:= 0;
                  waVCenter: Y:= Round((MaxHeight-Height) / 2);
                  waVBottom: Y:= MaxHeight-Height;
                end;

                //I prefer to be less restrictive and do the scan
                if (X < 0) then X:= 0;
                if (X > MaxWidth) then X:= MaxWidth;
                if (Y < 0) then Y:= 0;
                if (Y > MaxHeight) then Y:= MaxHeight;

                //if (X+Width > MaxWidth) or (Y+Height > MaxHeight) then Exit; //Be restrictive
              end;

         iPixels:= Trunc(X*XMux);
         if not(SetProperty(WIA_IPS_XPOS, VT_I4, iPixels, useRoot)) then Exit;

         iPixels:= Trunc(Y*YMux);
         if not(SetProperty(WIA_IPS_YPOS, VT_I4, iPixels, useRoot)) then Exit;

         iPixels:= Trunc(Width*XMux);
         if not(SetProperty(WIA_IPS_XEXTENT, VT_I4, iPixels, useRoot)) then Exit;

         iPixels:= Trunc(Height*YMux);
         Result:= SetProperty(WIA_IPS_YEXTENT, VT_I4, iPixels, useRoot);
       end
  else begin
         if (Value = wptMAX)
         then begin
                //assigns the entire area
                Result:= GetPaperSizeMax(MaxWidth, MaxHeight, useRoot);
                if not(Result) then Exit;

                Result:= False;

                if (rXRes = -1) or (rYRes = -1)
                then if not(GetResolution(rXRes, rYRes)) then Exit;

                iPixels:= 0;
                if not(SetProperty(WIA_IPS_XPOS, VT_I4, iPixels, useRoot)) then Exit;
                if not(SetProperty(WIA_IPS_YPOS, VT_I4, iPixels, useRoot)) then Exit;

                iPixels:= Trunc(MaxWidth * rXRes / 1000);
                if not(SetProperty(WIA_IPS_XEXTENT, VT_I4, iPixels, useRoot)) then Exit;

                iPixels:= Trunc(MaxHeight * rYRes / 1000);
                Result:= SetProperty(WIA_IPS_YEXTENT, VT_I4, iPixels, useRoot);
              end
         else Result:= SetProperty(WIA_IPS_PAGE_SIZE, VT_I4, Value, useRoot);
       end;
end;

function TWIADevice.SetPaperSize(Width, Height: Integer; useRoot: Boolean): Boolean;
var
   MaxWidth,
   MaxHeight,
   swapS,
   X, Y,
   iPixels: Integer;
   XMux,
   YMux: Single;

begin
  Result:= GetPaperSizeMax(MaxWidth, MaxHeight, useRoot);
  if not(Result) then Exit;

  Result:= False;

  if (rXRes = -1) or (rYRes = -1)
  then if not(GetResolution(rXRes, rYRes)) then Exit;

  XMux:= rXRes / 1000;
  YMux:= rYRes / 1000;

  //Check if the paper size fits within the entire area
  if (rPaperLandscape) then
  begin
    //Swap Width with Height
    swapS:= Width;
    Width:= Height;
    Height:= swapS;
  end;

  Case rHAlign of
  waHLeft: X:= 0;
  waHCenter: X:= Round((MaxWidth-Width) / 2);
  waHRight: X:= MaxWidth-Width;
  end;

  Case rVAlign of
  waVTop: Y:= 0;
  waVCenter: Y:= Round((MaxHeight-Height) / 2);
  waVBottom: Y:= MaxHeight-Height;
  end;

  //I prefer to be less restrictive and do the scan
  if (X < 0) then X:= 0;
  if (X > MaxWidth) then X:= MaxWidth;
  if (Y < 0) then Y:= 0;
  if (Y > MaxHeight) then Y:= MaxHeight;

  //if (X+Width > MaxWidth) or (Y+Height > MaxHeight) then Exit; //Be restrictive

  iPixels:= Trunc(X*XMux);
  if not(SetProperty(WIA_IPS_XPOS, VT_I4, iPixels, useRoot)) then Exit;

  iPixels:= Trunc(Y*YMux);
  if not(SetProperty(WIA_IPS_YPOS, VT_I4, iPixels, useRoot)) then Exit;

  iPixels:= Trunc(Width*XMux);
  if not(SetProperty(WIA_IPS_XEXTENT, VT_I4, iPixels, useRoot)) then Exit;

  iPixels:= Trunc(Height*YMux);
  Result:= SetProperty(WIA_IPS_YEXTENT, VT_I4, iPixels, useRoot);
end;

function TWIADevice.GetPaperLandscape(var Value: Boolean; useRoot: Boolean): Boolean;
var
   rotValue: TWIARotation;

begin
  if PaperSizes_Calculated
  then begin
         Value:= rPaperLandscape;
         Result:= True;
       end
  else begin
         Result:= GetRotation(rotValue, useRoot);
         Value:= (rotValue in [wrLandscape, wrRot270]); //270 is basically landscape
       end;
end;

function TWIADevice.SetPaperLandscape(const Value: Boolean; useRoot: Boolean): Boolean;
var
   rotValue: TWIARotation;

begin
  Result:= False;

  rPaperLandscape := Value;

  if PaperSizes_Calculated
  then rotValue:= wrPortrait //Always set Portrait because we do the position calculations
  else begin
         if Value
         then rotValue:= wrLandscape
         else rotValue:= wrPortrait;
       end;

  Result:= SetRotation(rotValue, useRoot);
end;

function TWIADevice.GetRotation(var Value: TWIARotation; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_ROTATION, propType, Value, useRoot);
end;

function TWIADevice.GetRotation(var Current, Default: TWIARotation; var Values: TWIARotationSet; useRoot: Boolean): Boolean;
var
   i: Integer;
   intValues: TArrayInteger;
   propType: TVarType;
   pFlags: TWIAPropertyFlags;

begin
  Result:= False;
  try
     Values:=[];

     pFlags:= GetProperty(WIA_IPS_ROTATION, propType, Current, Default, intValues, useRoot);
     if not(WIAProp_READ in pFlags) then Exit;

     if (WIAProp_LIST in pFlags) then
     begin
       for i:=0 to Length(intValues)-1 do Values:= Values+[TWIARotation(intValues[i])];

       Result:= True;
     end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetRotation(const Value: TWIARotation; useRoot: Boolean): Boolean;
var
   iValue: Integer;

begin
  iValue:= Integer(Value); // Avoid Delphi Release Optimization Error
  Result:= SetProperty(WIA_IPS_ROTATION, VT_I4, iValue, useRoot);
end;

function TWIADevice.GetDocumentHandling(var Value: TWIADocumentHandlingSet; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   iCurrent: Integer;

begin
  Result:= GetProperty(WIA_IPS_DOCUMENT_HANDLING_SELECT, propType, iCurrent, useRoot);
  if Result then Value:= WIADocumentHandling(iCurrent);
end;

function TWIADevice.GetDocumentHandling(var Current, Default, Values: TWIADocumentHandlingSet; useRoot: Boolean): Boolean;
var
   i,
   iCurrent,
   iDefault: Integer;
   intValues: TArrayInteger;
   propType: TVarType;
   pFlags: TWIAPropertyFlags;

begin
  Result:= False;
  try
     Values:=[];

     pFlags:= GetProperty(WIA_IPS_DOCUMENT_HANDLING_SELECT, propType, iCurrent, iDefault, intValues, useRoot);
     if not(WIAProp_READ in pFlags) then Exit;

     if (WIAProp_FLAG in pFlags) then
     begin
       Current:= WIADocumentHandling(iCurrent);
       Default:= WIADocumentHandling(iDefault);
       Values:= WIADocumentHandling(TArrayInteger(intValues)[0]);

       Result:= True;
     end;

   finally
     intValues:= nil;
   end;
end;

function TWIADevice.SetDocumentHandling(const Value: TWIADocumentHandlingSet; useRoot: Boolean): Boolean;
var
   iValue: Integer;

begin
  iValue:= WIADocumentHandlingInt(Value);
  Result:= SetProperty(WIA_IPS_DOCUMENT_HANDLING_SELECT, VT_I4, iValue, useRoot);
end;

function TWIADevice.GetPages(var Current: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_PAGES, propType, Current, useRoot);
end;

function TWIADevice.GetPages(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   pFlags: TWIAPropertyFlags;
   intValues: TArrayInteger;

begin
  Result:= False;
  try
     pFlags:= GetProperty(WIA_IPS_PAGES, propType, Current, Default, intValues, useRoot);

     Result:= (WIAProp_RANGE in pFlags) and (Length(intValues) = WIA_RANGE_NUM_ELEMS);

     if Result then
     begin
       AMin:= intValues[WIA_RANGE_MIN];
       AMax:= intValues[WIA_RANGE_MAX];
       AStep:= intValues[WIA_RANGE_STEP];
     end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetPages(const Value: Integer; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPS_PAGES, VT_I4, Value, useRoot);
end;

function TWIADevice.GetBrightness(var Current: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_BRIGHTNESS, propType, Current, useRoot);
end;

function TWIADevice.GetBrightness(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   pFlags: TWIAPropertyFlags;
   intValues: TArrayInteger;

begin
  Result:= False;
  try
     pFlags:= GetProperty(WIA_IPS_BRIGHTNESS, propType, Current, Default, intValues, useRoot);
     { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}

     Result:= (WIAProp_RANGE in pFlags) and (Length(intValues) = WIA_RANGE_NUM_ELEMS);

     if Result then
     begin
       AMin:= intValues[WIA_RANGE_MIN];
       AMax:= intValues[WIA_RANGE_MAX];
       AStep:= intValues[WIA_RANGE_STEP];
     end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetBrightness(const Value: Integer; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPS_BRIGHTNESS, VT_I4, Value, useRoot);
end;

function TWIADevice.GetContrast(var Current: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPS_CONTRAST, propType, Current, useRoot);
end;

function TWIADevice.GetContrast(var Current, Default, AMin, AMax, AStep: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   pFlags: TWIAPropertyFlags;
   intValues: TArrayInteger;

begin
  Result:= False;
  try
     pFlags:= GetProperty(WIA_IPS_CONTRAST, propType, Current, Default, intValues, useRoot);
     { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}

     Result:= (WIAProp_RANGE in pFlags) and (Length(intValues) = WIA_RANGE_NUM_ELEMS);

     if Result then
     begin
       AMin:= intValues[WIA_RANGE_MIN];
       AMax:= intValues[WIA_RANGE_MAX];
       AStep:= intValues[WIA_RANGE_STEP];
     end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetContrast(const Value: Integer; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPS_CONTRAST, VT_I4, Value, useRoot);
end;

function TWIADevice.GetImageFormat(var Current: TWIAImageFormat; useRoot: Boolean): Boolean;
var
   propType: TVarType;
   gValue: TGUID;

begin
  Result:= GetProperty(WIA_IPA_FORMAT, propType, gValue, useRoot);
  if Result
  then Result:= WIAImageFormat(gValue, Current);
end;

function TWIADevice.GetImageFormat(var Current, Default: TWIAImageFormat; var Values: TWIAImageFormatSet; useRoot: Boolean): Boolean;
var
   i: Integer;
   gValues: TArrayGUID;
   propType: TVarType;
   pFlags: TWIAPropertyFlags;
   curValue: TWIAImageFormat;
   gValue: TGUID;

begin
  Result:= False;
  try
     Values:= [];

     //Does not return the Current and Default Values
     pFlags:= GetProperty(WIA_IPA_FORMAT, propType, Current, Default, gValues, useRoot);
     if not(WIAProp_READ in pFlags) then Exit;

     if (WIAProp_LIST in pFlags) then
     for i:=0 to Length(gValues)-1 do
     begin
       Result:= WIAImageFormat(gValues[i], curValue);
       if Result
       then Values:= Values+[curValue];
       { #todo 2 -oMaxM : else Ignore it or return False? }
     end;

     //Default Values are not valid so we must take it in this way
     Current:= wifUNDEFINED;
     Default:= wifUNDEFINED;
     Result:= GetProperty(WIA_IPA_FORMAT, propType, gValue, useRoot) and
              WIAImageFormat(gValue, Current);
     if not(Result) then exit;

     Result:= GetProperty(WIA_IPA_PREFERRED_FORMAT, propType, gValue, useRoot) and
              WIAImageFormat(gValue, Default);

  finally
    gValues:= nil;
  end;
end;

function TWIADevice.SetImageFormat(const Value: TWIAImageFormat; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPA_FORMAT, VT_CLSID, WiaImageFormatGUID[Value], useRoot);
end;

function TWIADevice.GetDataType(var Current: TWIADataType; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPA_DATATYPE, propType, Current, useRoot);
end;

function TWIADevice.GetDataType(var Current, Default: TWIADataType; var Values: TWIADataTypeSet; useRoot: Boolean): Boolean;
var
   i: Integer;
   intValues: TArrayInteger;
   propType: TVarType;
   pFlags: TWIAPropertyFlags;

begin
  Result:= False;
  try
     Values:= [];

     pFlags:= GetProperty(WIA_IPA_DATATYPE, propType, Current, Default, intValues, useRoot);
     if not(WIAProp_READ in pFlags) then Exit;

     { #note 5 -oMaxM : what to do if the propType is not the expected one VT_I4}

     if (WIAProp_LIST in pFlags) then
     begin
       for i:=0 to Length(intValues)-1 do Values:= Values+[TWIADataType(intValues[i])];

       Result:= True;
     end;

  finally
    intValues:= nil;
  end;
end;

function TWIADevice.SetDataType(const Value: TWIADataType; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPA_DATATYPE, VT_I4, Value, useRoot);
end;

function TWIADevice.GetBitDepth(var Current, Default: Integer; var Values: TArrayInteger; useRoot: Boolean): TWIAPropertyFlags;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPA_DEPTH, propType, Current, Default, Values, useRoot);
end;

function TWIADevice.GetBitDepth(var Current: Integer; useRoot: Boolean): Boolean;
var
   propType: TVarType;

begin
  Result:= GetProperty(WIA_IPA_DEPTH, propType, Current, useRoot);
end;

function TWIADevice.SetBitDepth(const Value: Integer; useRoot: Boolean): Boolean;
begin
  Result:= SetProperty(WIA_IPA_DEPTH, VT_I4, Value, useRoot);
end;

function TWIADevice.GetParamsCapabilities(var Value: TWIAParamsCapabilities): Boolean;
var
   pFlags: TWIAPropertyFlags;
   subItemArray: TArrayWIAItem;

begin
  Result:= False;
  FillChar(Value, SizeOf(TWIAParamsCapabilities), 0);

  with Value do
  if (GetSelectedItemIntf <> nil) then
  begin
    if (rItemList[rSelectedItemIndex].ItemCategory <> wicAUTO) then
    begin
      Result:= GetPaperSizeMax(PaperSizeMaxWidth, PaperSizeMaxHeight);
      //if not(Result) then raise Exception.Create('GetPaperSizeMax');

      Result:= GetPaperType(PaperTypeCurrent, PaperTypeDefault, PaperTypeSet);
      //if not(Result) then raise Exception.Create('GetPaperType');

      Result:= GetRotation(RotationCurrent, RotationDefault, RotationSet);
      //if not(Result) then raise Exception.Create('GetRotation');

      pFlags:= GetResolutionsX(ResolutionCurrent, ResolutionDefault, ResolutionArray);
      Result:= (WIAProp_READ in pFlags);
      //if not(Result) then raise Exception.Create('GetResolutionsX');
      ResolutionRange:= WIAProp_RANGE in pFlags;

      Result:= GetBrightness(BrightnessCurrent, BrightnessDefault, BrightnessMin, BrightnessMax, BrightnessStep);
      //if not(Result) then raise Exception.Create('GetBrightness');

      Result:= GetContrast(ContrastCurrent, ContrastDefault, ContrastMin, ContrastMax, ContrastStep);
      //if not(Result) then raise Exception.Create('GetContrast');

      (*
      pFlags:= GetBitDepth(BitDepthCurrent, BitDepthDefault, BitDepthArray);
      Result:= (WIAProp_READ in pFlags);
      if not(Result) then exit;
      *)

      Result:= GetDataType(DataTypeCurrent, DataTypeDefault, DataTypeSet);
      //if not(Result) then raise Exception.Create('GetDataType');

      //Take the specific Capabilities according to the category (from this point onwards Result is True)
      Case rItemList[rSelectedItemIndex].ItemCategory of
        wicFEEDER: begin
          { #todo 5 -oMaxM : Must be tested in a Duplex Scanner }
          Result:= GetDocumentHandling(DocHandlingCurrent, DocHandlingDefault, DocHandlingSet);
          //if not(Result) then exit;

          (*Result:= GetSelectedItemSubItems(subItemArray); //return nil if non duplex
          if Result
          then begin
                 DocHandlingSet:= DocHandlingSet+[wdhAdvanced_Duplex];
               end
          else if (subItemArray = nil) then
               begin

               end;*)
        end
        else begin
             end;
      end;
    end;

    Result:= True;

  end;
end;

function TWIADevice.SetParams(const AParams: TWIAParams): Boolean;
begin
  Result:= False;

  with AParams do
  if (GetSelectedItemIntf <> nil) then
  begin
    if (rItemList[rSelectedItemIndex].ItemCategory <> wicAUTO) then
    begin
      Result:= SetResolution(Resolution, Resolution);
      //if not(Result) then raise Exception.Create('SetResolution');

      Result:= SetPaperAlign((Rotation in [wrLandscape, wrRot270]), HAlign, VAlign);
      //if not(Result) then raise Exception.Create('SetPaperAlign');

      if (PaperType = wptCUSTOM)
      then Result:= SetPaperSize(PaperW, PaperH)
      else Result:= SetPaperType(PaperType);
      //if not(Result) then raise Exception.Create('SetPaperType');

      //Set only if we are in Wia 1, in Wia 2 the user must specify what to download
      //from the feeder in the Download method using the ADocHandling parameter.
      if (*(rVersion = 1) and*) (rItemList[rSelectedItemIndex].ItemCategory = wicFEEDER) then
      begin
        Result:= SetDocumentHandling(DocHandling);
        //if not(Result) then raise Exception.Create('SetDocumentHandling');
      end;

      Result:= SetBrightness(Brightness);
      //if not(Result) then raise Exception.Create('SetBrightness');

      Result:= SetContrast(Contrast);
      //if not(Result) then raise Exception.Create('SetContrast');

      (*
      Result:= WIASource.SetBitDepth(BitDepth);
      if not(Result) then raise Exception.Create('SetBitDepth');
      *)

      Result:= SetDataType(DataType);
      //if not(Result) then raise Exception.Create('SetDataType');
    end
    else Result:= True;
  end;
end;

function TWIADevice.GetSelectedItemSubItems(var AItemArray: TArrayWIAItem): Boolean;
var
   pIEnumItem: IEnumWiaItem2;
   pItem: IWiaItem2;
   iCount,
   itemFetched: ULONG;
   itemType: LONG;
   itemCategory: TGUID;
   i: Integer;
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   pWiaPropertyStorage: IWiaPropertyStorage;

begin
  Result :=False;
  try
  AItemArray:= nil;

  if (GetSelectedItemIntf = nil) then exit;

  lres:= pSelectedItem.EnumChildItems(nil, pIEnumItem);
  if (lres = S_OK) then
  begin
    lres:= pIEnumItem.GetCount(iCount);
    if (lres = S_OK) then
    begin
      SetLength(AItemArray, iCount);

      for i:=0 to iCount-1 do
      begin
        lres:= pIEnumItem.Next(1, pItem, itemFetched);

        Result := (lres = S_OK);
        if Result then
        begin
          lres:= pItem.GetItemType(itemType);
          if (lres = S_OK)
          then AItemArray[i].ItemType :=WIAItemTypes(itemType)
          else AItemArray[i].ItemType :=[];

          lres:= pItem.GetItemCategory(itemCategory);
          if (lres = S_OK)
          then AItemArray[i].ItemCategory :=WIAItemCategory(itemCategory)
          else AItemArray[i].ItemCategory :=wicNULL;

          lres:= pItem.QueryInterface(IID_IWiaPropertyStorage, pWiaPropertyStorage);
          if (lres = S_OK) and (pWiaPropertyStorage <> nil) then
          begin
            pPropSpec.ulKind := PRSPEC_PROPID;
            pPropSpec.propid := WIA_IPA_ITEM_NAME; //WIA_IPA_FULL_ITEM_NAME

            lres := pWiaPropertyStorage.ReadMultiple(1, @pPropSpec, @pPropVar);

            if (VT_BSTR = pPropVar.vt)
            then AItemArray[i].Name :=pPropVar.bstrVal
            else AItemArray[i].Name :='?';
          end;
        end
        else break;
      end;

      Result :=True;
    end;

    pIEnumItem:= nil;
  end;

  except
    AItemArray:= nil;
    Result:= False;
  end;
end;

(*
function TWIADevice.SetProperty(APropId: PROPID; APropValue: Smallint; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Integer; useRoot: Boolean): Boolean;
var
   pPropSpec: PROPSPEC;
   pPropVar: PROPVARIANT;
   curProp: IWiaPropertyStorage;

begin
  try
     Result:= False;

     if useRoot
     then curProp:= GetRootProperties
     else curProp:= GetSelectedProperties;

     if (curProp <> nil) then
     begin
       pPropSpec.ulKind:= PRSPEC_PROPID;
       pPropSpec.propid:= APropId;
       pPropVar.vt:= VT_INT;   //VT_I4?
       pPropVar.lVal:= APropValue;    //intVal

       lres:= curProp.WriteMultiple(1, @pPropSpec, @pPropVar, 2);

       Result:= (lres = S_OK);
     end;
  except
  end;
end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Single; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Double; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Currency; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: TDateTime; useRoot: Boolean): Boolean;
begin
    { #note 3 -oMaxM : use ActiveX.DATE = DOUBLE }
end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: BSTR; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Boolean; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Word; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: DWord; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: Int64; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: UInt64; useRoot: Boolean): Boolean;
begin

end;

function TWIADevice.SetProperty(APropId: PROPID; APropValue: LPSTR; useRoot: Boolean): Boolean;
begin

end;
*)

{ TWIAManager }

function TWIAManager.GetDevMgrIntf: WIA_LH.IWiaDevMgr2;
begin
  if (pDevMgr = nil)
  then pDevMgr:= WIA_LH.IWiaDevMgr2(CreateDevManager);

  Result:=pDevMgr;
end;

procedure TWIAManager.SetEnumAll(AValue: Boolean);
begin
  if (rEnumAll <> AValue) then
  begin
    rEnumAll:= AValue;
    HasEnumerated:= EnumerateDevices;
  end;
end;

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
  //Enumerate devices if needed
  if not(HasEnumerated)
  then HasEnumerated:= EnumerateDevices;

  if (Index >= 0) and (Index < Length(rDeviceList))
  then Result:= rDeviceList[Index]
  else Result:= nil;
end;

function TWIAManager.GetDevicesCount: Integer;
begin
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
  //SetLength(rDeviceList, 0);
  EmptyDeviceList(True);

  try
    if (pDevMgr = nil)
    then pDevMgr:= WIA_LH.IWiaDevMgr2(CreateDevManager);

    if pDevMgr<>nil then
    begin
      if EnumAll
      then lres :=pDevMgr.EnumDeviceInfo(WIA_DEVINFO_ENUM_ALL, ppIEnum)
      else lres :=pDevMgr.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL, ppIEnum);

      if (lres=S_OK) and (ppIEnum<>nil) then
      begin
        lres :=ppIEnum.GetCount(devCount);

        if (lres<>S_OK)
        then Exception.Create('Number of WIA Devices not available');

        EmptyDeviceList(True);

        if (devCount > 0) then
        begin
          SetLength(rDeviceList, devCount);
          //EmptyDeviceList(False);

          // Define which properties you want to read:
          // Device ID.  This is what you would use to create
          // the device.
          pPropSpec[0].ulKind := PRSPEC_PROPID;
          pPropSpec[0].propid := WIA_DIP_DEV_ID;
          //pPropIDS[0] :=WIA_DIP_DEV_ID;

          // Device Manufacturer
          pPropSpec[1].ulKind := PRSPEC_PROPID;
          pPropSpec[1].propid := WIA_DIP_VEND_DESC;
          //pPropIDS[1] :=WIA_DIP_VEND_DESC;

          // Device Name
          pPropSpec[2].ulKind := PRSPEC_PROPID;
          pPropSpec[2].propid := WIA_DIP_DEV_NAME;
          //pPropIDS[2] :=WIA_DIP_DEV_NAME;

          // Device Type
          pPropSpec[3].ulKind := PRSPEC_PROPID;
          pPropSpec[3].propid := WIA_DIP_DEV_TYPE;
          //pPropIDS[3] :=WIA_DIP_DEV_TYPE;

          // Device Wia Version
          pPropSpec[4].ulKind := PRSPEC_PROPID;
          pPropSpec[4].propid := WIA_DIP_WIA_VERSION;
          //pPropIDS[4] :=WIA_DIP_WIA_VERSION;

          pWiaPropertyStorage :=nil;

          for i:=0 to devCount-1 do
          begin
            FillChar(pPropVar, Sizeof(pPropVar), 0);
            //FillChar(pPropNames, Sizeof(pPropNames), 0);

            lres :=ppIEnum.Next(1, pWiaPropertyStorage, devFetched);

            if (lres<>S_OK)
            then Exception.Create('pWiaPropertyStorage for Device '+IntToStr(i)+' not available');

            // Ask for the property values
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
            then rDeviceList[i].rType :=TWIADeviceType(pPropVar[3].iVal)
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
        end;

        ppIEnum :=nil;
      end;

      Result :=True;
    end;

  except
    EmptyDeviceList(True);
    Result :=False;
  end;
end;

function TWIAManager.SelectDeviceDialog: Integer;
begin
  try
     Result:= TWIASelectForm.Execute(Self);
     FreeAndNil(WIASelectForm);

  except
    Result:= -1;
  end;
end;

constructor TWIAManager.Create(AEnumAll: Boolean);
begin
  inherited Create;

  HasEnumerated:= False;
  pDevMgr:= nil;
  rSelectedDeviceIndex:= -1;
  rEnumAll:= AEnumAll;
end;

destructor TWIAManager.Destroy;
begin
  EmptyDeviceList(True);
  if (pDevMgr<>nil) then pDevMgr :=nil; //Free the Interface

  inherited Destroy;
end;

procedure TWIAManager.ClearDeviceList;
begin
  EmptyDeviceList(True);
end;

procedure TWIAManager.RefreshDeviceList;
begin
  HasEnumerated:= EnumerateDevices;
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

procedure TWIAManager.SelectDeviceItem(ADeviceID, ADeviceItem: String; var ADevice: TWIADevice; var AIndex: Integer);
var
   i      :Integer;
   curDev :TWIADevice;

begin
  ADevice:= nil;
  AIndex:= -1;

  try
    for i:=0 to DevicesCount-1 do
    begin
      curDev:= Devices[i];

      if (curDev <> nil) and
         (curDev.ID = ADeviceID) then
      begin
        ADevice:= curDev;

        if curDev.SelectItem(ADeviceItem)
        then AIndex:= ADevice.SelectedItemIndex
        else AIndex:= -1;

        break;
      end;
    end;

  except
    ADevice:= nil;
    AIndex:= -1;
  end;
end;

end.

