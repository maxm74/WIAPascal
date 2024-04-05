(****************************************************************************
*
*
*  FILE: WIA_LH.pas
*
*  VERSION:     1.0
*
*  DESCRIPTION:
*    Definitions for the WIA interfaces.
*
*****************************************************************************
*
*  2024-04-05 Translation and adaptation by Massimo Magnano
*
*****************************************************************************)

unit WIA_LH;

{$MODE DELPHI}

interface

uses Windows, ActiveX, WiaDef;

const
  IID_IWiaDevMgr : TGUID = '{5eb2502a-8cf1-11d1-bf92-0060081ed811}';
  IID_IEnumWIA_DEV_INFO : TGUID = '{5e38b83c-8cf1-11d1-bf92-0060081ed811}';
  IID_IWiaEventCallback : TGUID = '{ae6287b0-0084-11d2-973b-00a0c9068f2e}';
  IID_IWiaDataCallback : TGUID = '{a558a866-a5b0-11d2-a08f-00c04f72dc3c}';
  IID_IWiaDataTransfer : TGUID = '{a6cef998-a5b0-11d2-a08f-00c04f72dc3c}';

  IID_IWiaItem : TGUID = '{4db1ad10-3391-11d2-9a33-00c04fa36145}';
  IID_IWiaItem2 : TGUID = '{6CBA0075-1287-407d-9B77-CF0E030435CC}';

type
  IWiaDevMgr = interface;
  IEnumWIA_DEV_INFO = interface;
  IWiaEventCallback = interface;
  IWiaDataCallback = interface;
  IWiaDataTransfer = interface;
  IWiaItem = interface;
  IWiaPropertyStorage = interface;
  IEnumWiaItem = interface;
  IEnumWIA_DEV_CAPS = interface;
  IEnumWIA_FORMAT_INFO = interface;
  IWiaLog = interface;
  IWiaLogEx = interface;
  IWiaNotifyDevMgr = interface;
  IWiaItemExtras = interface;
  IWiaAppErrorHandler = interface;
  IWiaErrorHandler = interface;
  IWiaTransfer = interface;
  IWiaTransferCallback = interface;
  IWiaSegmentationFilter = interface;
  IWiaImageFilter = interface;
  IWiaPreview = interface;
  IEnumWiaItem2 = interface;
  IWiaItem2 = interface;
  IWiaDevMgr2 = interface;
  WiaDevMgr = class;
  WiaDevMgr2 = class;
  WiaLog = class;

  TarrayBSTR = array [0..0] of BSTR;
  PBSTR = ^TarrayBSTR; //^BSTR;

  TarrayIWiaItem = array [0..0] of IWiaItem;
  PIWiaItem = ^TarrayIWiaItem; //^IWiaItem;

  _WIA_DITHER_PATTERN_DATA = record
      lSize : LONG;
      bstrPatternName : BSTR;
      lPatternWidth : LONG;
      lPatternLength : LONG;
      cbPattern : LONG;
      pbPattern : PBYTE;
  end;
  WIA_DITHER_PATTERN_DATA = _WIA_DITHER_PATTERN_DATA;
  PWIA_DITHER_PATTERN_DATA = ^WIA_DITHER_PATTERN_DATA;

  _WIA_PROPID_TO_NAME = record
      propid : PROPID;
      pszName : LPOLESTR;
  end;
  WIA_PROPID_TO_NAME = _WIA_PROPID_TO_NAME;
  PWIA_PROPID_TO_NAME = ^WIA_PROPID_TO_NAME;

  _WIA_FORMAT_INFO = record
      guidFormatID : GUID;
      lTymed : LONG;
    end;
  WIA_FORMAT_INFO = _WIA_FORMAT_INFO;
  PWIA_FORMAT_INFO = ^WIA_FORMAT_INFO;

  IWiaDevMgr = interface(IUnknown)
  ['{5eb2502a-8cf1-11d1-bf92-0060081ed811}']
    function EnumDeviceInfo(lFlags: LONG;
                            var ppIEnum: IEnumWIA_DEV_INFO): HRESULT; stdcall;

    function CreateDevice(bstrDeviceID: BSTR;
                          var ppWiaItemRoot: IWiaItem): HRESULT; stdcall;

    function SelectDeviceDlg(hwndParent: HWND;
                             lDeviceType,
                             lFlags: LONG;
                             var pbstrDeviceID: BSTR;
                             var ppWiaItemRoot: IWiaItem): HRESULT; stdcall;

    function SelectDeviceDlgID(hwndParent: HWND;
                               lDeviceType,
                               lFlags: LONG;
                               var pbstrDeviceID: BSTR): HRESULT; stdcall;

    function GetImageDlg(hwndParent: HWND;
                         lDeviceType,
                         lFlags,
                         lIntent: LONG;
                         pItemRoot: IWiaItem;
                         bstrFilename: BSTR;
                         var pguidFormat: TGUID): HRESULT; stdcall;

    function RegisterEventCallbackProgram(lFlags: LONG;
                                          bstrDeviceID: BSTR;
                                          const pEventGUID: PGUID;
                                          bstrCommandline,
                                          bstrName,
                                          bstrDescription,
                                          bstrIcon: BSTR): HRESULT; stdcall;

    function RegisterEventCallbackInterface(lFlags: LONG;
                                            bstrDeviceID: BSTR;
                                            const pEventGUID: PGUID;
                                            pIWiaEventCallback: IWiaEventCallback;
                                            var pEventObject: IUnknown): HRESULT; stdcall;

    function RegisterEventCallbackCLSID(lFlags: LONG;
                                        bstrDeviceID: BSTR;
                                        const pEventGUID: PGUID;
                                        const pClsID: PGUID;
                                        bstrName,
                                        bstrDescription,
                                        bstrIcon: BSTR): HRESULT; stdcall;

    function AddDeviceDlg(hwndParent: HWND;
                          lFlags: LONG): HRESULT; stdcall;
  end;

  IEnumWIA_DEV_INFO = interface(IUnknown)
  ['{5e38b83c-8cf1-11d1-bf92-0060081ed811}']
    function Next(celt: ULONG;
                  var rgelt: IWiaPropertyStorage;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(var ppIEnum: IEnumWIA_DEV_INFO): HRESULT; stdcall;

    function GetCount(var celt: ULONG): HRESULT; stdcall;
  end;

  IWiaEventCallback = interface(IUnknown)
  ['{ae6287b0-0084-11d2-973b-00a0c9068f2e}']
    function ImageEventCallback(const pEventGUID: PGUID;
                                bstrEventDescription,
                                bstrDeviceID,
                                bstrDeviceDescription: BSTR;
                                dwDeviceType: DWORD;
                                bstrFullItemName: BSTR;
                                var pulEventType: ULONG;
                                ulReserved: ULONG): HRESULT; stdcall;

  end;

  _WIA_DATA_CALLBACK_HEADER = record
      lSize : LONG;
      guidFormatID : GUID;
      lBufferSize : LONG;
      lPageCount : LONG;
  end;
  WIA_DATA_CALLBACK_HEADER = _WIA_DATA_CALLBACK_HEADER;
  PWIA_DATA_CALLBACK_HEADER = ^WIA_DATA_CALLBACK_HEADER;

  IWiaDataCallback = interface(IUnknown)
  ['{ae6287b0-0084-11d2-973b-00a0c9068f2e}']
    function BandedDataCallback(lMessage,
                                lStatus,
                                lPercentComplete,
                                lOffset,
                                lLength,
                                lReserved,
                                lResLength: LONG;
                                pbBuffer: PBYTE): HRESULT; stdcall;

  end;

  _WIA_DATA_TRANSFER_INFO = record
      ulSize : ULONG;
      ulSection : ULONG;
      ulBufferSize : ULONG;
      bDoubleBuffer : BOOL;
      ulReserved1 : ULONG;
      ulReserved2 : ULONG;
      ulReserved3 : ULONG;
    end;
  WIA_DATA_TRANSFER_INFO = _WIA_DATA_TRANSFER_INFO;
  PWIA_DATA_TRANSFER_INFO = ^WIA_DATA_TRANSFER_INFO;

  _WIA_EXTENDED_TRANSFER_INFO = record
      ulSize : ULONG;
      ulMinBufferSize : ULONG;
      ulOptimalBufferSize : ULONG;
      ulMaxBufferSize : ULONG;
      ulNumBuffers : ULONG;
    end;
  WIA_EXTENDED_TRANSFER_INFO = _WIA_EXTENDED_TRANSFER_INFO;
  PWIA_EXTENDED_TRANSFER_INFO = ^WIA_EXTENDED_TRANSFER_INFO;

  IWiaDataTransfer = interface(IUnknown)
  ['{a6cef998-a5b0-11d2-a08f-00c04f72dc3c}']
    function idtGetData(var pMedium: STGMEDIUM;
                        pIWiaDataCallback: IWiaDataCallback): HRESULT; stdcall;

    function idtGetBandedData(const pWiaDataTransInfo: PWIA_DATA_TRANSFER_INFO;
                              pIWiaDataCallback: IWiaDataCallback): HRESULT; stdcall;

    function idtQueryGetData(const pfe: PWIA_FORMAT_INFO): HRESULT; stdcall;

    function idtEnumWIA_FORMAT_INFO(var ppEnum: IEnumWIA_FORMAT_INFO): HRESULT; stdcall;

    function idtGetExtendedTransferInfo(var pExtendedTransferInfo: WIA_EXTENDED_TRANSFER_INFO): HRESULT; stdcall;
  end;

  IWiaItem = interface(IUnknown)
  ['{4db1ad10-3391-11d2-9a33-00c04fa36145}']
    function GetItemType(var pItemType: LONG): HRESULT; stdcall;

    function AnalyzeItem(lFlags: LONG): HRESULT; stdcall;

    function EnumChildItems(var ppIEnumWiaItem: IEnumWiaItem): HRESULT; stdcall;

    function DeleteItem(lFlags: LONG): HRESULT; stdcall;

    function CreateChildItem(lFlags: LONG;
                             bstrItemName: BSTR;
                             bstrFullItemName: BSTR;
                             var ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function EnumRegisterEventInfo(lFlags: LONG;
                                   const pEventGUID: PGUID;
                                   var ppIEnum: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function FindItemByName(lFlags: LONG;
                            bstrFullItemName: BSTR;
                            var ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function DeviceDlg(hwndParent: HWND;
                       lFlags,
                       lIntent: LONG;
                       var plItemCount: LONG;
                       var ppIWiaItem: PIWiaItem): HRESULT; stdcall;

    function DeviceCommand(lFlags: LONG;
                           const pCmdGUID: PGUID;
                           var pIWiaItem: IWiaItem): HRESULT; stdcall;

    function GetRootItem(var ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function EnumDeviceCapabilities(lFlags: LONG;
                                    var ppIEnumWIA_DEV_CAPS: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function DumpItemData(var bstrData: BSTR): HRESULT; stdcall;

    function DumpDrvItemData(var bstrData: BSTR): HRESULT; stdcall;

    function DumpTreeItemData(var bstrData: BSTR): HRESULT; stdcall;

    function Diagnostic(ulSize: ULONG;
                        pBuffer: PBYTE): HRESULT; stdcall;
  end;


{ #todo 10 -oMaxM : Continua Line 1530 of .h MIDL_INTERFACE("98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6") }

  IWiaItem2 = interface(IUnknown)
  ['{6CBA0075-1287-407d-9B77-CF0E030435CC}']
    function CreateChildItem(lItemFlags, lCreationFlags: LONG;
                             bstrItemName: BSTR;
                             var ppIWiaItem2: IWiaItem2): HRESULT; stdcall;
    function DeleteItem(lFlags: LONG): HRESULT; stdcall;
    function EnumChildItems(const pCategoryGUID: PGUID;
                            var ppIEnumWiaItem2: IEnumWiaItem2): HRESULT; stdcall;
    function FindItemByName(lFlags: LONG;
                            bstrFullItemName: BSTR;
                            var ppIWiaItem2: IWiaItem2): HRESULT; stdcall;
    function GetItemCategory(var pItemCategoryGUID: TGUID): HRESULT; stdcall;
    function GetItemType(var pItemType: LONG): HRESULT; stdcall;
    function DeviceDlg(lFlags: LONG;
                       hwndParent: HWND;
                       bstrFolderName, bstrFilename: BSTR;
                       var plNumFiles: LONG; var ppbstrFilePaths: PBSTR;
                       var ppItem: IWiaItem2): HRESULT; stdcall;
    function DeviceCommand(lFlags: LONG;
                           const pCmdGUID: PGUID;
                           var pIWiaItem2: IWiaItem2): HRESULT; stdcall;
    function EnumDeviceCapabilities(lFlags: LONG;
                                    var ppIEnumWIA_DEV_CAPS: IEnumWIA_DEV_CAPS): HRESULT; stdcall;
    function CheckExtension(lFlags: LONG;
                            bstrName: BSTR;
                            riidExtensionInterface: REFIID;
                            var pbExtensionExists: BOOL): HRESULT; stdcall;
    function GetExtension(lFlags: LONG;
                          bstrName: BSTR;
                          riidExtensionInterface: REFIID;
                          var ppOut: PInterface): HRESULT; stdcall;
    function GetParentItem(var ppIWiaItem2: IWiaItem2): HRESULT; stdcall;
    function GetRootItem(var ppIWiaItem2: IWiaItem2): HRESULT; stdcall;
    function GetPreviewComponent(lFlags: LONG;
                                 var ppWiaPreview: IWiaPreview): HRESULT; stdcall;
    function EnumRegisterEventInfo(lFlags: LONG;
                                   const pEventGUID: PGUID;
                                   var ppIEnum: IEnumWIA_DEV_CAPS): HRESULT; stdcall;
    function Diagnostic(ulSize: ULONG; pBuffer: Pointer): HRESULT; stdcall;
  end;

implementation

end.

