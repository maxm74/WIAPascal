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
  IID_IWiaPropertyStorage : TGUID = '{98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6}';
  IID_IEnumWiaItem : TGUID = '{5e8383fc-3391-11d2-9a33-00c04fa36145}';
  IID_IEnumWIA_DEV_CAPS : TGUID = '{1fcc4287-aca6-11d2-a093-00c04f72dc3c}';
  IID_IEnumWIA_FORMAT_INFO : TGUID = '{81BEFC5B-656D-44f1-B24C-D41D51B4DC81}';
  IID_IWiaLog : TGUID = '{A00C10B6-82A1-452f-8B6C-86062AAD6890}';
  IID_IWiaLogEx : TGUID = '{AF1F22AC-7A40-4787-B421-AEb47A1FBD0B}';
  IID_IWiaNotifyDevMgr : TGUID = '{70681EA0-E7BF-4291-9FB1-4E8813A3F78E}';
  IID_IWiaItemExtras : TGUID = '{6291ef2c-36ef-4532-876a-8e132593778d}';
  IID_IWiaAppErrorHandler : TGUID = '{6C16186C-D0A6-400c-80F4-D26986A0E734}';
  IID_IWiaErrorHandler : TGUID = '{0e4a51b1-bc1f-443d-a835-72e890759ef3}';
  IID_IWiaTransfer : TGUID = '{c39d6942-2f4e-4d04-92fe-4ef4d3a1de5a}';
  IID_IWiaTransferCallback : TGUID = '{27d4eaaf-28a6-4ca5-9aab-e678168b9527}';
  IID_IWiaSegmentationFilter : TGUID = '{EC46A697-AC04-4447-8F65-FF63D5154B21}';
  IID_IWiaImageFilter : TGUID = '{A8A79FFA-450B-41f1-8F87-849CCD94EBF6}';
  IID_IWiaPreview : TGUID = '{95C2B4FD-33F2-4d86-AD40-9431F0DF08F7}';
  IID_IEnumWiaItem2 : TGUID = '{59970AF4-CD0D-44d9-AB24-52295630E582}';
  IID_IWiaItem2 : TGUID = '{6CBA0075-1287-407d-9B77-CF0E030435CC}';
  IID_IWiaDevMgr2 : TGUID = '{79C07CF1-CBDD-41ee-8EC3-F00080CADA7A}';

  LIBID_WiaDevMgr : TGUID = '{a1f4e726-8cf1-11d1-bf92-0060081ed811}';
  CLSID_WiaDevMgr : TGUID = '{a1f4e726-8cf1-11d1-bf92-0060081ed811}';
  CLSID_WiaDevMgr2 : TGUID = '{B6C292BC-7C88-41ee-8B54-8EC92617E599}';
  CLSID_WiaLog : TGUID = '{A1E75357-881A-419e-83E2-BB16DB197C68}';

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
                            out ppIEnum: IEnumWIA_DEV_INFO): HRESULT; stdcall;

    function CreateDevice(bstrDeviceID: BSTR;
                          out ppWiaItemRoot: IWiaItem): HRESULT; stdcall;

    function SelectDeviceDlg(hwndParent: HWND;
                             lDeviceType,
                             lFlags: LONG;
                             var pbstrDeviceID: BSTR;
                             out ppWiaItemRoot: IWiaItem): HRESULT; stdcall;

    function SelectDeviceDlgID(hwndParent: HWND;
                               lDeviceType,
                               lFlags: LONG;
                               out pbstrDeviceID: BSTR): HRESULT; stdcall;

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
                                            out pEventObject: IUnknown): HRESULT; stdcall;

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
                  out rgelt: IWiaPropertyStorage;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(out ppIEnum: IEnumWIA_DEV_INFO): HRESULT; stdcall;

    function GetCount(out celt: ULONG): HRESULT; stdcall;
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

    function idtEnumWIA_FORMAT_INFO(out ppEnum: IEnumWIA_FORMAT_INFO): HRESULT; stdcall;

    function idtGetExtendedTransferInfo(out pExtendedTransferInfo: WIA_EXTENDED_TRANSFER_INFO): HRESULT; stdcall;
  end;

  IWiaItem = interface(IUnknown)
  ['{4db1ad10-3391-11d2-9a33-00c04fa36145}']
    function GetItemType(out pItemType: LONG): HRESULT; stdcall;

    function AnalyzeItem(lFlags: LONG): HRESULT; stdcall;

    function EnumChildItems(out ppIEnumWiaItem: IEnumWiaItem): HRESULT; stdcall;

    function DeleteItem(lFlags: LONG): HRESULT; stdcall;

    function CreateChildItem(lFlags: LONG;
                             bstrItemName: BSTR;
                             bstrFullItemName: BSTR;
                             out ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function EnumRegisterEventInfo(lFlags: LONG;
                                   const pEventGUID: PGUID;
                                   out ppIEnum: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function FindItemByName(lFlags: LONG;
                            bstrFullItemName: BSTR;
                            out ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function DeviceDlg(hwndParent: HWND;
                       lFlags,
                       lIntent: LONG;
                       out plItemCount: LONG;
                       out ppIWiaItem: PIWiaItem): HRESULT; stdcall;

    function DeviceCommand(lFlags: LONG;
                           const pCmdGUID: PGUID;
                           var pIWiaItem: IWiaItem): HRESULT; stdcall;

    function GetRootItem(out ppIWiaItem: IWiaItem): HRESULT; stdcall;

    function EnumDeviceCapabilities(lFlags: LONG;
                                    out ppIEnumWIA_DEV_CAPS: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function DumpItemData(out bstrData: BSTR): HRESULT; stdcall;

    function DumpDrvItemData(out bstrData: BSTR): HRESULT; stdcall;

    function DumpTreeItemData(out bstrData: BSTR): HRESULT; stdcall;

    function Diagnostic(ulSize: ULONG;
                        pBuffer: PBYTE): HRESULT; stdcall;
  end;

  IWiaPropertyStorage = interface(IUnknown)
  ['{98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6}']
    function ReadMultiple(cpspec: ULONG;
                          const rgpspec: PPROPSPEC;
                          rgpropvar: PPROPVARIANT): HRESULT; stdcall; //out

    function WriteMultiple(cpspec: ULONG;
                           const rgpspec: PPROPSPEC;
                           const rgpropvar: PPROPVARIANT;
                           propidNameFirst: PROPID): HRESULT; stdcall;

    function DeleteMultiple(cpspec: ULONG;
                            const rgpspec: PPROPSPEC): HRESULT; stdcall;

    function ReadPropertyNames(cpspec: ULONG;
                               const rgpropid: PPROPID;
                               rglpwstrName: PLPOLESTR): HRESULT; stdcall; //out

    function WritePropertyNames(cpropid: ULONG;
                                const rgpropid: PPROPID;
                                const rglpwstrName: PLPOLESTR): HRESULT;stdcall;

    function DeletePropertyNames(cpropid: ULONG;
                                 const rgpropid: PPROPID): HRESULT; stdcall;

    function Commit(grfCommitFlags: DWORD): HRESULT; stdcall;

    function Revert: HRESULT; stdcall;

    function Enum(out ppenum: IEnumSTATPROPSTG): HRESULT; stdcall;

    function SetTimes(const pctime: PFILETIME;
                      const patime: PFILETIME;
                      const pmtime: PFILETIME): HRESULT; stdcall;

    function SetClass(clsid: PCLSID): HRESULT; stdcall;

    function Stat(pstatpsstg: PSTATPROPSETSTG): HRESULT; stdcall; //out

    function GetPropertyAttributes(cpspec: ULONG;
                                   rgpspec: PPROPSPEC;
                                   rgflags: PULONG;                           //out
                                   rgpropvar: PPROPVARIANT): HRESULT; stdcall;//out

    function GetCount(out pulNumProps: ULONG): HRESULT; stdcall;

    function GetPropertyStream(out pCompatibilityId: TGUID;
                               out ppIStream: IStream): HRESULT; stdcall;

    function SetPropertyStream(pCompatibilityId: PGUID;
                               ppIStream: IStream): HRESULT; stdcall;
  end;

  IEnumWiaItem = interface(IUnknown)
  ['{5e8383fc-3391-11d2-9a33-00c04fa36145}']
    function Next(celt: ULONG;
                  out ppIWiaItem: IWiaItem;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(out ppIEnum: IEnumWiaItem): HRESULT; stdcall;

    function GetCount(out celt: ULONG): HRESULT; stdcall;
  end;

  _WIA_DEV_CAP = record
        guid : GUID;
        ulFlags : ULONG;
        bstrName : BSTR;
        bstrDescription : BSTR;
        bstrIcon : BSTR;
        bstrCommandline : BSTR;
      end;
  WIA_DEV_CAP = _WIA_DEV_CAP;
  PWIA_DEV_CAP = ^WIA_DEV_CAP;

  IEnumWIA_DEV_CAPS = interface(IUnknown)
  ['{1fcc4287-aca6-11d2-a093-00c04f72dc3c}']
    function Next(celt: ULONG;
                  out rgelt: WIA_DEV_CAP;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(out ppIEnum: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function GetCount(out celt: ULONG): HRESULT; stdcall;
  end;

  IEnumWIA_FORMAT_INFO = interface(IUnknown)
  ['{81BEFC5B-656D-44f1-B24C-D41D51B4DC81}']
    function Next(celt: ULONG;
                  out rgelt: WIA_FORMAT_INFO;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(out ppIEnum: IEnumWIA_FORMAT_INFO): HRESULT; stdcall;

    function GetCount(out celt: ULONG): HRESULT; stdcall;
  end;

  IWiaLog = interface(IUnknown)
  ['{A00C10B6-82A1-452f-8B6C-86062AAD6890}']
    function InitializeLog(hInstance: LONG): HRESULT; stdcall;

    function hResult(hResult: HRESULT): HRESULT; stdcall;

    function Log(lFlags,
                 lResID,
                 lDetail: LONG;
                 bstrText: BSTR): HRESULT; stdcall;
  end;

  IWiaLogEx = interface(IUnknown)
  ['{AF1F22AC-7A40-4787-B421-AEb47A1FBD0B}']
    function InitializeLogEx(hInstance: PBYTE): HRESULT; stdcall;

    function hResult(hResult: HRESULT): HRESULT; stdcall;

    function Log(lFlags,
                 lResID,
                 lDetail: LONG;
                 bstrText: BSTR): HRESULT; stdcall;

    function hResultEx(lMethodId: LONG;
                       hResult: HRESULT): HRESULT; stdcall;

    function LogEx(lMethodId: LONG;
                   lFlags,
                   lResID,
                   lDetail: LONG;
                   bstrText: BSTR): HRESULT; stdcall;
  end;

  IWiaNotifyDevMgr = interface(IUnknown)
  ['{70681EA0-E7BF-4291-9FB1-4E8813A3F78E}']
    function NewDeviceArrival(): HRESULT; stdcall;
  end;

  IWiaItemExtras = interface(IUnknown)
  ['{6291ef2c-36ef-4532-876a-8e132593778d}']
    function GetExtendedErrorInfo(out bstrErrorText: BSTR): HRESULT; stdcall;

    function Escape(dwEscapeCode: DWORD;
                    lpInData: PBYTE;
                    cbInDataSize: DWORD;
                    out pOutData: PBYTE;
                    dwOutDataSize: DWORD;
                    out pdwActualDataSize: DWORD): HRESULT; stdcall;

    function CancelPendingIO(): HRESULT; stdcall;
  end;

  IWiaAppErrorHandler = interface(IUnknown)
  ['{6C16186C-D0A6-400c-80F4-D26986A0E734}']
    function GetWindow(out phwnd: HWND): HRESULT; stdcall;

    function ReportStatus(lFlags: LONG;
                          pWiaItem2: IWiaItem2;
                          hrStatus: HRESULT;
                          lPercentComplete: LONG): HRESULT; stdcall;
  end;

  IWiaErrorHandler = interface(IUnknown)
  ['{0e4a51b1-bc1f-443d-a835-72e890759ef3}']
    function ReportStatus(lFlags: LONG;
                          hwndParent: HWND;
                          pWiaItem2: IWiaItem2;
                          hrStatus: HRESULT;
                          lPercentComplete: LONG): HRESULT; stdcall;

    function GetStatusDescription(lFlags: LONG;
                          pWiaItem2: IWiaItem2;
                          hrStatus: HRESULT;
                          out pbstrDescription: BSTR): HRESULT; stdcall;
  end;

  IWiaTransfer = interface(IUnknown)
  ['{c39d6942-2f4e-4d04-92fe-4ef4d3a1de5a}']
    function Download(lFlags: LONG;
                      pIWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;

    function Upload(lFlags: LONG;
                    pSource: IStream;
                    pIWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;

    function Cancel(): HRESULT; stdcall;

    function EnumWIA_FORMAT_INFO(out ppEnum: IEnumWIA_FORMAT_INFO): HRESULT; stdcall;
  end;

  _WiaTransferParams = record
      lMessage : LONG;
      lPercentComplete : LONG;
      ulTransferredBytes : ULONG64;
      hrErrorStatus : HRESULT;
    end;
  WiaTransferParams = _WiaTransferParams;
  PWiaTransferParams = ^WiaTransferParams;

  IWiaTransferCallback = interface(IUnknown)
  ['{27d4eaaf-28a6-4ca5-9aab-e678168b9527}']
    function TransferCallback(lFlags: LONG;
                              pWiaTransferParams: PWiaTransferParams): HRESULT; stdcall;

    function GetNextStream(lFlags: LONG;
                           bstrItemName,
                           bstrFullItemName: BSTR;
                           out ppDestination: IStream): HRESULT; stdcall;
  end;

  IWiaSegmentationFilter = interface(IUnknown)
  ['{EC46A697-AC04-4447-8F65-FF63D5154B21}']
    function DetectRegions(lFlags: LONG;
                           pInputStream: IStream;
                           pWiaItem2: IWiaItem2): HRESULT; stdcall;
  end;

  IWiaImageFilter = interface(IUnknown)
  ['{A8A79FFA-450B-41f1-8F87-849CCD94EBF6}']
    function InitializeFilter(pWiaItem2: IWiaItem2;
                              pWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;

    function SetNewCallback(pWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;


    function FilterPreviewImage(lFlags: LONG;
                                pWiaChildItem2: IWiaItem2;
                                InputImageExtents: RECT;
                                pInputStream: IStream): HRESULT; stdcall;

    function ApplyProperties(pWiaPropertyStorage: IWiaPropertyStorage): HRESULT; stdcall;
  end;

  IWiaPreview = interface(IUnknown)
  ['{95C2B4FD-33F2-4d86-AD40-9431F0DF08F7}']
    function GetNewPreview(lFlags: LONG;
                           pWiaItem2: IWiaItem2;
                           pWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;

    function UpdatePreview(lFlags: LONG;
                           pChildWiaItem2: IWiaItem2;
                           pWiaTransferCallback: IWiaTransferCallback): HRESULT; stdcall;

    function DetectRegions(lFlags: LONG): HRESULT; stdcall;

    function Clear(): HRESULT; stdcall;
  end;

  IEnumWiaItem2 = interface(IUnknown)
  ['{59970AF4-CD0D-44d9-AB24-52295630E582}']
    function Next(celt: ULONG;
                  out ppIWiaItem2: IWiaItem2;
                  var pceltFetched: ULONG): HRESULT; stdcall;

    function Skip(celt: ULONG): HRESULT; stdcall;

    function Reset(): HRESULT; stdcall;

    function Clone(out ppIEnum: IEnumWiaItem2): HRESULT; stdcall;

    function GetCount(out celt: ULONG): HRESULT; stdcall;
  end;

  IWiaItem2 = interface(IUnknown)
  ['{6CBA0075-1287-407d-9B77-CF0E030435CC}']
    function CreateChildItem(lItemFlags,
                             lCreationFlags: LONG;
                             bstrItemName: BSTR;
                             out ppIWiaItem2: IWiaItem2): HRESULT; stdcall;

    function DeleteItem(lFlags: LONG): HRESULT; stdcall;

    function EnumChildItems(const pCategoryGUID: PGUID;
                            out ppIEnumWiaItem2: IEnumWiaItem2): HRESULT; stdcall;

    function FindItemByName(lFlags: LONG;
                            bstrFullItemName: BSTR;
                            out ppIWiaItem2: IWiaItem2): HRESULT; stdcall;

    function GetItemCategory(out pItemCategoryGUID: TGUID): HRESULT; stdcall;

    function GetItemType(out pItemType: LONG): HRESULT; stdcall;

    function DeviceDlg(lFlags: LONG;
                       hwndParent: HWND;
                       bstrFolderName,
                       bstrFilename: BSTR;
                       out plNumFiles: LONG;
                       out ppbstrFilePaths: PBSTR;
                       var ppItem: IWiaItem2): HRESULT; stdcall;

    function DeviceCommand(lFlags: LONG;
                           const pCmdGUID: PGUID;
                           out pIWiaItem2: IWiaItem2): HRESULT; stdcall;

    function EnumDeviceCapabilities(lFlags: LONG;
                                    out ppIEnumWIA_DEV_CAPS: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function CheckExtension(lFlags: LONG;
                            bstrName: BSTR;
                            riidExtensionInterface: REFIID;
                            out pbExtensionExists: BOOL): HRESULT; stdcall;

    function GetExtension(lFlags: LONG;
                          bstrName: BSTR;
                          riidExtensionInterface: REFIID;
                          out ppOut: Pointer): HRESULT; stdcall;

    function GetParentItem(out ppIWiaItem2: IWiaItem2): HRESULT; stdcall;

    function GetRootItem(out ppIWiaItem2: IWiaItem2): HRESULT; stdcall;

    function GetPreviewComponent(lFlags: LONG;
                                 out ppWiaPreview: IWiaPreview): HRESULT; stdcall;

    function EnumRegisterEventInfo(lFlags: LONG;
                                   const pEventGUID: PGUID;
                                   out ppIEnum: IEnumWIA_DEV_CAPS): HRESULT; stdcall;

    function Diagnostic(ulSize: ULONG;
                        pBuffer: PBYTE): HRESULT; stdcall;
  end;

  IWiaDevMgr2 = interface(IUnknown)
  ['{79C07CF1-CBDD-41ee-8EC3-F00080CADA7A}']
    function EnumDeviceInfo(lFlags: LONG;
                            out ppIEnum: IEnumWIA_DEV_INFO): HRESULT; stdcall;

    function CreateDevice(lFlags: LONG;
                          bstrDeviceID: BSTR;
                          out ppWiaItem2Root: IWiaItem2): HRESULT; stdcall;

    function SelectDeviceDlg(hwndParent: HWND;
                             lDeviceType,
                             lFlags: LONG;
                             var pbstrDeviceID: BSTR;
                             out ppItemRoot: IWiaItem2): HRESULT; stdcall;

    function SelectDeviceDlgID(hwndParent: HWND;
                               lDeviceType,
                               lFlags: LONG;
                               out pbstrDeviceID: BSTR): HRESULT; stdcall;

    function RegisterEventCallbackInterface(lFlags: LONG;
                                            bstrDeviceID: BSTR;
                                            const pEventGUID: PGUID;
                                            pIWiaEventCallback: IWiaEventCallback;
                                            out pEventObject: IUnknown): HRESULT; stdcall;

    function RegisterEventCallbackProgram(lFlags: LONG;
                                          bstrDeviceID: BSTR;
                                          const pEventGUID: PGUID;
                                          bstrFullAppName,
                                          bstrCommandLineArg,
                                          bstrName,
                                          bstrDescription,
                                          bstrIcon: BSTR): HRESULT; stdcall;

    function RegisterEventCallbackCLSID(lFlags: LONG;
                                        bstrDeviceID: BSTR;
                                        const pEventGUID: PGUID;
                                        const pClsID: PGUID;
                                        bstrName,
                                        bstrDescription,
                                        bstrIcon: BSTR): HRESULT; stdcall;

    function GetImageDlg(lFlags: LONG;
                         bstrDeviceID: BSTR;
                         hwndParent: HWND;
                         bstrFolderName,
                         bstrFilename: BSTR;
                         out plNumFiles: LONG;
                         out ppbstrFilePaths: PBSTR;
                         var ppItem: IWiaItem2): HRESULT; stdcall;
  end;

implementation

end.

