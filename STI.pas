(****************************************************************************
*
* Module Name:
*
*    STI.pas
*
*  Abstract:
*
*    This module contains the user mode Still Image APIs in COM format
*****************************************************************************
*
*  2024-04-19 Translation and adaptation by Massimo Magnano
*
*****************************************************************************)

unit STI;

{$ALIGN 8}
{$H+}
{$POINTERMATH ON}

interface

uses Windows;


const
  //
  // Only use UNICODE STI interfaces
  //
  STI_UNICODE = 1;

  //
  // Class IID's
  //
  CLSID_Sti : TGUID = '{B323F8E0-2E68-11D0-90EA-00AA0060F86C}';

  //
  // Interface IID's
  //
  IID_IStillImageW : TGUID = '{641BD880-2DC8-11D0-90EA-00AA0060F86C}';
  //{$ifdef UNICODE}
  IID_IStillImage : TGUID = '{641BD880-2DC8-11D0-90EA-00AA0060F86C}';
  //{$endif}
  IID_IStiDevice : TGUID = '{6CFA5A80-2DC8-11D0-90EA-00AA0060F86C}';

  //
  // Standard event GUIDs
  //
  GUID_DeviceArrivedLaunch : TGUID = '{740D9EE6-70F1-11d1-AD10-00A02438AD48}';
  GUID_ScanImage : TGUID = '{A6C5A715-8C6E-11d2-977A-0000F87A926F}';
  GUID_ScanPrintImage : TGUID = '{B441F425-8C6E-11d2-977A-0000F87A926F}';
  GUID_ScanFaxImage : TGUID = '{C00EB793-8C6E-11d2-977A-0000F87A926F}';
  GUID_STIUserDefined1 : TGUID = '{C00EB795-8C6E-11d2-977A-0000F87A926F}';
  GUID_STIUserDefined2 : TGUID = '{C77AE9C5-8C6E-11d2-977A-0000F87A926F}';
  GUID_STIUserDefined3 : TGUID = '{C77AE9C6-8C6E-11d2-977A-0000F87A926F}';

  //
  // Generic constants and definitions
  //
  STI_VERSION_FLAG_MASK = $ff000000;
  STI_VERSION_FLAG_UNICODE = $01000000;

  STI_VERSION_REAL = $00000002;
  //{$if (_WIN32_WINNT >= 0x0600) // Windows Vista and later}
  STI_VERSION_3 = $00000003 or STI_VERSION_FLAG_UNICODE;
  //{$endif}

  STI_VERSION_MIN_ALLOWED = $00000002;

  {$ifdef UNICODE}
  STI_VERSION = STI_VERSION_REAL or STI_VERSION_FLAG_UNICODE;
  {$else}
  STI_VERSION = STI_VERSION_REAL;
  {$endif}

  //
  // Maximum length of internal device name
  //
  STI_MAX_INTERNAL_NAME_LENGTH = 128;

type
{$IFNDEF FPC}
{$IFDEF CPUX64}
  PtrInt = Int64;
  PtrUInt = UInt64;
{$ELSE}
  PtrInt = longint;
  PtrUInt = Longword;
{$ENDIF}
{$ENDIF}
  // begin sti_device_information

  //
  //  Device information definitions and prototypes
  // ----------------------------------------------
  //

  //
  //  Following information is used for enumerating still image devices , currently configured
  //  in the system. Presence of the device in the enumerated list does not mean availability
  // of the device, it only means that device was installed at least once and had not been removed since.
  //

  //
  // Type of device ( scanner, camera) is represented by DWORD value with
  // hi word containing generic device type , and lo word containing sub-type
  //
  _STI_DEVICE_MJ_TYPE = (
       StiDeviceTypeDefault          = 0,
       StiDeviceTypeScanner          = 1,
       StiDeviceTypeDigitalCamera    = 2,
       //#if (_WIN32_WINNT >= 0x0501) // Windows XP and later
       StiDeviceTypeStreamingVideo   = 3
       //#endif //#if (_WIN32_WINNT >= 0x0501)
  );
  STI_DEVICE_MJ_TYPE = _STI_DEVICE_MJ_TYPE;

  STI_DEVICE_TYPE = DWORD;

  //
  // Device capabilities bits.
  // Various capabilities are grouped into separate bitmasks
  //
  _STI_DEV_CAPS = record
       dwGeneric : DWORD;
  end;
  STI_DEV_CAPS = _STI_DEV_CAPS;
  PSTI_DEV_CAPS = ^_STI_DEV_CAPS;

const
  STI_GENCAP_COMMON_MASK : DWORD = $00ff;

  //
  // Notifications are supported.
  // If this capability set , device can be subscribed to .
  //
  STI_GENCAP_NOTIFICATIONS = $00000001;

  //
  // Polling required .
  // This capability is used when previous is set to TRUE. Presence of it means
  // that device is not capable of issuing "truly" asyncronous notifications, but can
  // be polled to determine the moment when event happened
  STI_GENCAP_POLLING_NEEDED = $00000002;

  //
  // Generate event on device arrival
  // If this capability is set, still image service will generate event when device
  // instance is successfully initialized ( typically in response to PnP arrival)

  //
  // Note: on initial service enumeration events will nto be generated to avoid
  // end-user confusion.
  //
  STI_GENCAP_GENERATE_ARRIVALEVENT = $00000004;

  //
  // Auto port selection on non-PnP buses
  // This capability indicates that USD is able to detect non-PnP device on a
  // bus , device is supposed to be attached to.
  //
  STI_GENCAP_AUTO_PORTSELECT = $00000008;

  //#if (_WIN32_WINNT >= 0x0501) // Windows XP and later

  //
  // WIA capability bit.
  // This capability indicates that USD is WIA capable.
  //
  STI_GENCAP_WIA = $00000010;

  //
  // Subset driver bit.
  // This capability indicates that there is more featured driver exists. All
  // of inbox driver has this bit set. Fully featured (IHV) driver shouldn't have
  // this bit set.
  //
  STI_GENCAP_SUBSET = $00000020;

  //#endif //#if (_WIN32_WINNT >= 0x0501)
  //#if (_WIN32_WINNT >= 0x0600) // Windows Vista and later
  WIA_INCOMPAT_XP = $00000001;
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  //
  // Type of bus connection for those in need to know
  //
  STI_HW_CONFIG_UNKNOWN = $0001;
  STI_HW_CONFIG_SCSI = $0002;
  STI_HW_CONFIG_USB = $0004;
  STI_HW_CONFIG_SERIAL = $0008;
  STI_HW_CONFIG_PARALLEL = $0010;

type
  //
  // Device information structure, this is not configurable. This data is returned from
  // device enumeration API and is used for populating UI or selecting which device
  // should be used in current session
  //
  P_STI_DEVICE_INFORMATIONW = ^_STI_DEVICE_INFORMATIONW;
  _STI_DEVICE_INFORMATIONW = record
       dwSize : DWORD;

       // Type of the hardware imaging device
       DeviceType : STI_DEVICE_TYPE;

       // Device identifier for reference when creating device object
       szDeviceInternalName : array[0..(STI_MAX_INTERNAL_NAME_LENGTH)-1] of WCHAR;

       // Set of capabilities flags
       DeviceCapabilities : STI_DEV_CAPS;

       // This includes bus type
       dwHardwareConfiguration : DWORD;

       // Vendor description string
       pszVendorDescription : LPWSTR;

       // Device description , provided by vendor
       pszDeviceDescription : LPWSTR;

       // String , representing port on which device is accessible.
       pszPortName : LPWSTR;

       // Control panel propery provider
       pszPropProvider : LPWSTR;

       // Local specific ("friendly") name of the device, mainly used for showing in the UI
       pszLocalName : LPWSTR;
  end;
  STI_DEVICE_INFORMATIONW = _STI_DEVICE_INFORMATIONW;
  PSTI_DEVICE_INFORMATIONW = ^STI_DEVICE_INFORMATIONW;
  //{$ifdef UNICODE}
   STI_DEVICE_INFORMATION = STI_DEVICE_INFORMATIONW;
   PSTI_DEVICE_INFORMATION = PSTI_DEVICE_INFORMATIONW;
  //{$endif}

  //
  // EXTENDED STI INFORMATION TO COVER WIA
  //
  _STI_WIA_DEVICE_INFORMATIONW = record
       dwSize : DWORD;

       // Type of the hardware imaging device
       DeviceType : STI_DEVICE_TYPE;

       // Device identifier for reference when creating device object
       szDeviceInternalName : array[0..(STI_MAX_INTERNAL_NAME_LENGTH)-1] of WCHAR;

       // Set of capabilities flags
       DeviceCapabilities : STI_DEV_CAPS;

       // This includes bus type
       dwHardwareConfiguration : DWORD;

       // Vendor description string
       pszVendorDescription : LPWSTR;

       // Device description , provided by vendor
       pszDeviceDescription : LPWSTR;

       // String , representing port on which device is accessible.
       pszPortName : LPWSTR;

       // Control panel propery provider
       pszPropProvider : LPWSTR;

       // Local specific ("friendly") name of the device, mainly used for showing in the UI
       pszLocalName : LPWSTR;

       //
       // WIA values
       //

       pszUiDll : LPWSTR;
       pszServer : LPWSTR;
  end;
  STI_WIA_DEVICE_INFORMATIONW = _STI_WIA_DEVICE_INFORMATIONW;
  PSTI_WIA_DEVICE_INFORMATIONW = ^STI_WIA_DEVICE_INFORMATIONW;
  //{$ifdef UNICODE}
  STI_WIA_DEVICE_INFORMATION = STI_WIA_DEVICE_INFORMATIONW;
  PSTI_WIA_DEVICE_INFORMATION = PSTI_WIA_DEVICE_INFORMATIONW;
  // {$endif}

// end sti_device_information

const
  //
  // Device state information.
  // ------------------------
  //
  // Following types  are used to inquire state characteristics of the device after
  // it had been opened.
  //
  // Device configuration structure contains configurable parameters reflecting
  // current state of the device
  //
  //
  // Device hardware status.
  //

  //
  // Individual bits for state acquiring  through StatusMask
  //

  // State of hardware as known to USD
  STI_DEVSTATUS_ONLINE_STATE = $0001;

  // State of pending events ( as known to USD)
  STI_DEVSTATUS_EVENTS_STATE = $0002;

  //
  // Online state values
  //
  STI_ONLINESTATE_OPERATIONAL = $00000001;
  STI_ONLINESTATE_PENDING = $00000002;
  STI_ONLINESTATE_ERROR = $00000004;
  STI_ONLINESTATE_PAUSED = $00000008;
  STI_ONLINESTATE_PAPER_JAM = $00000010;
  STI_ONLINESTATE_PAPER_PROBLEM = $00000020;
  STI_ONLINESTATE_OFFLINE = $00000040;
  STI_ONLINESTATE_IO_ACTIVE = $00000080;
  STI_ONLINESTATE_BUSY = $00000100;
  STI_ONLINESTATE_TRANSFERRING = $00000200;
  STI_ONLINESTATE_INITIALIZING = $00000400;
  STI_ONLINESTATE_WARMING_UP = $00000800;
  STI_ONLINESTATE_USER_INTERVENTION = $00001000;
  STI_ONLINESTATE_POWER_SAVE = $00002000;

  //
  // Event processing parameters
  //
  STI_EVENTHANDLING_ENABLED = $00000001;
  STI_EVENTHANDLING_POLLING = $00000002;
  STI_EVENTHANDLING_PENDING = $00000004;

type
  _STI_DEVICE_STATUS = record
      dwSize : DWORD;

      // Request field - bits of status to verify
      StatusMask : DWORD;

      //
      // Fields are set when status mask contains STI_DEVSTATUS_ONLINE_STATE bit set
      //
      // Bitmask describing  device state
      dwOnlineState : DWORD;

      // Device status code as defined by vendor
      dwHardwareStatusCode : DWORD;

      //
      // Fields are set when status mask contains STI_DEVSTATUS_EVENTS_STATE bit set
      //

      // State of device notification processing (enabled, pending)
      dwEventHandlingState : DWORD;

      // If device is polled, polling interval in ms
      dwPollingInterval : DWORD;
    end;
  STI_DEVICE_STATUS = _STI_DEVICE_STATUS;
  PSTI_DEVICE_STATUS = ^STI_DEVICE_STATUS;

//
// Structure to describe diagnostic ( test ) request to be processed by USD
//

const
  // Basic test for presence of associated hardware
  STI_DIAGCODE_HWPRESENCE = $00000001;

type
  //
  // Status bits for diagnostic
  //

  //
  // generic diagnostic errors
  //
  _ERROR_INFOW = record
      dwSize : DWORD;

      // Generic error , describing results of last operation
      dwGenericError : DWORD;

      // vendor specific error code
      dwVendorError : DWORD;

      // String, describing in more details results of last operation if it failed
      szExtendedErrorText : array[0..254] of WCHAR;
  end;
  STI_ERROR_INFOW = _ERROR_INFOW;
  PSTI_ERROR_INFOW = ^STI_ERROR_INFOW;
  //{$ifdef UNICODE}
  STI_ERROR_INFO = STI_ERROR_INFOW;
  PSTI_ERROR_INFO = PSTI_ERROR_INFOW;
  //{$endif}

  _STI_DIAG = record
      dwSize : DWORD;

      // Diagnostic request fields. Are set on request by caller

      // One of the
      dwBasicDiagCode : DWORD;
      dwVendorDiagCode : DWORD;

      // Response fields
      dwStatusMask : DWORD;

      sErrorInfo : STI_ERROR_INFO;
    end;
  STI_DIAG = _STI_DIAG;
  LPSTI_DIAG = ^STI_DIAG;
  PSTI_DIAG = ^STI_DIAG;
  DIAG = STI_DIAG;
  PDIAG = LPSTI_DIAG;
  LPDIAG = LPSTI_DIAG;

// end device state information.

const
  //
  // Flags passed to WriteToErrorLog call in a first parameter, indicating type of the message
  // which needs to be logged
  //
  STI_TRACE_INFORMATION = $00000001;
  STI_TRACE_WARNING = $00000002;
  STI_TRACE_ERROR = $00000004;

//
// Event notification mechansims.
// ------------------------------
//
// Those are used to inform last subscribed caller of the changes in device state, initiated by
// device.
//
// The only supported discipline of notification is stack. The last caller to subscribe will be notified
// and will receive notification data. After caller unsubscribes , the previously subscribed caller will
// become active.
//

  // Notifications are sent to subscriber via window message. Window handle is passed as
  // parameter
  STI_SUBSCRIBE_FLAG_WINDOW = $0001;

  // Device notification is signalling Win32 event ( auto-set event). Event handle
  // is passed as a parameter
  STI_SUBSCRIBE_FLAG_EVENT = $0002;

type
  P_STISUBSCRIBE = ^_STISUBSCRIBE;
  _STISUBSCRIBE = record
      dwSize : DWORD;
      dwFlags : DWORD;

      // Not used . Will be used for subscriber to set bit mask filtering different events
      dwFilter : DWORD;

      // When STI_SUBSCRIBE_FLAG_WINDOW bit is set, following fields should be set
      // Handle of the window which will receive notification message
      hWndNotify : HWND;

      // Handle of Win32 auto-reset event , which will be signalled whenever device has
      // notification pending
      hEvent : THANDLE;

      // Code of notification message, sent to window
      uiNotificationMessage : UINT;
    end;
  STISUBSCRIBE = _STISUBSCRIBE;
  PSTISUBSCRIBE = ^STISUBSCRIBE;
  LPSTISUBSCRIBE = PSTISUBSCRIBE;

const
  MAX_NOTIFICATION_DATA = 64;

type
  //
  // Structure to describe notification information
  //
  _STINOTIFY = record
      // Total size of the notification structure
      dwSize : DWORD;

      // GUID of the notification being retrieved
      guidNotificationCode : TGUID;

      // Vendor specific notification description
      abNotificationData : array[0..(MAX_NOTIFICATION_DATA)-1] of BYTE;
    end;
  STINOTIFY = _STINOTIFY;
  PSTINOTIFY = ^STINOTIFY;
  LPSTINOTIFY = PSTINOTIFY;

// end event_mechanisms

//
// STI device broadcasting
//

//
// When STI Device is being added or removed, PnP broadacst is being sent , but it is not obvious
// for application code to recognize if it is STI device and if so, what is the name of the
// device. STI subsystem will analyze PnP broadcasts and rebroadcast another message via
// BroadcastSystemMessage / WM_DEVICECHANGE / DBT_USERDEFINED .

const
  // String passed as user defined message contains STI prefix, action and device name

  STI_ADD_DEVICE_BROADCAST_ACTION = 'Arrival';
  STI_REMOVE_DEVICE_BROADCAST_ACTION = 'Removal';

  STI_ADD_DEVICE_BROADCAST_STRING = 'STI\\'+STI_ADD_DEVICE_BROADCAST_ACTION+'\\%s';
  STI_REMOVE_DEVICE_BROADCAST_STRING = 'STI\\'+STI_REMOVE_DEVICE_BROADCAST_ACTION+'\\%s';

// end STI broadcasting


   //
   // Device create modes
   //

   // Device is being opened only for status querying and notifications receiving
   STI_DEVICE_CREATE_STATUS = $00000001;

   // Device is being opened for data transfer ( supersedes status mode)
   STI_DEVICE_CREATE_DATA = $00000002;
   STI_DEVICE_CREATE_BOTH = $00000003;

   //
   // Bit mask for legitimate mode bits, which can be used when calling CreateDevice
   //
   STI_DEVICE_CREATE_MASK = $0000FFFF;

   //
   // Flags controlling device enumeration
   //
   STIEDFL_ALLDEVICES = $00000000;
   STIEDFL_ATTACHEDONLY = $00000001;

//
// Control code , sent to the device through raw control interface
//
type
  STI_RAW_CONTROL_CODE = DWORD;

const
  //
  // All raw codes below this one are reserved for future use.
  //
  STI_RAW_RESERVED = $1000;

type
  IStillImageW = interface;
  IStiDeviceW = interface;

  //{$ifdef UNICODE}
  IStillImage = IStillImageW;
  LPSTILLIMAGE = IStillImage;
  IStiDevice = IStiDeviceW;
  //{$endif}

  LPSTILLIMAGEDEVICE = IStiDevice;
  PSTI = IStillImage;
  PSTIDEVICE = IStiDevice;
  PSTIW = IStillImageW;
  PSTIDEVICEW = IStiDeviceW;

  //
  // IStillImage interface
  //
  // Top level STI access interface.
  //
  //
  IStillImageW = interface(IUnknown)
  ['{641BD880-2DC8-11D0-90EA-00AA0060F86C}']
    function Initialize(hinst: PtrUInt; dwVersion: DWORD): HRESULT; stdcall;
    function GetDeviceList(dwType, dwFlags: DWORD; out pdwItemsReturned: DWORD; out ppBuffer: LPVOID): HRESULT; stdcall;
    function GetDeviceInfo(pwszDeviceName: LPWSTR; out ppBuffer: LPVOID): HRESULT; stdcall;
    function CreateDevice(pwszDeviceName: LPWSTR; dwMode: DWORD; out pDevice: PSTIDEVICE; punkOuter: IUnknown): HRESULT; stdcall;

    //
    // Device instance values. Used to associate various data with device.
    //
    function GetDeviceValue(pwszDeviceName, pValueName: LPWSTR; out pType: DWORD; out pData: LPBYTE; var cbData: DWORD): HRESULT; stdcall;
    function SetDeviceValue(pwszDeviceName, pValueName: LPWSTR; Type_: DWORD; pData: LPBYTE; cbData: DWORD): HRESULT; stdcall;

    //
    // For appllication started through push model launch, returns associated information
    //
    function GetSTILaunchInformation(out pwszDeviceName: LPWSTR; out pdwEventCode: DWORD; out pwszEventName: LPWSTR): HRESULT; stdcall;
    function RegisterLaunchApplication(pwszAppName, pwszCommandLine: LPWSTR): HRESULT; stdcall;
    function UnregisterLaunchApplication(pwszAppName: LPWSTR): HRESULT; stdcall;

    //
    // To control state of notification handling. For polled devices this means state of monitor
    // polling, for true notification devices means enabling/disabling notification flow
    // from monitor to registered applications
    //
    function EnableHwNotifications(pwszDeviceName: LPCWSTR; bNewState: BOOL): HRESULT; stdcall;
    function GetHwNotificationState(pwszDeviceName: LPCWSTR; out pbCurrentState: BOOL): HRESULT; stdcall;

    //
    // When device is installed but not accessible, application may request bus refresh
    // which in some cases will make device known. This is mainly used for nonPnP buses
    // like SCSI, when device was powered on after PnP enumeration
    //
    //
    function RefreshDeviceBus(pwszDeviceName: LPCWSTR): HRESULT; stdcall;

    //
    // Launch application to emulate event on a device. Used by "control center" style components,
    // which intercept device event , analyze and later force launch based on certain criteria.
    //
    function LaunchApplicationForDevice(pwszDeviceName, pwszAppName: LPWSTR; pStiNotify: LPSTINOTIFY): HRESULT; stdcall;

    //
    // For non-PnP devices with non-known bus type connection, setup extension, associated with the
    // device can set it's parameters
    //
    function SetupDeviceParameters(var unnamedParam1: STI_DEVICE_INFORMATIONW): HRESULT; stdcall;

    //
    // Write message to STI error log
    //
    function WriteToErrorLog(dwMessageType: DWORD; pszMessage: LPCWSTR): HRESULT; stdcall;

    //
    // TO register application for receiving various STI notifications
    //
    function RegisterDeviceNotification(pwszAppName: LPWSTR; var lpSubscribe: STISUBSCRIBE): HRESULT; stdcall;
    function UnregisterDeviceNotification(): HRESULT; stdcall;
  end;

  //
  // IStillImage_Device interface
  //
  // This is generic per device interface. Specialized interfaces are also
  // available
  //
  IStiDevice = interface(IUnknown)
  ['{6CFA5A80-2DC8-11D0-90EA-00AA0060F86C}']
    function Initialize(hinst: PtrUInt; pwszDeviceName: LPCWSTR; dwVersion, dwMode: DWORD): HRESULT; stdcall;
    function GetCapabilities(var pDevCaps: STI_DEV_CAPS): HRESULT; stdcall;
    function GetStatus(var pDevStatus: STI_DEVICE_STATUS): HRESULT; stdcall;
    function DeviceReset(): HRESULT; stdcall;
    function Diagnostic(var pBuffer: STI_DIAG): HRESULT; stdcall;
    function Escape(EscapeFunction: STI_RAW_CONTROL_CODE; lpInData: LPVOID; cbInDataSize: DWORD;
                    var pOutData: LPVOID; dwOutDataSize: DWORD; out pdwActualData: DWORD): HRESULT; stdcall;
    function GetLastError(out pdwLastDeviceError: DWORD): HRESULT; stdcall;
    function LockDevice(dwTimeOut: DWORD): HRESULT; stdcall;
    function UnLockDevice(): HRESULT; stdcall;
    function RawReadData(var lpBuffer: LPVOID; var lpdwNumberOfBytes: DWORD; lpOverlapped: POVERLAPPED): HRESULT; stdcall;
    function RawWriteData(lpBuffer: LPVOID; nNumberOfBytes: DWORD; lpOverlapped: POVERLAPPED): HRESULT; stdcall;
    function RawReadCommand(var lpBuffer: LPVOID; var lpdwNumberOfBytes: DWORD; lpOverlapped: POVERLAPPED): HRESULT; stdcall;
    function RawWriteCommand(lpBuffer: LPVOID; nNumberOfBytes: DWORD; lpOverlapped: POVERLAPPED): HRESULT; stdcall;

    //
    // Subscription is used to enable "control center" style applications , where flow of
    // notifications should be redirected from monitor itself to another "launcher"
    //
    function Subscribe(var lpSubsribe: STISUBSCRIBE): HRESULT; stdcall;
    function GetLastNotificationData(out lpNotify: STINOTIFY): HRESULT; stdcall;
    function UnSubscribe(): HRESULT; stdcall;

    function GetLastErrorInfo(out pLastErrorInfo: STI_ERROR_INFO): HRESULT; stdcall;
  end;


function GET_STIVER_MAJOR(dwVersion : longint) : DWORD;
function GET_STIVER_MINOR(dwVersion : longint) : WORD;

//
// Macros to extract device type/subtype from single type field
//
function GET_STIDEVICE_TYPE(dwDevType : longint) : WORD;
function GET_STIDEVICE_SUBTYPE(dwDevType : longint) : WORD;

//
// Generic capabilities mask contain 16 bits , common for all devices, maintained by MS
// and 16 bits , which USD can use for proprietary capbailities reporting.
//
function GET_STIDCOMMON_CAPS(dwGenericCaps : longint) : WORD;
function GET_STIVENDOR_CAPS(dwGenericCaps : longint) : WORD;

function StiCreateInstanceW(hinst: PtrUInt; dwVer: DWORD; out ppSti: IStillImageW; punkOuter: IUnknown): HRESULT; stdcall; external 'Sti.dll';

//{$ifdef UNICODE}
function StiCreateInstance(hinst: PtrUInt; dwVer: DWORD; out ppSti: IStillImageW; punkOuter: IUnknown): HRESULT; stdcall; external 'Sti.dll' name 'StiCreateInstanceW';
//{$endif}

implementation

function GET_STIVER_MAJOR(dwVersion : longint) : DWORD;
begin
  Result:= HIWORD(dwVersion) and not(STI_VERSION_FLAG_MASK);
end;

function GET_STIVER_MINOR(dwVersion : longint) : WORD;
begin
  Result:= LOWORD(dwVersion);
end;

function GET_STIDEVICE_TYPE(dwDevType : longint) : WORD;
begin
  Result:= HIWORD(dwDevType);
end;

function GET_STIDEVICE_SUBTYPE(dwDevType : longint) : WORD;
begin
  Result:= LOWORD(dwDevType);
end;

function GET_STIDCOMMON_CAPS(dwGenericCaps : longint) : WORD;
begin
  Result:= LOWORD(dwGenericCaps);
end;

function GET_STIVENDOR_CAPS(dwGenericCaps : longint) : WORD;
begin
  Result:= HIWORD(dwGenericCaps);
end;

end.
