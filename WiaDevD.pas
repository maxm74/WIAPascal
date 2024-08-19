(****************************************************************************
*
*
*  FILE: Translation of wiadevd.h
*
*  VERSION:     1.0
*
*  DESCRIPTION:
*    Device Dialog and UI extensibility declarations.
*
*****************************************************************************
*
*  2024-04-05 Translation and adaptation by Massimo Magnano
*
*****************************************************************************)

unit WiaDevD;

{$H+}
{$ALIGN 8}

interface

uses Windows, ActiveX, WIA_LH;

const
  IID_IWiaUIExtension2 : TGUID ='{305600d7-5088-46d7-9a15-b77b09cdba7a}';
  IID_IWiaUIExtension  : TGUID ='{da319113-50ee-4c80-b460-57d005d44a2c}';

type
  PtagDEVICEDIALOGDATA2 = ^tagDEVICEDIALOGDATA2;
  tagDEVICEDIALOGDATA2 = record
    cbSize : DWORD;              // Size of the structure in bytes
    pIWiaItemRoot : IWiaItem2;   // Valid root item
    dwFlags : DWORD;             // Flags
    hwndParent : HWND;           // Parent window
    bstrFolderName : BSTR;       // Folder name where the files are transferred
    bstrFilename : BSTR;         // template file name.
    lNumFiles : LONG;            // Number of items in ppbstrFilePaths array.  Filled on return.
    pbstrFilePaths : PBSTR;      // file names created after successful transfers.
    pWiaItem : IWiaItem2;        // IWiaItem2 interface pointer.  This is the IWiaItem2 used for transfer.
  end;
  DEVICEDIALOGDATA2 = tagDEVICEDIALOGDATA2;
  LPDEVICEDIALOGDATA2 = PtagDEVICEDIALOGDATA2;
  PDEVICEDIALOGDATA2 = PtagDEVICEDIALOGDATA2;

  IWiaUIExtension2 = interface(IUnknown)
    ['{305600d7-5088-46d7-9a15-b77b09cdba7a}']
    procedure DeviceDialog(const pDeviceDialogData: PDEVICEDIALOGDATA2); stdcall;
    procedure GetDeviceIcon(const bstrDeviceId: BSTR; var phIcon: HICON; nSize: ULONG); stdcall;
  end;

  PtagDEVICEDIALOGDATA = ^tagDEVICEDIALOGDATA;
  tagDEVICEDIALOGDATA = record
    cbSize : DWORD;              // Size of the structure in bytes
    hwndParent : HWND;           // Parent window
    pIWiaItemRoot : IWiaItem;    // Valid root item
    dwFlags : DWORD;             // Flags
    lIntent : LONG;              // Intent flags
    lItemCount : LONG;           // Number of items in ppWiaItems array.  Filled on return.
    ppWiaItems : PIWiaItem;      // Array of IWiaItem interface pointers.  Array must
                                 // be allocated using CoTaskMemAlloc, and all interface pointers must be AddRef'ed
  end;
  DEVICEDIALOGDATA = tagDEVICEDIALOGDATA;
  LPDEVICEDIALOGDATA = PtagDEVICEDIALOGDATA;
  PDEVICEDIALOGDATA = PtagDEVICEDIALOGDATA;

  // IWiaUIExtension provides a means to replace a device's image acquisition dialog
  // and to provide custom icons and logo bitmaps to appear on the standard dialog
  IWiaUIExtension = interface(IUnknown)
    ['{da319113-50ee-4c80-b460-57d005d44a2c}']
    procedure DeviceDialog(const pDeviceDialogData: PDEVICEDIALOGDATA); stdcall;
    procedure GetDeviceIcon(const bstrDeviceId: BSTR; var phIcon: HICON; nSize: ULONG); stdcall;
    procedure GetDeviceBitmapLogo(const bstrDeviceId: BSTR; var phBitmap: HBITMAP; nMaxWidth, nMaxHeight: ULONG); stdcall;
  end;

implementation

end.
