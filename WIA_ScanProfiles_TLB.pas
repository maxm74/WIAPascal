unit WIA_ScanProfiles_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 98336 $
// File generated on 03/04/2024 12:53:44 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Windows\SysWOW64\wiascanprofiles.dll (1)
// LIBID: {77A6BD8A-AB60-49FF-853C-B6EE7BABAF96}
// LCID: 0
// Helpfile: 
// HelpString: ScanProfiles 1.0 Type Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// Errors:
//   Hint: Symbol 'type' renamed to 'type_'
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}
{$H+}

interface

uses Windows, Classes, Variants, OleServer, ActiveX;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  ScanProfilesLibMajorVersion = 1;
  ScanProfilesLibMinorVersion = 0;

  LIBID_ScanProfilesLib: TGUID = '{77A6BD8A-AB60-49FF-853C-B6EE7BABAF96}';

  IID_IScanProfileMgr: TGUID = '{34EAAE27-2D89-4278-84EF-61DEFA323FC1}';
  CLASS_ScanProfileMgr: TGUID = '{CB0FC8E5-686A-478B-A252-FDECF8E167B7}';
  IID_IScanProfile: TGUID = '{D68A6C07-9FF8-47D1-9A2A-429D28FBC6A4}';
  IID_IStorage: TGUID = '{0000000B-0000-0000-C000-000000000046}';
  IID_IEnumSTATSTG: TGUID = '{0000000D-0000-0000-C000-000000000046}';
  IID_IRecordInfo: TGUID = '{0000002F-0000-0000-C000-000000000046}';
  IID_ITypeInfo: TGUID = '{00020401-0000-0000-C000-000000000046}';
  IID_ITypeComp: TGUID = '{00020403-0000-0000-C000-000000000046}';
  IID_ITypeLib: TGUID = '{00020402-0000-0000-C000-000000000046}';
  IID_IScanProfileUI: TGUID = '{B67CDDB7-2B20-473E-8D6C-3F1BD202E78D}';
  CLASS_ScanProfileUI: TGUID = '{19603261-6059-43DF-B9E1-8B4352825A90}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum tagTYPEKIND
type
  tagTYPEKIND = TOleEnum;
const
  TKIND_ENUM = $00000000;
  TKIND_RECORD = $00000001;
  TKIND_MODULE = $00000002;
  TKIND_INTERFACE = $00000003;
  TKIND_DISPATCH = $00000004;
  TKIND_COCLASS = $00000005;
  TKIND_ALIAS = $00000006;
  TKIND_UNION = $00000007;
  TKIND_MAX = $00000008;

// Constants for enum tagDESCKIND
type
  tagDESCKIND = TOleEnum;
const
  DESCKIND_NONE = $00000000;
  DESCKIND_FUNCDESC = $00000001;
  DESCKIND_VARDESC = $00000002;
  DESCKIND_TYPECOMP = $00000003;
  DESCKIND_IMPLICITAPPOBJ = $00000004;
  DESCKIND_MAX = $00000005;

// Constants for enum tagFUNCKIND
type
  tagFUNCKIND = TOleEnum;
const
  FUNC_VIRTUAL = $00000000;
  FUNC_PUREVIRTUAL = $00000001;
  FUNC_NONVIRTUAL = $00000002;
  FUNC_STATIC = $00000003;
  FUNC_DISPATCH = $00000004;

// Constants for enum tagINVOKEKIND
type
  tagINVOKEKIND = TOleEnum;
const
  INVOKE_FUNC = $00000001;
  INVOKE_PROPERTYGET = $00000002;
  INVOKE_PROPERTYPUT = $00000004;
  INVOKE_PROPERTYPUTREF = $00000008;

// Constants for enum tagCALLCONV
type
  tagCALLCONV = TOleEnum;
const
  CC_FASTCALL = $00000000;
  CC_CDECL = $00000001;
  CC_MSCPASCAL = $00000002;
  CC_PASCAL = $00000002;
  CC_MACPASCAL = $00000003;
  CC_STDCALL = $00000004;
  CC_FPFASTCALL = $00000005;
  CC_SYSCALL = $00000006;
  CC_MPWCDECL = $00000007;
  CC_MPWPASCAL = $00000008;
  CC_MAX = $00000009;

// Constants for enum tagVARKIND
type
  tagVARKIND = TOleEnum;
const
  VAR_PERINSTANCE = $00000000;
  VAR_STATIC = $00000001;
  VAR_CONST = $00000002;
  VAR_DISPATCH = $00000003;

// Constants for enum tagSYSKIND
type
  tagSYSKIND = TOleEnum;
const
  SYS_WIN16 = $00000000;
  SYS_WIN32 = $00000001;
  SYS_MAC = $00000002;
  SYS_WIN64 = $00000003;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IScanProfileMgr = interface;
  IScanProfileMgrDisp = dispinterface;
  IScanProfile = interface;
  IScanProfileDisp = dispinterface;
  IStorage = interface;
  IEnumSTATSTG = interface;
  IRecordInfo = interface;
  ITypeInfo = interface;
  ITypeComp = interface;
  ITypeLib = interface;
  IScanProfileUI = interface;
  IScanProfileUIDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  ScanProfileMgr = IScanProfileMgr;
  ScanProfileUI = IScanProfileUI;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  wirePSAFEARRAY = ^PUserType4; 
  wireSNB = ^tagRemSNB; 
  PUserType5 = ^_FLAGGED_WORD_BLOB; {*}
  PUserType6 = ^_wireVARIANT; {*}
  PUserType13 = ^_wireBRECORD; {*}
  PUserType4 = ^_wireSAFEARRAY; {*}
  PPUserType1 = ^PUserType4; {*}
  PUserType10 = ^tagTYPEDESC; {*}
  PUserType11 = ^tagARRAYDESC; {*}
  PUserType1 = ^tag_inner_PROPVARIANT; {*}
  wireHWND = ^_RemotableHandle; 
  PUINT1 = ^LongWord; {*}
  PByte1 = ^Byte; {*}
  PUserType2 = ^TGUID; {*}
  PUserType3 = ^_FILETIME; {*}
  POleVariant1 = ^OleVariant; {*}
  PUserType7 = ^tagTYPEATTR; {*}
  PUserType8 = ^tagFUNCDESC; {*}
  PUserType9 = ^tagVARDESC; {*}
  PUserType12 = ^tagTLIBATTR; {*}

{$ALIGN 8}
  _LARGE_INTEGER = record
    QuadPart: Int64;
  end;

  _ULARGE_INTEGER = record
    QuadPart: Largeuint;
  end;

{$ALIGN 4}
  _FILETIME = record
    dwLowDateTime: LongWord;
    dwHighDateTime: LongWord;
  end;

  tagCLIPDATA = record
    cbSize: LongWord;
    ulClipFmt: Integer;
    pClipData: ^Byte;
  end;

  tagBSTRBLOB = record
    cbSize: LongWord;
    pData: ^Byte;
  end;

  tagBLOB = record
    cbSize: LongWord;
    pBlobData: ^Byte;
  end;

  tagVersionedStream = record
    guidVersion: TGUID;
    pStream: IStream;
  end;


{$ALIGN 8}
  tagSTATSTG = record
    pwcsName: PWideChar;
    type_: LongWord;
    cbSize: _ULARGE_INTEGER;
    mtime: _FILETIME;
    ctime: _FILETIME;
    atime: _FILETIME;
    grfMode: LongWord;
    grfLocksSupported: LongWord;
    clsid: TGUID;
    grfStateBits: LongWord;
    reserved: LongWord;
  end;


{$ALIGN 4}
  tagRemSNB = record
    ulCntStr: LongWord;
    ulCntChar: LongWord;
    rgString: ^Word;
  end;

  tagCAC = record
    cElems: LongWord;
    pElems: ^Shortint;
  end;

  tagCAUB = record
    cElems: LongWord;
    pElems: ^Byte;
  end;


  _wireSAFEARR_BSTR = record
    Size: LongWord;
    aBstr: ^PUserType5;
  end;

  _wireSAFEARR_UNKNOWN = record
    Size: LongWord;
    apUnknown: ^IUnknown;
  end;

  _wireSAFEARR_DISPATCH = record
    Size: LongWord;
    apDispatch: ^IDispatch;
  end;

  _FLAGGED_WORD_BLOB = record
    fFlags: LongWord;
    clSize: LongWord;
    asData: ^Word;
  end;


  _wireSAFEARR_VARIANT = record
    Size: LongWord;
    aVariant: ^PUserType6;
  end;


  _wireBRECORD = record
    fFlags: LongWord;
    clSize: LongWord;
    pRecInfo: IRecordInfo;
    pRecord: ^Byte;
  end;


  __MIDL_IOleAutomationTypes_0005 = record
    case Integer of
      0: (lptdesc: PUserType10);
      1: (lpadesc: PUserType11);
      2: (hreftype: LongWord);
  end;

  tagTYPEDESC = record
    DUMMYUNIONNAME: __MIDL_IOleAutomationTypes_0005;
    vt: Word;
  end;

  tagSAFEARRAYBOUND = record
    cElements: LongWord;
    lLbound: Integer;
  end;

  ULONG_PTR = LongWord; 

  tagIDLDESC = record
    dwReserved: ULONG_PTR;
    wIDLFlags: Word;
  end;

  DWORD = LongWord; 

{$ALIGN 8}
  tagPARAMDESCEX = record
    cBytes: LongWord;
    varDefaultValue: OleVariant;
  end;

{$ALIGN 4}
  tagPARAMDESC = record
    pparamdescex: ^tagPARAMDESCEX;
    wParamFlags: Word;
  end;

  tagELEMDESC = record
    tdesc: tagTYPEDESC;
    paramdesc: tagPARAMDESC;
  end;

  tagFUNCDESC = record
    memid: Integer;
    lprgscode: ^SCODE;
    lprgelemdescParam: ^tagELEMDESC;
    funckind: tagFUNCKIND;
    invkind: tagINVOKEKIND;
    callconv: tagCALLCONV;
    cParams: Smallint;
    cParamsOpt: Smallint;
    oVft: Smallint;
    cScodes: Smallint;
    elemdescFunc: tagELEMDESC;
    wFuncFlags: Word;
  end;

  __MIDL_IOleAutomationTypes_0006 = record
    case Integer of
      0: (oInst: LongWord);
      1: (lpvarValue: ^OleVariant);
  end;

  tagVARDESC = record
    memid: Integer;
    lpstrSchema: PWideChar;
    DUMMYUNIONNAME: __MIDL_IOleAutomationTypes_0006;
    elemdescVar: tagELEMDESC;
    wVarFlags: Word;
    varkind: tagVARKIND;
  end;

  tagTLIBATTR = record
    guid: TGUID;
    lcid: LongWord;
    syskind: tagSYSKIND;
    wMajorVerNum: Word;
    wMinorVerNum: Word;
    wLibFlags: Word;
  end;

  _wireSAFEARR_BRECORD = record
    Size: LongWord;
    aRecord: ^PUserType13;
  end;

  _wireSAFEARR_HAVEIID = record
    Size: LongWord;
    apUnknown: ^IUnknown;
    iid: TGUID;
  end;

  _BYTE_SIZEDARR = record
    clSize: LongWord;
    pData: ^Byte;
  end;

  _SHORT_SIZEDARR = record
    clSize: LongWord;
    pData: ^Word;
  end;

  _LONG_SIZEDARR = record
    clSize: LongWord;
    pData: ^LongWord;
  end;

  _HYPER_SIZEDARR = record
    clSize: LongWord;
    pData: ^Int64;
  end;

  tagCAI = record
    cElems: LongWord;
    pElems: ^Smallint;
  end;

  tagCAUI = record
    cElems: LongWord;
    pElems: ^Word;
  end;

  tagCAL = record
    cElems: LongWord;
    pElems: ^Integer;
  end;

  tagCAUL = record
    cElems: LongWord;
    pElems: ^LongWord;
  end;

  tagCAH = record
    cElems: LongWord;
    pElems: ^_LARGE_INTEGER;
  end;

  tagCAUH = record
    cElems: LongWord;
    pElems: ^_ULARGE_INTEGER;
  end;

  tagCAFLT = record
    cElems: LongWord;
    pElems: ^Single;
  end;

  tagCADBL = record
    cElems: LongWord;
    pElems: ^Double;
  end;

  tagCABOOL = record
    cElems: LongWord;
    pElems: ^WordBool;
  end;

  tagCASCODE = record
    cElems: LongWord;
    pElems: ^SCODE;
  end;

  tagCACY = record
    cElems: LongWord;
    pElems: ^Currency;
  end;

  tagCADATE = record
    cElems: LongWord;
    pElems: ^TDateTime;
  end;

  tagCAFILETIME = record
    cElems: LongWord;
    pElems: ^_FILETIME;
  end;

  tagCACLSID = record
    cElems: LongWord;
    pElems: ^TGUID;
  end;

  tagCACLIPDATA = record
    cElems: LongWord;
    pElems: ^tagCLIPDATA;
  end;

  tagCABSTR = record
    cElems: LongWord;
    pElems: ^WideString;
  end;

  tagCABSTRBLOB = record
    cElems: LongWord;
    pElems: ^tagBSTRBLOB;
  end;

  tagCALPSTR = record
    cElems: LongWord;
    pElems: ^PAnsiChar;
  end;

  tagCALPWSTR = record
    cElems: LongWord;
    pElems: ^PWideChar;
  end;


  tagCAPROPVARIANT = record
    cElems: LongWord;
    pElems: PUserType1;
  end;

{$ALIGN 8}
  __MIDL___MIDL_itf_scanprofiles_0001_0066_0001 = record
    case Integer of
      0: (cVal: Shortint);
      1: (bVal: Byte);
      2: (iVal: Smallint);
      3: (uiVal: Word);
      4: (lVal: Integer);
      5: (ulVal: LongWord);
      6: (intVal: SYSINT);
      7: (uintVal: SYSUINT);
      8: (hVal: _LARGE_INTEGER);
      9: (uhVal: _ULARGE_INTEGER);
      10: (fltVal: Single);
      11: (dblVal: Double);
      12: (boolVal: WordBool);
      13: (__OBSOLETE__VARIANT_BOOL: WordBool);
      14: (scode: SCODE);
      15: (cyVal: Currency);
      16: (date: TDateTime);
      17: (filetime: _FILETIME);
      18: (puuid: ^TGUID);
      19: (pClipData: ^tagCLIPDATA);
      20: (bstrVal: {NOT_UNION(WideString)}Pointer);
      21: (bstrblobVal: tagBSTRBLOB);
      22: (blob: tagBLOB);
      23: (pszVal: PAnsiChar);
      24: (pwszVal: PWideChar);
      25: (punkVal: {NOT_UNION(IUnknown)}Pointer);
      26: (pdispVal: {NOT_UNION(IDispatch)}Pointer);
      27: (pStream: {NOT_UNION(IStream)}Pointer);
      28: (pStorage: {NOT_UNION(IStorage)}Pointer);
      29: (pVersionedStream: ^tagVersionedStream);
      30: (parray: wirePSAFEARRAY);
      31: (cac: tagCAC);
      32: (caub: tagCAUB);
      33: (cai: tagCAI);
      34: (caui: tagCAUI);
      35: (cal: tagCAL);
      36: (caul: tagCAUL);
      37: (cah: tagCAH);
      38: (cauh: tagCAUH);
      39: (caflt: tagCAFLT);
      40: (cadbl: tagCADBL);
      41: (cabool: tagCABOOL);
      42: (cascode: tagCASCODE);
      43: (cacy: tagCACY);
      44: (cadate: tagCADATE);
      45: (cafiletime: tagCAFILETIME);
      46: (cauuid: tagCACLSID);
      47: (caclipdata: tagCACLIPDATA);
      48: (cabstr: tagCABSTR);
      49: (cabstrblob: tagCABSTRBLOB);
      50: (calpstr: tagCALPSTR);
      51: (calpwstr: tagCALPWSTR);
      52: (capropvar: tagCAPROPVARIANT);
      53: (pcVal: ^Shortint);
      54: (pbVal: ^Byte);
      55: (piVal: ^Smallint);
      56: (puiVal: ^Word);
      57: (plVal: ^Integer);
      58: (pulVal: ^LongWord);
      59: (pintVal: ^SYSINT);
      60: (puintVal: ^SYSUINT);
      61: (pfltVal: ^Single);
      62: (pdblVal: ^Double);
      63: (pboolVal: ^WordBool);
      64: (pdecVal: ^TDecimal);
      65: (pscode: ^SCODE);
      66: (pcyVal: ^Currency);
      67: (pdate: ^TDateTime);
      68: (pbstrVal: ^WideString);
      69: (ppunkVal: {NOT_UNION(^IUnknown)}Pointer);
      70: (ppdispVal: {NOT_UNION(^IDispatch)}Pointer);
      71: (pparray: ^wirePSAFEARRAY);
      72: (pvarVal: PUserType1);
  end;


{$ALIGN 4}
  __MIDL_IWinTypes_0009 = record
    case Integer of
      0: (hInproc: Integer);
      1: (hRemote: Integer);
  end;

  _RemotableHandle = record
    fContext: Integer;
    u: __MIDL_IWinTypes_0009;
  end;


{$ALIGN 8}
  tag_inner_PROPVARIANT = record
    vt: Word;
    wReserved1: Byte;
    wReserved2: Byte;
    wReserved3: LongWord;
    __MIDL____MIDL_itf_scanprofiles_0001_00660001: __MIDL___MIDL_itf_scanprofiles_0001_0066_0001;
  end;


  __MIDL_IOleAutomationTypes_0004 = record
    case Integer of
      0: (llVal: Int64);
      1: (lVal: Integer);
      2: (bVal: Byte);
      3: (iVal: Smallint);
      4: (fltVal: Single);
      5: (dblVal: Double);
      6: (boolVal: WordBool);
      7: (scode: SCODE);
      8: (cyVal: Currency);
      9: (date: TDateTime);
      10: (bstrVal: ^_FLAGGED_WORD_BLOB);
      11: (punkVal: {NOT_UNION(IUnknown)}Pointer);
      12: (pdispVal: {NOT_UNION(IDispatch)}Pointer);
      13: (parray: ^PUserType4);
      14: (brecVal: ^_wireBRECORD);
      15: (pbVal: ^Byte);
      16: (piVal: ^Smallint);
      17: (plVal: ^Integer);
      18: (pllVal: ^Int64);
      19: (pfltVal: ^Single);
      20: (pdblVal: ^Double);
      21: (pboolVal: ^WordBool);
      22: (pscode: ^SCODE);
      23: (pcyVal: ^Currency);
      24: (pdate: ^TDateTime);
      25: (pbstrVal: ^PUserType5);
      26: (ppunkVal: {NOT_UNION(^IUnknown)}Pointer);
      27: (ppdispVal: {NOT_UNION(^IDispatch)}Pointer);
      28: (pparray: ^PPUserType1);
      29: (pvarVal: ^PUserType6);
      30: (cVal: Shortint);
      31: (uiVal: Word);
      32: (ulVal: LongWord);
      33: (ullVal: Largeuint);
      34: (intVal: SYSINT);
      35: (uintVal: SYSUINT);
      36: (decVal: TDecimal);
      37: (pdecVal: ^TDecimal);
      38: (pcVal: ^Shortint);
      39: (puiVal: ^Word);
      40: (pulVal: ^LongWord);
      41: (pullVal: ^Largeuint);
      42: (pintVal: ^SYSINT);
      43: (puintVal: ^SYSUINT);
  end;

{$ALIGN 4}
  __MIDL_IOleAutomationTypes_0001 = record
    case Integer of
      0: (BstrStr: _wireSAFEARR_BSTR);
      1: (UnknownStr: _wireSAFEARR_UNKNOWN);
      2: (DispatchStr: _wireSAFEARR_DISPATCH);
      3: (VariantStr: _wireSAFEARR_VARIANT);
      4: (RecordStr: _wireSAFEARR_BRECORD);
      5: (HaveIidStr: _wireSAFEARR_HAVEIID);
      6: (ByteStr: _BYTE_SIZEDARR);
      7: (WordStr: _SHORT_SIZEDARR);
      8: (LongStr: _LONG_SIZEDARR);
      9: (HyperStr: _HYPER_SIZEDARR);
  end;

  _wireSAFEARRAY_UNION = record
    sfType: LongWord;
    u: __MIDL_IOleAutomationTypes_0001;
  end;

{$ALIGN 8}
  _wireVARIANT = record
    clSize: LongWord;
    rpcReserved: LongWord;
    vt: Word;
    wReserved1: Word;
    wReserved2: Word;
    wReserved3: Word;
    DUMMYUNIONNAME: __MIDL_IOleAutomationTypes_0004;
  end;


{$ALIGN 4}
  tagTYPEATTR = record
    guid: TGUID;
    lcid: LongWord;
    dwReserved: LongWord;
    memidConstructor: Integer;
    memidDestructor: Integer;
    lpstrSchema: PWideChar;
    cbSizeInstance: LongWord;
    typekind: tagTYPEKIND;
    cFuncs: Word;
    cVars: Word;
    cImplTypes: Word;
    cbSizeVft: Word;
    cbAlignment: Word;
    wTypeFlags: Word;
    wMajorVerNum: Word;
    wMinorVerNum: Word;
    tdescAlias: tagTYPEDESC;
    idldescType: tagIDLDESC;
  end;

  tagARRAYDESC = record
    tdescElem: tagTYPEDESC;
    cDims: Word;
    rgbounds: ^tagSAFEARRAYBOUND;
  end;


  _wireSAFEARRAY = record
    cDims: Word;
    fFeatures: Word;
    cbElements: LongWord;
    cLocks: LongWord;
    uArrayStructs: _wireSAFEARRAY_UNION;
    rgsabound: ^tagSAFEARRAYBOUND;
  end;


// *********************************************************************//
// Interface: IScanProfileMgr
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {34EAAE27-2D89-4278-84EF-61DEFA323FC1}
// *********************************************************************//
  IScanProfileMgr = interface(IDispatch)
    ['{34EAAE27-2D89-4278-84EF-61DEFA323FC1}']
    procedure GetNumProfiles(out pulNumProfiles: LongWord); safecall;
    procedure GetNumProfilesforDeviceID(const bstrDeviceID: WideString; out pulNumProfiles: LongWord); safecall;
    procedure GetProfiles(var pulNumProfiles: LongWord; out ppScanProfile: IScanProfile); safecall;
    procedure GetProfilesforDeviceID(const bstrDeviceID: WideString; var pulNumProfiles: LongWord; 
                                     out ppScanProfile: IScanProfile); safecall;
    procedure GetDefaultProfile(const bstrDeviceID: WideString; out ppScanProfile: IScanProfile); safecall;
    procedure CreateProfile(const bstrDeviceID: WideString; const bstrName: WideString; 
                            guidCategory: TGUID; out ppScanProfile: IScanProfile); safecall;
    procedure OpenProfile(guid: TGUID; out ppScanProfile: IScanProfile); safecall;
    procedure SetDefault(const pScanProfile: IScanProfile); safecall;
    procedure DeleteProfile(const pScanProfile: IScanProfile); safecall;
    procedure DeleteAllProfiles(const bstrDeviceID: WideString); safecall;
    procedure DeleteAllProfilesForUser; safecall;
    procedure Refresh; safecall;
  end;

// *********************************************************************//
// DispIntf:  IScanProfileMgrDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {34EAAE27-2D89-4278-84EF-61DEFA323FC1}
// *********************************************************************//
  IScanProfileMgrDisp = dispinterface
    ['{34EAAE27-2D89-4278-84EF-61DEFA323FC1}']
    procedure GetNumProfiles(out pulNumProfiles: LongWord); dispid 1610743808;
    procedure GetNumProfilesforDeviceID(const bstrDeviceID: WideString; out pulNumProfiles: LongWord); dispid 1610743809;
    procedure GetProfiles(var pulNumProfiles: LongWord; out ppScanProfile: IScanProfile); dispid 1610743810;
    procedure GetProfilesforDeviceID(const bstrDeviceID: WideString; var pulNumProfiles: LongWord; 
                                     out ppScanProfile: IScanProfile); dispid 1610743811;
    procedure GetDefaultProfile(const bstrDeviceID: WideString; out ppScanProfile: IScanProfile); dispid 1610743812;
    procedure CreateProfile(const bstrDeviceID: WideString; const bstrName: WideString; 
                            guidCategory: {NOT_OLEAUTO(TGUID)}OleVariant; 
                            out ppScanProfile: IScanProfile); dispid 1610743813;
    procedure OpenProfile(guid: {NOT_OLEAUTO(TGUID)}OleVariant; out ppScanProfile: IScanProfile); dispid 1610743814;
    procedure SetDefault(const pScanProfile: IScanProfile); dispid 1610743815;
    procedure DeleteProfile(const pScanProfile: IScanProfile); dispid 1610743816;
    procedure DeleteAllProfiles(const bstrDeviceID: WideString); dispid 1610743817;
    procedure DeleteAllProfilesForUser; dispid 1610743818;
    procedure Refresh; dispid 1610743819;
  end;

// *********************************************************************//
// Interface: IScanProfile
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {D68A6C07-9FF8-47D1-9A2A-429D28FBC6A4}
// *********************************************************************//
  IScanProfile = interface(IDispatch)
    ['{D68A6C07-9FF8-47D1-9A2A-429D28FBC6A4}']
    procedure GetGUID(out pGUID: TGUID); safecall;
    procedure GetDeviceID(out pbstrDeviceID: WideString); safecall;
    procedure IsDefault(out pbDefault: Integer); safecall;
    procedure GetProperty(num: LongWord; var pid: LongWord; out pvar: tag_inner_PROPVARIANT); safecall;
    procedure SetProperty(num: LongWord; var pid: LongWord; var pvar: tag_inner_PROPVARIANT); safecall;
    procedure GetAllPropIDs(var num: LongWord; out ppid: LongWord); safecall;
    procedure GetNumPropIDS(out num: LongWord); safecall;
    procedure GetName(out pbstrName: WideString); safecall;
    procedure SetName(const pbstrName: WideString); safecall;
    procedure GetItem(out pguidCategory: TGUID); safecall;
    procedure SetItem(guidCategory: TGUID); safecall;
    procedure Save; safecall;
    procedure RemoveProperty(num: LongWord; var pid: LongWord); safecall;
  end;

// *********************************************************************//
// DispIntf:  IScanProfileDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {D68A6C07-9FF8-47D1-9A2A-429D28FBC6A4}
// *********************************************************************//
  IScanProfileDisp = dispinterface
    ['{D68A6C07-9FF8-47D1-9A2A-429D28FBC6A4}']
    procedure GetGUID(out pGUID: {NOT_OLEAUTO(TGUID)}OleVariant); dispid 1610743808;
    procedure GetDeviceID(out pbstrDeviceID: WideString); dispid 1610743809;
    procedure IsDefault(out pbDefault: Integer); dispid 1610743810;
    procedure GetProperty(num: LongWord; var pid: LongWord; 
                          out pvar: {NOT_OLEAUTO(tag_inner_PROPVARIANT)}OleVariant); dispid 1610743811;
    procedure SetProperty(num: LongWord; var pid: LongWord; 
                          var pvar: {NOT_OLEAUTO(tag_inner_PROPVARIANT)}OleVariant); dispid 1610743812;
    procedure GetAllPropIDs(var num: LongWord; out ppid: LongWord); dispid 1610743813;
    procedure GetNumPropIDS(out num: LongWord); dispid 1610743814;
    procedure GetName(out pbstrName: WideString); dispid 1610743815;
    procedure SetName(const pbstrName: WideString); dispid 1610743816;
    procedure GetItem(out pguidCategory: {NOT_OLEAUTO(TGUID)}OleVariant); dispid 1610743817;
    procedure SetItem(guidCategory: {NOT_OLEAUTO(TGUID)}OleVariant); dispid 1610743818;
    procedure Save; dispid 1610743819;
    procedure RemoveProperty(num: LongWord; var pid: LongWord); dispid 1610743820;
  end;

// *********************************************************************//
// Interface: IStorage
// Flags:     (0)
// GUID:      {0000000B-0000-0000-C000-000000000046}
// *********************************************************************//
  IStorage = interface(IUnknown)
    ['{0000000B-0000-0000-C000-000000000046}']
    function CreateStream(pwcsName: PWideChar; grfMode: LongWord; reserved1: LongWord; 
                          reserved2: LongWord; out ppstm: IStream): HResult; stdcall;
    function RemoteOpenStream(pwcsName: PWideChar; cbReserved1: LongWord; var reserved1: Byte; 
                              grfMode: LongWord; reserved2: LongWord; out ppstm: IStream): HResult; stdcall;
    function CreateStorage(pwcsName: PWideChar; grfMode: LongWord; reserved1: LongWord; 
                           reserved2: LongWord; out ppstg: IStorage): HResult; stdcall;
    function OpenStorage(pwcsName: PWideChar; const pstgPriority: IStorage; grfMode: LongWord; 
                         var snbExclude: tagRemSNB; reserved: LongWord; out ppstg: IStorage): HResult; stdcall;
    function RemoteCopyTo(ciidExclude: LongWord; var rgiidExclude: TGUID; 
                          var snbExclude: tagRemSNB; const pstgDest: IStorage): HResult; stdcall;
    function MoveElementTo(pwcsName: PWideChar; const pstgDest: IStorage; pwcsNewName: PWideChar; 
                           grfFlags: LongWord): HResult; stdcall;
    function Commit(grfCommitFlags: LongWord): HResult; stdcall;
    function Revert: HResult; stdcall;
    function RemoteEnumElements(reserved1: LongWord; cbReserved2: LongWord; var reserved2: Byte; 
                                reserved3: LongWord; out ppenum: IEnumSTATSTG): HResult; stdcall;
    function DestroyElement(pwcsName: PWideChar): HResult; stdcall;
    function RenameElement(pwcsOldName: PWideChar; pwcsNewName: PWideChar): HResult; stdcall;
    function SetElementTimes(pwcsName: PWideChar; var pctime: _FILETIME; var patime: _FILETIME; 
                             var pmtime: _FILETIME): HResult; stdcall;
    function SetClass(var clsid: TGUID): HResult; stdcall;
    function SetStateBits(grfStateBits: LongWord; grfMask: LongWord): HResult; stdcall;
    function Stat(out pstatstg: tagSTATSTG; grfStatFlag: LongWord): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IEnumSTATSTG
// Flags:     (0)
// GUID:      {0000000D-0000-0000-C000-000000000046}
// *********************************************************************//
  IEnumSTATSTG = interface(IUnknown)
    ['{0000000D-0000-0000-C000-000000000046}']
    function RemoteNext(celt: LongWord; out rgelt: tagSTATSTG; out pceltFetched: LongWord): HResult; stdcall;
    function Skip(celt: LongWord): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppenum: IEnumSTATSTG): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IRecordInfo
// Flags:     (0)
// GUID:      {0000002F-0000-0000-C000-000000000046}
// *********************************************************************//
  IRecordInfo = interface(IUnknown)
    ['{0000002F-0000-0000-C000-000000000046}']
    function RecordInit(pvNew: Pointer): HResult; stdcall;
    function RecordClear(pvExisting: Pointer): HResult; stdcall;
    function RecordCopy(pvExisting: Pointer; pvNew: Pointer): HResult; stdcall;
    function GetGUID(out pGUID: TGUID): HResult; stdcall;
    function GetName(out pbstrName: WideString): HResult; stdcall;
    function GetSize(out pcbSize: LongWord): HResult; stdcall;
    function GetTypeInfo(out ppTypeInfo: ITypeInfo): HResult; stdcall;
    function GetField(pvData: Pointer; szFieldName: PWideChar; out pvarField: OleVariant): HResult; stdcall;
    function GetFieldNoCopy(pvData: Pointer; szFieldName: PWideChar; out pvarField: OleVariant; 
                            out ppvDataCArray: Pointer): HResult; stdcall;
    function PutField(wFlags: LongWord; pvData: Pointer; szFieldName: PWideChar; 
                      const pvarField: OleVariant): HResult; stdcall;
    function PutFieldNoCopy(wFlags: LongWord; pvData: Pointer; szFieldName: PWideChar; 
                            const pvarField: OleVariant): HResult; stdcall;
    function GetFieldNames(var pcNames: LongWord; out rgBstrNames: WideString): HResult; stdcall;
    function IsMatchingType(const pRecordInfo: IRecordInfo): Integer; stdcall;
    function RecordCreate: Pointer; stdcall;
    function RecordCreateCopy(pvSource: Pointer; out ppvDest: Pointer): HResult; stdcall;
    function RecordDestroy(pvRecord: Pointer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: ITypeInfo
// Flags:     (0)
// GUID:      {00020401-0000-0000-C000-000000000046}
// *********************************************************************//
  ITypeInfo = interface(IUnknown)
    ['{00020401-0000-0000-C000-000000000046}']
    function RemoteGetTypeAttr(out ppTypeAttr: PUserType7; out pDummy: DWORD): HResult; stdcall;
    function GetTypeComp(out ppTComp: ITypeComp): HResult; stdcall;
    function RemoteGetFuncDesc(index: SYSUINT; out ppFuncDesc: PUserType8; out pDummy: DWORD): HResult; stdcall;
    function RemoteGetVarDesc(index: SYSUINT; out ppVarDesc: PUserType9; out pDummy: DWORD): HResult; stdcall;
    function RemoteGetNames(memid: Integer; out rgBstrNames: WideString; cMaxNames: SYSUINT; 
                            out pcNames: SYSUINT): HResult; stdcall;
    function GetRefTypeOfImplType(index: SYSUINT; out pRefType: LongWord): HResult; stdcall;
    function GetImplTypeFlags(index: SYSUINT; out pImplTypeFlags: SYSINT): HResult; stdcall;
    function LocalGetIDsOfNames: HResult; stdcall;
    function LocalInvoke: HResult; stdcall;
    function RemoteGetDocumentation(memid: Integer; refPtrFlags: LongWord; 
                                    out pbstrName: WideString; out pBstrDocString: WideString; 
                                    out pdwHelpContext: LongWord; out pBstrHelpFile: WideString): HResult; stdcall;
    function RemoteGetDllEntry(memid: Integer; invkind: tagINVOKEKIND; refPtrFlags: LongWord; 
                               out pBstrDllName: WideString; out pbstrName: WideString; 
                               out pwOrdinal: Word): HResult; stdcall;
    function GetRefTypeInfo(hreftype: LongWord; out ppTInfo: ITypeInfo): HResult; stdcall;
    function LocalAddressOfMember: HResult; stdcall;
    function RemoteCreateInstance(var riid: TGUID; out ppvObj: IUnknown): HResult; stdcall;
    function GetMops(memid: Integer; out pBstrMops: WideString): HResult; stdcall;
    function RemoteGetContainingTypeLib(out ppTLib: ITypeLib; out pIndex: SYSUINT): HResult; stdcall;
    function LocalReleaseTypeAttr: HResult; stdcall;
    function LocalReleaseFuncDesc: HResult; stdcall;
    function LocalReleaseVarDesc: HResult; stdcall;
  end;

// *********************************************************************//
// Interface: ITypeComp
// Flags:     (0)
// GUID:      {00020403-0000-0000-C000-000000000046}
// *********************************************************************//
  ITypeComp = interface(IUnknown)
    ['{00020403-0000-0000-C000-000000000046}']
    function RemoteBind(szName: PWideChar; lHashVal: LongWord; wFlags: Word; 
                        out ppTInfo: ITypeInfo; out pDescKind: tagDESCKIND; 
                        out ppFuncDesc: PUserType8; out ppVarDesc: PUserType9; 
                        out ppTypeComp: ITypeComp; out pDummy: DWORD): HResult; stdcall;
    function RemoteBindType(szName: PWideChar; lHashVal: LongWord; out ppTInfo: ITypeInfo): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: ITypeLib
// Flags:     (0)
// GUID:      {00020402-0000-0000-C000-000000000046}
// *********************************************************************//
  ITypeLib = interface(IUnknown)
    ['{00020402-0000-0000-C000-000000000046}']
    function RemoteGetTypeInfoCount(out pcTInfo: SYSUINT): HResult; stdcall;
    function GetTypeInfo(index: SYSUINT; out ppTInfo: ITypeInfo): HResult; stdcall;
    function GetTypeInfoType(index: SYSUINT; out pTKind: tagTYPEKIND): HResult; stdcall;
    function GetTypeInfoOfGuid(var guid: TGUID; out ppTInfo: ITypeInfo): HResult; stdcall;
    function RemoteGetLibAttr(out ppTLibAttr: PUserType12; out pDummy: DWORD): HResult; stdcall;
    function GetTypeComp(out ppTComp: ITypeComp): HResult; stdcall;
    function RemoteGetDocumentation(index: SYSINT; refPtrFlags: LongWord; 
                                    out pbstrName: WideString; out pBstrDocString: WideString; 
                                    out pdwHelpContext: LongWord; out pBstrHelpFile: WideString): HResult; stdcall;
    function RemoteIsName(szNameBuf: PWideChar; lHashVal: LongWord; out pfName: Integer; 
                          out pBstrLibName: WideString): HResult; stdcall;
    function RemoteFindName(szNameBuf: PWideChar; lHashVal: LongWord; out ppTInfo: ITypeInfo; 
                            out rgMemId: Integer; var pcFound: Word; out pBstrLibName: WideString): HResult; stdcall;
    function LocalReleaseTLibAttr: HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IScanProfileUI
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {B67CDDB7-2B20-473E-8D6C-3F1BD202E78D}
// *********************************************************************//
  IScanProfileUI = interface(IDispatch)
    ['{B67CDDB7-2B20-473E-8D6C-3F1BD202E78D}']
    procedure ScanProfileDialog(var hwndParent: _RemotableHandle); safecall;
  end;

// *********************************************************************//
// DispIntf:  IScanProfileUIDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {B67CDDB7-2B20-473E-8D6C-3F1BD202E78D}
// *********************************************************************//
  IScanProfileUIDisp = dispinterface
    ['{B67CDDB7-2B20-473E-8D6C-3F1BD202E78D}']
    procedure ScanProfileDialog(var hwndParent: {NOT_OLEAUTO(_RemotableHandle)}OleVariant); dispid 1610743808;
  end;

// *********************************************************************//
// The Class CoScanProfileMgr provides a Create and CreateRemote method to          
// create instances of the default interface IScanProfileMgr exposed by              
// the CoClass ScanProfileMgr. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoScanProfileMgr = class
    class function Create: IScanProfileMgr;
    class function CreateRemote(const MachineName: string): IScanProfileMgr;
  end;

// *********************************************************************//
// The Class CoScanProfileUI provides a Create and CreateRemote method to          
// create instances of the default interface IScanProfileUI exposed by              
// the CoClass ScanProfileUI. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoScanProfileUI = class
    class function Create: IScanProfileUI;
    class function CreateRemote(const MachineName: string): IScanProfileUI;
  end;

implementation

uses ComObj;

class function CoScanProfileMgr.Create: IScanProfileMgr;
begin
  Result := CreateComObject(CLASS_ScanProfileMgr) as IScanProfileMgr;
end;

class function CoScanProfileMgr.CreateRemote(const MachineName: string): IScanProfileMgr;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ScanProfileMgr) as IScanProfileMgr;
end;

class function CoScanProfileUI.Create: IScanProfileUI;
begin
  Result := CreateComObject(CLASS_ScanProfileUI) as IScanProfileUI;
end;

class function CoScanProfileUI.CreateRemote(const MachineName: string): IScanProfileUI;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ScanProfileUI) as IScanProfileUI;
end;

end.
