(****************************************************************************
*                FreePascal \ Delphi WIA Implementation
*
*  FILE: WIA_PaperSizes.pas
*
*  VERSION:     0.0.2
*
*  DESCRIPTION:
*    WIA Paper Sizes consts in thousandths of an inch.
*
*****************************************************************************
*
*  (c) 2024 Massimo Magnano
*
*  See changelog.txt for Change Log
*
*****************************************************************************)
unit WIA_PaperSizes;

{$H+}

interface

uses WIADef;

type
  {Paper size}
  TPaperSize = packed record
    name: String[16];
    w, h: Integer;
  end;

  TWIAPaperSize = (
   wpsMAX = $FF, // Use  WIA_IPS_MAX_HORIZONTAL/VERTICAL_SIZE
   wpsA4 = WIA_PAGE_A4, //  8267 x 11692
   wpsLETTER = WIA_PAGE_LETTER, //  8500 x 11000
   wpsCUSTOM = WIA_PAGE_CUSTOM, // Use a Range from  WIA_IPS_MIN_*_SIZE to WIA_IPS_MAX_*_SIZE
   wpsUSLEGAL = WIA_PAGE_USLEGAL, //  8500 x 14000
   wpsUSLEDGER = WIA_PAGE_USLEDGER, // 11000 x 17000
   wpsUSSTATEMENT = WIA_PAGE_USSTATEMENT, //  5500 x  8500
   wpsBUSINESSCARD = WIA_PAGE_BUSINESSCARD, //  3543 x  2165
   wpsISO_A0 = WIA_PAGE_ISO_A0, // 33110 x 46811
   wpsISO_A1 = WIA_PAGE_ISO_A1, // 23385 x 33110
   wpsISO_A2 = WIA_PAGE_ISO_A2, // 16535 x 23385
   wpsISO_A3 = WIA_PAGE_ISO_A3, // 11692 x 16535
   wpsISO_A5 = WIA_PAGE_ISO_A5, //  5826 x  8267
   wpsISO_A6 = WIA_PAGE_ISO_A6, //  4133 x  5826
   wpsISO_A7 = WIA_PAGE_ISO_A7, //  2913 x  4133
   wpsISO_A8 = WIA_PAGE_ISO_A8, //  2047 x  2913
   wpsISO_A9 = WIA_PAGE_ISO_A9, //  1456 x  2047
   wpsISO_A10 = WIA_PAGE_ISO_A10, //  1023 x  1456
   wpsISO_B0 = WIA_PAGE_ISO_B0, //  39370 x 55669
   wpsISO_B1 = WIA_PAGE_ISO_B1, //  27834 x 39370
   wpsISO_B2 = WIA_PAGE_ISO_B2, //  19685 x 27834
   wpsISO_B3 = WIA_PAGE_ISO_B3, //  13897 x 19685
   wpsISO_B4 = WIA_PAGE_ISO_B4, //   9842 x 13897
   wpsISO_B5 = WIA_PAGE_ISO_B5, //   6929 x  9842
   wpsISO_B6 = WIA_PAGE_ISO_B6, //   4921 x  6929
   wpsISO_B7 = WIA_PAGE_ISO_B7, //   3464 x  4921
   wpsISO_B8 = WIA_PAGE_ISO_B8, //   2440 x  3464
   wpsISO_B9 = WIA_PAGE_ISO_B9, //   1732 x  2440
   wpsISO_B10 = WIA_PAGE_ISO_B10, //   1220 x  1732
   wpsISO_C0 = WIA_PAGE_ISO_C0, //  36102 x 51062
   wpsISO_C1 = WIA_PAGE_ISO_C1, //  25511 x 36102
   wpsISO_C2 = WIA_PAGE_ISO_C2, //  18031 x 25511
   wpsISO_C3 = WIA_PAGE_ISO_C3, //  12755 x 18031
   wpsISO_C4 = WIA_PAGE_ISO_C4, //   9015 x 12755 (unfolded)
   wpsISO_C5 = WIA_PAGE_ISO_C5, //   6377 x  9015 (folded once)
   wpsISO_C6 = WIA_PAGE_ISO_C6, //   4488 x  6377 (folded twice)
   wpsISO_C7 = WIA_PAGE_ISO_C7, //   3188 x  4488
   wpsISO_C8 = WIA_PAGE_ISO_C8, //   2244 x  3188
   wpsISO_C9 = WIA_PAGE_ISO_C9, //   1574 x  2244
   wpsISO_C10 = WIA_PAGE_ISO_C10, //   1102 x  1574
   wpsJIS_B0 = WIA_PAGE_JIS_B0, //  40551 x 57322
   wpsJIS_B1 = WIA_PAGE_JIS_B1, //  28661 x 40551
   wpsJIS_B2 = WIA_PAGE_JIS_B2, //  20275 x 28661
   wpsJIS_B3 = WIA_PAGE_JIS_B3, //  14330 x 20275
   wpsJIS_B4 = WIA_PAGE_JIS_B4, //  10118 x 14330
   wpsJIS_B5 = WIA_PAGE_JIS_B5, //   7165 x 10118
   wpsJIS_B6 = WIA_PAGE_JIS_B6, //   5039 x  7165
   wpsJIS_B7 = WIA_PAGE_JIS_B7, //   3582 x  5039
   wpsJIS_B8 = WIA_PAGE_JIS_B8, //   2519 x  3582
   wpsJIS_B9 = WIA_PAGE_JIS_B9, //   1771 x  2519
   wpsJIS_B10 = WIA_PAGE_JIS_B10, //   1259 x  1771
   wpsJIS_2A = WIA_PAGE_JIS_2A, //  46811 x 66220
   wpsJIS_4A = WIA_PAGE_JIS_4A, //  66220 x  93622
   wpsDIN_2B = WIA_PAGE_DIN_2B, //  55669 x 78740
   wpsDIN_4B = WIA_PAGE_DIN_4B, //  78740 x 111338
   wpsAUTO = WIA_PAGE_AUTO
  );
  TWIAPaperSizeSet = set of TWIAPaperSize;

const
  PaperSizesWIA_MaxIndex = wpsDIN_4B;

  //Sizes of Papers in thousandths of an inch
  PaperSizesWIA: array [wpsA4..wpsDIN_4B] of TPaperSize = (
   (name:'A4'; w: 8267; h: 11692),
   (name:'US Letter'; w: 8500; h: 11000),
   (name:''; w: 0; h: 0),
   (name:'US Legal'; w: 8500; h: 14000),
   (name:'US Ledger'; w:11000; h: 17000),
   (name:'US Statement'; w: 5500; h:  8500),
   (name:'Business card'; w: 3543; h:  2165),
   (name:'A0'; w:33110; h: 46811),
   (name:'A1'; w:23385; h: 33110),
   (name:'A2'; w:16535; h: 23385),
   (name:'A3'; w:11692; h: 16535),
   (name:'A5'; w: 5826; h:  8267),
   (name:'A6'; w: 4133; h:  5826),
   (name:'A7'; w: 2913; h:  4133),
   (name:'A8'; w: 2047; h:  2913),
   (name:'A9'; w: 1456; h:  2047),
   (name:'A10'; w: 1023; h:  1456),
   (name:'B0'; w: 39370; h: 55669),
   (name:'B1'; w: 27834; h: 39370),
   (name:'B2'; w: 19685; h: 27834),
   (name:'B3'; w: 13897; h: 19685),
   (name:'B4'; w:  9842; h: 13897),
   (name:'B5'; w:  6929; h:  9842),
   (name:'B6'; w:  4921; h:  6929),
   (name:'B7'; w:  3464; h:  4921),
   (name:'B8'; w:  2440; h:  3464),
   (name:'B9'; w:  1732; h:  2440),
   (name:'B10'; w:  1220; h:  1732),
   (name:'C0'; w: 36102; h: 51062),
   (name:'C1'; w: 25511; h: 36102),
   (name:'C2'; w: 18031; h: 25511),
   (name:'C3'; w: 12755; h: 18031),
   (name:'C4'; w:  9015; h: 12755),
   (name:'C5'; w:  6377; h:  9015),
   (name:'C6'; w:  4488; h:  6377),
   (name:'C7'; w:  3188; h:  4488),
   (name:'C8'; w:  2244; h:  3188),
   (name:'C9'; w:  1574; h:  2244),
   (name:'C10'; w:  1102; h:  1574),
   (name:'JIS B0'; w: 40551; h: 57322),
   (name:'JIS B1'; w: 28661; h: 40551),
   (name:'JIS B2'; w: 20275; h: 28661),
   (name:'JIS B3'; w: 14330; h: 20275),
   (name:'JIS B4'; w: 10118; h: 14330),
   (name:'JIS B5'; w:  7165; h: 10118),
   (name:'JIS B6'; w:  5039; h:  7165),
   (name:'JIS B7'; w:  3582; h:  5039),
   (name:'JIS B8'; w:  2519; h:  3582),
   (name:'JIS B9'; w:  1771; h:  2519),
   (name:'JIS B10'; w:  1259; h:  1771),
   (name:'JIS 2A'; w: 46811; h: 66220),
   (name:'JIS 4A'; w: 66220; h:  93622),
   (name:'DIN 2B'; w: 55669; h: 78740),
   (name:'DIN 4B'; w: 78740; h: 111338)
  );

(* //in cm
   (name:'A4'; w:21.0; h:29.7),
   (name:'US Letter'; w:21.6; h:27.9),
   (name:''; w:0; h:0), //Place order for Custom
   (name:'US Legal'; w:21.6; h:35.6),
   (name:'US Ledger'; w:43.2; h:27.9),
   (name:'US Statement'; w:14.0; h:21.6),
   (name:'Business card'; w:9.0; h:5.5),
   (name:'A0'; w:84.1; h:118.9),
   (name:'A1'; w:59.4; h:84.1),
   (name:'A2'; w:42.0; h:59.4),
   (name:'A3'; w:29.7; h:42.0),
   (name:'A5'; w:14.8; h:21.0),
   (name:'A6'; w:4.1; h:5.8),
   (name:'A7'; w:7.4; h:10.5),
   (name:'A8'; w:5.2; h:7.4),
   (name:'A9'; w:3.7; h:5.2),
   (name:'A10'; w:2.6; h:3.7),
   (name:'B0'; w:100.0; h:141.4),
   (name:'B1'; w:70.7; h:100.0),
   (name:'B2'; w:50.0; h:70.7),
   (name:'B3'; w:35.3; h:50.0),
   (name:'B4'; w:25.0; h:35.3),
   (name:'B5'; w:17.6; h:25.0),
   (name:'B6'; w:12.5; h:17.6),
   (name:'B7'; w:8.8; h:12.5),
   (name:'B8'; w:6.2; h:8.8),
   (name:'B9'; w:4.4; h:6.2),
   (name:'B10'; w:3.1; h:4.4),
   (name:'C0'; w:91.7; h:129.7),
   (name:'C1'; w:64.8; h:91.7),
   (name:'C2'; w:45.8; h:64.8),
   (name:'C3'; w:32.4; h:45.8),
   (name:'C4'; w:22.9; h:32.4),
   (name:'C5'; w:16.2; h:22.9),
   (name:'C6'; w:11.4; h:16.2),
   (name:'C7'; w:8.1; h:11.4),
   (name:'C8'; w:5.7; h:8.1),
   (name:'C9'; w:4.0; h:5.7),
   (name:'C10'; w:2.8; h:4.0),
   (name:'JIS B0'; w:103.0; h:145.6),
   (name:'JIS B1'; w:72.8; h:103.0),
   (name:'JIS B2'; w:51.5; h:72.8),
   (name:'JIS B3'; w:36.4; h:51.5),
   (name:'JIS B4'; w:25.7; h:36.4),
   (name:'JIS B5'; w:18.2; h:25.7),
   (name:'JIS B6'; w:12.8; h:18.2),
   (name:'JIS B7'; w:9.1; h:12.8),
   (name:'JIS B8'; w:6.4; h:9.1),
   (name:'JIS B9'; w:4.5; h:6.4),
   (name:'JIS B10'; w:3.2; h:4.5),
   (name:'JIS 2A'; w:118.9; h:168.2),
   (name:'JIS 4A'; w:168.2; h:237.8),
   (name:'DIN 2B'; w:141.4; h:200),
   (name:'DIN 4B'; w:200; h:282.8)
*)

function CalculatePaperSizeSet(Max_Width, Max_Height: Integer): TWIAPaperSizeSet;
function CalculatePaperSize(AWidth, AHeight: Integer): TWIAPaperSize;
function THInchToCm(ASize: Integer): Single;
function THInchToCmStr(ASize: Integer): String;
function THInchToInch(ASize: Integer): Single;
function THInchToInchStr(ASize: Integer): String;

implementation

uses SysUtils;

function CalculatePaperSizeSet(Max_Width, Max_Height: Integer): TWIAPaperSizeSet;
var
   iSwap: Integer;
   i: TWIAPaperSize;

begin
  Result:= [wpsMAX];

  if (Max_Width > Max_Height) then
  begin
    iSwap:= Max_Height;
    Max_Height:= Max_Width;
    Max_Width:= iSwap;
  end;

  for i:=wpsA4 to PaperSizesWIA_MaxIndex do
  begin
    if (PaperSizesWIA[i].w <= Max_Width) and (PaperSizesWIA[i].h <= Max_Height)
    then Result:= Result + [i];
  end;
end;

function CalculatePaperSize(AWidth, AHeight: Integer): TWIAPaperSize;
var
   i: TWIAPaperSize;

begin
  Result:= wpsCUSTOM;
  for i:=wpsA4 to PaperSizesWIA_MaxIndex do
  begin
    if (PaperSizesWIA[i].w = AWidth) and (PaperSizesWIA[i].h = AHeight)
    then begin Result:= i; break; end;
  end;
end;

function THInchToCm(ASize: Integer): Single;
begin
  Result:= (ASize / 1000) * 2.54;
end;

function THInchToCmStr(ASize: Integer): String;
begin
  Result:= FloatToStrF((ASize / 1000) * 2.54, ffFixed, 15, 2);
end;

function THInchToInch(ASize: Integer): Single;
begin
  Result:= (ASize / 1000);
end;

function THInchToInchStr(ASize: Integer): String;
begin
  Result:= FloatToStrF((ASize / 1000), ffFixed, 15, 2);
end;

end.

