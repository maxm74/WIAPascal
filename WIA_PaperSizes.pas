(****************************************************************************
*                FreePascal \ Delphi WIA Implementation
*
*  FILE: WIA_PaperSizes.pas
*
*  VERSION:     0.0.2
*
*  DESCRIPTION:
*    WIA Paper Sizes consts in cm.
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

uses WIA;

const
  { #todo 10 -oMaxM : do it in inchs for use in Pages Sizes Set with IPS_XEXTENT }
  //Sizes of Papers in cm
  PaperSizesWIA: array [wpsA4..wpsDIN_4B] of TPaperSize =
   ((name:'A4'; w:21.0; h:29.7),
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
   );

implementation

end.

