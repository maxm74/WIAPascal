(****************************************************************************
*                FreePascal \ Delphi WIA Implementation
*
*  FILE: WIA_UI_Common.pas
*
*  VERSION:     1.0.3
*
*  DESCRIPTION:
*    WIA UI Basic Types and Functions for registering User Interfaces.
*
*****************************************************************************
*
*  (c) 2026 Massimo Magnano
*
*  See changelog.txt for Change Log
*
*****************************************************************************)

unit WIA_UI_Common;

interface

uses
  Classes, SysUtils, WIA;

resourcestring
  rsLandscape = 'Landscape';
  rsPortrait = 'Portrait';
  rsAutotype = 'Auto type';

type
  TWIASelectDialogFunc = function (AWIAManager: TWIAManager): Integer;

  TWIASettingsDialogFunc = function (AWIASource: TWIADevice;
                            var ASelectedItemIndex: Integer;
                            { #todo -oMaxM : Possibly Filters for which Items Kinds to Show? How manage AParams without Indexes? }
                            AInitItemValues: TInitialItemValues;
                            var AParams: TArrayWIAParams;
                            AOnInitDefaultValues: TInitDefaultValuesEvent=nil): Boolean;

var
   WIASelectDialogFunc: TWIASelectDialogFunc = nil;
   WIASettingsDialogFunc: TWIASettingsDialogFunc = nil;
   WIASettings_Unit_cm: Boolean = True; //False to show then measurement in fucking inches



implementation

end.

