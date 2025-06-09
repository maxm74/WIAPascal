(****************************************************************************
*
*
*  File: WiaWSDsc.pas
*
*  Version: 2.0
*
*  Description: contains custom WIA definitions for the WSD scan class driver
*
*****************************************************************************
*
*  2025-06-09 Translation and adaptation by Massimo Magnano
*
*****************************************************************************)

unit WiaWSDsc;


interface

uses WiaDef;

const
  //
  // Custom WIA property IDs (see wiadef.h)
  //
  // These custom properties describe PnP-X device properties
  // read at run time from Function Discovery, along with:
  //
  // WIA_DPS_SERVICE_ID
  // WIA_DPS_DEVICE_ID
  // WIA_DPS_GLOBAL_IDENTITY
  // WIA_DPS_FIRMWARE_VERSION
  //
  // All are read-only Root item properties maintained by the driver.
  //
  // Property Type: VT_BSTR
  // Valid Values:  WIA_PROP_NONE
  // Access Rights: READONLY
  //

  WIA_WSD_MANUFACTURER             = WIA_PRIVATE_DEVPROP;
  WIA_WSD_MANUFACTURER_STR         = 'Device manufacturer';

  WIA_WSD_MANUFACTURER_URL         = (WIA_PRIVATE_DEVPROP + 1);
  WIA_WSD_MANUFACTURER_URL_STR     = 'Manufacurer URL';

  WIA_WSD_MODEL_NAME               = (WIA_PRIVATE_DEVPROP + 2);
  WIA_WSD_MODEL_NAME_STR           = 'Model name';

  WIA_WSD_MODEL_NUMBER             = (WIA_PRIVATE_DEVPROP + 3);
  WIA_WSD_MODEL_NUMBER_STR         = 'Model number';

  WIA_WSD_MODEL_URL                = (WIA_PRIVATE_DEVPROP + 4);
  WIA_WSD_MODEL_URL_STR            = 'Model URL';

  WIA_WSD_PRESENTATION_URL         = (WIA_PRIVATE_DEVPROP + 5);
  WIA_WSD_PRESENTATION_URL_STR     = 'Presentation URL';

  WIA_WSD_FRIENDLY_NAME            = (WIA_PRIVATE_DEVPROP + 6);
  WIA_WSD_FRIENDLY_NAME_STR        = 'Friendly name';

  WIA_WSD_SERIAL_NUMBER            = (WIA_PRIVATE_DEVPROP + 7);
  WIA_WSD_SERIAL_NUMBER_STR        = 'Serial number';

  //
  // Obsolete custom WIA property for automatic input-source selection
  // during programmed push (device initiated) scanning, currently
  // replaced by the standard WIA_DPS_SCAN_AVAILABLE_ITEM (defined
  // in wiadef.h) and kept only for backwards compatibility.
  // Use WIA_DPS_SCAN_AVAILABLE_ITEM in all new code:
  //
  WIA_WSD_SCAN_AVAILABLE_ITEM      = (WIA_PRIVATE_DEVPROP + 8);
  WIA_WSD_SCAN_AVAILABLE_ITEM_STR  = WIA_DPS_SCAN_AVAILABLE_ITEM_STR;

implementation

end.

