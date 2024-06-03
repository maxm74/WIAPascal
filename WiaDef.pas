(****************************************************************************
*
*
*  File: WiaDef.pas
*
*  Version: 4.0
*
*  Description: WIA constant definitions
*
*****************************************************************************
*
*  2024-04-03 Translation and adaptation by Massimo Magnano
*
*****************************************************************************)

unit WiaDef;

{$MODE DELPHI}

interface

uses Windows, ActiveX;

type
  //MaxM: originally in "wia_lh.h" but moved here to avoid circular unit reference
  _WIA_PROPID_TO_NAME = record
      propid : PROPID;
      pszName : LPOLESTR;
  end;
  WIA_PROPID_TO_NAME = _WIA_PROPID_TO_NAME;
  PWIA_PROPID_TO_NAME = ^WIA_PROPID_TO_NAME;

{$ALIGN 8}

const
  //MaxM: originally in "sti.h" moved here to avoid use of sti unit
  // Type of device ( scanner, camera) is represented by DWORD value with
  // hi word containing generic device type , and lo word containing sub-type
  //
  StiDeviceTypeDefault          = 0;
  StiDeviceTypeScanner          = 1;
  StiDeviceTypeDigitalCamera    = 2;
  //#if (_WIN32_WINNT >= 0x0501) // Windows XP and later
  StiDeviceTypeStreamingVideo   = 3;
  //#endif //#if (_WIN32_WINNT >= 0x0501)

  //
  // WIA property ID and string constants
  //

  WIA_DIP_DEV_ID                             = 2; //= $2
  WIA_DIP_DEV_ID_STR                         = 'Unique Device ID';

  WIA_DIP_VEND_DESC                          = 3; //= $3
  WIA_DIP_VEND_DESC_STR                      = 'Manufacturer';

  WIA_DIP_DEV_DESC                           = 4; //= $4
  WIA_DIP_DEV_DESC_STR                       = 'Description';

  WIA_DIP_DEV_TYPE                           = 5; //= $5
  WIA_DIP_DEV_TYPE_STR                       = 'Type';

  WIA_DIP_PORT_NAME                          = 6; //= $6
  WIA_DIP_PORT_NAME_STR                      = 'Port';

  WIA_DIP_DEV_NAME                            = 7; //= $7
  WIA_DIP_DEV_NAME_STR                       = 'Name';

  WIA_DIP_SERVER_NAME                        = 8; //= $8
  WIA_DIP_SERVER_NAME_STR                    = 'Server';

  WIA_DIP_REMOTE_DEV_ID                       = 9; //= $9
  WIA_DIP_REMOTE_DEV_ID_STR                  = 'Remote Device ID';

  WIA_DIP_UI_CLSID                           = 10; //= $a
  WIA_DIP_UI_CLSID_STR                       = 'UI Class ID';

  WIA_DIP_HW_CONFIG                          = 11; //= $b
  WIA_DIP_HW_CONFIG_STR                      = 'Hardware Configuration';

  WIA_DIP_BAUDRATE                           = 12; //= $c
  WIA_DIP_BAUDRATE_STR                       = 'BaudRate';

  WIA_DIP_STI_GEN_CAPABILITIES               = 13; //= $d
  WIA_DIP_STI_GEN_CAPABILITIES_STR           = 'STI Generic Capabilities';

  WIA_DIP_WIA_VERSION                        = 14; //= $e
  WIA_DIP_WIA_VERSION_STR                    = 'WIA Version';

  WIA_DIP_DRIVER_VERSION                     = 15; //= $f
  WIA_DIP_DRIVER_VERSION_STR                 = 'Driver Version';

  WIA_DIP_PNP_ID                             = 16; //= $10
  WIA_DIP_PNP_ID_STR                         = 'PnP ID String';

  WIA_DIP_STI_DRIVER_VERSION                 = 17; //= $11
  WIA_DIP_STI_DRIVER_VERSION_STR             = 'STI Driver Version';

  WIA_DPA_FIRMWARE_VERSION                   = 1026; //= $402
  WIA_DPA_FIRMWARE_VERSION_STR               = 'Firmware Version';

  WIA_DPA_CONNECT_STATUS                     = 1027; //= $403
  WIA_DPA_CONNECT_STATUS_STR                 = 'Connect Status';

  WIA_DPA_DEVICE_TIME                        = 1028; //= $404
  WIA_DPA_DEVICE_TIME_STR                    = 'Device Time';

  WIA_DPC_PICTURES_TAKEN                      = 2050; //= $802
  WIA_DPC_PICTURES_TAKEN_STR                 = 'Pictures Taken';

  WIA_DPC_PICTURES_REMAINING                 = 2051; //= $803
  WIA_DPC_PICTURES_REMAINING_STR             = 'Pictures Remaining';

  WIA_DPC_EXPOSURE_MODE                      = 2052; //= $804
  WIA_DPC_EXPOSURE_MODE_STR                  = 'Exposure Mode';

  WIA_DPC_EXPOSURE_COMP                      = 2053; //= $805
  WIA_DPC_EXPOSURE_COMP_STR                  = 'Exposure Compensation';

  WIA_DPC_EXPOSURE_TIME                      = 2054; //= $806
  WIA_DPC_EXPOSURE_TIME_STR                  = 'Exposure Time';

  WIA_DPC_FNUMBER                            = 2055; //= $807
  WIA_DPC_FNUMBER_STR                        = 'F Number';

  WIA_DPC_FLASH_MODE                         = 2056; //= $808
  WIA_DPC_FLASH_MODE_STR                     = 'Flash Mode';

  WIA_DPC_FOCUS_MODE                         = 2057; //= $809
  WIA_DPC_FOCUS_MODE_STR                     = 'Focus Mode';

  WIA_DPC_FOCUS_MANUAL_DIST                  = 2058; //= $80a
  WIA_DPC_FOCUS_MANUAL_DIST_STR              = 'Focus Manual Dist';

  WIA_DPC_ZOOM_POSITION                      = 2059; //= $80b
  WIA_DPC_ZOOM_POSITION_STR                  = 'Zoom Position';

  WIA_DPC_PAN_POSITION                       = 2060; //= $80c
  WIA_DPC_PAN_POSITION_STR                   = 'Pan Position';

  WIA_DPC_TILT_POSITION                      = 2061; //= $80d
  WIA_DPC_TILT_POSITION_STR                  = 'Tilt Position';

  WIA_DPC_TIMER_MODE                         = 2062; //= $80e
  WIA_DPC_TIMER_MODE_STR                     = 'Timer Mode';

  WIA_DPC_TIMER_VALUE                        = 2063; //= $80f
  WIA_DPC_TIMER_VALUE_STR                    = 'Timer Value';

  WIA_DPC_POWER_MODE                         = 2064; //= $810
  WIA_DPC_POWER_MODE_STR                     = 'Power Mode';

  WIA_DPC_BATTERY_STATUS                     = 2065; //= $811
  WIA_DPC_BATTERY_STATUS_STR                 = 'Battery Status';

  WIA_DPC_THUMB_WIDTH                        = 2066; //= $812
  WIA_DPC_THUMB_WIDTH_STR                    = 'Thumbnail Width';

  WIA_DPC_THUMB_HEIGHT                       = 2067; //= $813
  WIA_DPC_THUMB_HEIGHT_STR                   = 'Thumbnail Height';

  WIA_DPC_PICT_WIDTH                         = 2068; //= $814
  WIA_DPC_PICT_WIDTH_STR                     = 'Picture Width';

  WIA_DPC_PICT_HEIGHT                        = 2069; //= $815
  WIA_DPC_PICT_HEIGHT_STR                    = 'Picture Height';

  WIA_DPC_DIMENSION                          = 2070; //= $816
  WIA_DPC_DIMENSION_STR                      = 'Dimension';

  WIA_DPC_COMPRESSION_SETTING                = 2071; //= $817
  WIA_DPC_COMPRESSION_SETTING_STR            = 'Compression Setting';

  WIA_DPC_FOCUS_METERING                     = 2072; //= $818
  WIA_DPC_FOCUS_METERING_STR                 = 'Focus Metering Mode';

  WIA_DPC_TIMELAPSE_INTERVAL                 = 2073; //= $819
  WIA_DPC_TIMELAPSE_INTERVAL_STR             = 'Timelapse Interval';

  WIA_DPC_TIMELAPSE_NUMBER                   = 2074; //= $81a
  WIA_DPC_TIMELAPSE_NUMBER_STR               = 'Timelapse Number';

  WIA_DPC_BURST_INTERVAL                     = 2075; //= $81b
  WIA_DPC_BURST_INTERVAL_STR                 = 'Burst Interval';

  WIA_DPC_BURST_NUMBER                       = 2076; //= $81c
  WIA_DPC_BURST_NUMBER_STR                   = 'Burst Number';

  WIA_DPC_EFFECT_MODE                        = 2077; //= $81d
  WIA_DPC_EFFECT_MODE_STR                    = 'Effect Mode';

  WIA_DPC_DIGITAL_ZOOM                       = 2078; //= $81e
  WIA_DPC_DIGITAL_ZOOM_STR                   = 'Digital Zoom';

  WIA_DPC_SHARPNESS                          = 2079; //= $81f
  WIA_DPC_SHARPNESS_STR                      = 'Sharpness';

  WIA_DPC_CONTRAST                           = 2080; //= $820
  WIA_DPC_CONTRAST_STR                       = 'Contrast';

  WIA_DPC_CAPTURE_MODE                       = 2081; //= $821
  WIA_DPC_CAPTURE_MODE_STR                   = 'Capture Mode';

  WIA_DPC_CAPTURE_DELAY                      = 2082; //= $822
  WIA_DPC_CAPTURE_DELAY_STR                  = 'Capture Delay';

  WIA_DPC_EXPOSURE_INDEX                     = 2083; //= $823
  WIA_DPC_EXPOSURE_INDEX_STR                 = 'Exposure Index';

  WIA_DPC_EXPOSURE_METERING_MODE             = 2084; //= $824
  WIA_DPC_EXPOSURE_METERING_MODE_STR         = 'Exposure Metering Mode';

  WIA_DPC_FOCUS_METERING_MODE                = 2085; //= $825
  WIA_DPC_FOCUS_METERING_MODE_STR            = 'Focus Metering Mode';

  WIA_DPC_FOCUS_DISTANCE                     = 2086; //= $826
  WIA_DPC_FOCUS_DISTANCE_STR                 = 'Focus Distance';

  WIA_DPC_FOCAL_LENGTH                       = 2087; //= $827
  WIA_DPC_FOCAL_LENGTH_STR                   = 'Focus Length';

  WIA_DPC_RGB_GAIN                           = 2088; //= $828
  WIA_DPC_RGB_GAIN_STR                       = 'RGB Gain';

  WIA_DPC_WHITE_BALANCE                      = 2089; //= $829
  WIA_DPC_WHITE_BALANCE_STR                  = 'White Balance';

  WIA_DPC_UPLOAD_URL                         = 2090; //= $82a
  WIA_DPC_UPLOAD_URL_STR                     = 'Upload URL';

  WIA_DPC_ARTIST                             = 2091; //= $82b
  WIA_DPC_ARTIST_STR                         = 'Artist';

  WIA_DPC_COPYRIGHT_INFO                     = 2092; //= $82c
  WIA_DPC_COPYRIGHT_INFO_STR                 = 'Copyright Info';

  WIA_DPS_HORIZONTAL_BED_SIZE                = 3074; //= $c02
  WIA_DPS_HORIZONTAL_BED_SIZE_STR            = 'Horizontal Bed Size';

  WIA_DPS_VERTICAL_BED_SIZE                  = 3075; //= $c03
  WIA_DPS_VERTICAL_BED_SIZE_STR              = 'Vertical Bed Size';

  WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE         = 3076; //= $c04
  WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE_STR     = 'Horizontal Sheet Feed Size';

  WIA_DPS_VERTICAL_SHEET_FEED_SIZE           = 3077; //= $c05
  WIA_DPS_VERTICAL_SHEET_FEED_SIZE_STR       = 'Vertical Sheet Feed Size';

  WIA_DPS_SHEET_FEEDER_REGISTRATION          = 3078; //= $c06
  WIA_DPS_SHEET_FEEDER_REGISTRATION_STR      = 'Sheet Feeder Registration';

  WIA_DPS_HORIZONTAL_BED_REGISTRATION        = 3079; //= $c07
  WIA_DPS_HORIZONTAL_BED_REGISTRATION_STR    = 'Horizontal Bed Registration';

  WIA_DPS_VERTICAL_BED_REGISTRATION          = 3080; //= $c08
  WIA_DPS_VERTICAL_BED_REGISTRATION_STR      = 'Vertical Bed Registration';

  WIA_DPS_PLATEN_COLOR                       = 3081; //= $c09
  WIA_DPS_PLATEN_COLOR_STR                   = 'Platen Color';

  WIA_DPS_PAD_COLOR                          = 3082; //= $c0a
  WIA_DPS_PAD_COLOR_STR                      = 'Pad Color';

  WIA_DPS_FILTER_SELECT                      = 3083; //= $c0b
  WIA_DPS_FILTER_SELECT_STR                  = 'Filter Select';

  WIA_DPS_DITHER_SELECT                      = 3084; //= $c0c
  WIA_DPS_DITHER_SELECT_STR                  = 'Dither Select';

  WIA_DPS_DITHER_PATTERN_DATA                = 3085; //= $c0d
  WIA_DPS_DITHER_PATTERN_DATA_STR            = 'Dither Pattern Data';

  WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES     = 3086; //= $c0e
  WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES_STR = 'Document Handling Capabilities';

  WIA_DPS_DOCUMENT_HANDLING_STATUS           = 3087; //= $c0f
  WIA_DPS_DOCUMENT_HANDLING_STATUS_STR       = 'Document Handling Status';

  WIA_DPS_DOCUMENT_HANDLING_SELECT           = 3088; //= $c10
  WIA_DPS_DOCUMENT_HANDLING_SELECT_STR       = 'Document Handling Select';

  WIA_DPS_DOCUMENT_HANDLING_CAPACITY         = 3089; //= $c11
  WIA_DPS_DOCUMENT_HANDLING_CAPACITY_STR     = 'Document Handling Capacity';

  WIA_DPS_OPTICAL_XRES                       = 3090; //= $c12
  WIA_DPS_OPTICAL_XRES_STR                   = 'Horizontal Optical Resolution';

  WIA_DPS_OPTICAL_YRES                       = 3091; //= $c13
  WIA_DPS_OPTICAL_YRES_STR                   = 'Vertical Optical Resolution';

  WIA_DPS_ENDORSER_CHARACTERS                = 3092; //= $c14, superseded by WIA_IPS_PRINTER_ENDORSER_VALID_CHARACTERS
  WIA_DPS_ENDORSER_CHARACTERS_STR            = 'Endorser Characters';

  WIA_DPS_ENDORSER_STRING                    = 3093; //= $c15, superseded by WIA_IPS_PRINTER_ENDORSER_STRING
  WIA_DPS_ENDORSER_STRING_STR                = 'Endorser String';

  WIA_DPS_SCAN_AHEAD_PAGES                   = 3094; //= $c16, superseded by WIA_IPS_SCAN_AHEAD
  WIA_DPS_SCAN_AHEAD_PAGES_STR               = 'Scan Ahead Pages';

  WIA_DPS_MAX_SCAN_TIME                      = 3095; //= $c17
  WIA_DPS_MAX_SCAN_TIME_STR                  = 'Max Scan Time';

  WIA_DPS_PAGES                              = 3096; //= $c18
  WIA_DPS_PAGES_STR                          = 'Pages';

  WIA_DPS_PAGE_SIZE                          = 3097; //= $c19
  WIA_DPS_PAGE_SIZE_STR                      = 'Page Size';

  WIA_DPS_PAGE_WIDTH                         = 3098; //= $c1a
  WIA_DPS_PAGE_WIDTH_STR                     = 'Page Width';

  WIA_DPS_PAGE_HEIGHT                        = 3099; //= $c1b
  WIA_DPS_PAGE_HEIGHT_STR                    = 'Page Height';

  WIA_DPS_PREVIEW                            = 3100; //= $c1c
  WIA_DPS_PREVIEW_STR                        = 'Preview';

  WIA_DPS_TRANSPARENCY                       = 3101; //= $c1d
  WIA_DPS_TRANSPARENCY_STR                   = 'Transparency Adapter';

  WIA_DPS_TRANSPARENCY_SELECT                = 3102; //= $c1e
  WIA_DPS_TRANSPARENCY_SELECT_STR            = 'Transparency Adapter Select';

  WIA_DPS_SHOW_PREVIEW_CONTROL               = 3103; //= $c1f
  WIA_DPS_SHOW_PREVIEW_CONTROL_STR           = 'Show preview control';

  WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE     = 3104; //= $c20
  WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE_STR = 'Minimum Horizontal Sheet Feed Size';

  WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE       = 3105; //= $c21
  WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE_STR   = 'Minimum Vertical Sheet Feed Size';

  WIA_DPS_TRANSPARENCY_CAPABILITIES          = 3106; //= $c22
  WIA_DPS_TRANSPARENCY_CAPABILITIES_STR      = 'Transparency Adapter Capabilities';

  WIA_DPS_TRANSPARENCY_STATUS                = 3107; //= $c23
  WIA_DPS_TRANSPARENCY_STATUS_STR            = 'Transparency Adapter Status';

  WIA_DPF_MOUNT_POINT                        = 3330; //= $d02
  WIA_DPF_MOUNT_POINT_STR                    = 'Directory mount point';

  WIA_DPV_LAST_PICTURE_TAKEN                 = 3586; //= $e02
  WIA_DPV_LAST_PICTURE_TAKEN_STR             = 'Last Picture Taken';

  WIA_DPV_IMAGES_DIRECTORY                   = 3587; //= $e03
  WIA_DPV_IMAGES_DIRECTORY_STR               = 'Images Directory';

  WIA_DPV_DSHOW_DEVICE_PATH                  = 3588; //= $e04
  WIA_DPV_DSHOW_DEVICE_PATH_STR              = 'Directshow Device Path';

  WIA_IPA_ITEM_NAME                          = 4098; //= $1002
  WIA_IPA_ITEM_NAME_STR                      = 'Item Name';

  WIA_IPA_FULL_ITEM_NAME                     = 4099; //= $1003
  WIA_IPA_FULL_ITEM_NAME_STR                 = 'Full Item Name';

  WIA_IPA_ITEM_TIME                          = 4100; //= $1004
  WIA_IPA_ITEM_TIME_STR                      = 'Item Time Stamp';

  WIA_IPA_ITEM_FLAGS                         = 4101; //= $1005
  WIA_IPA_ITEM_FLAGS_STR                     = 'Item Flags';

  WIA_IPA_ACCESS_RIGHTS                      = 4102; //= $1006
  WIA_IPA_ACCESS_RIGHTS_STR                  = 'Access Rights';

  WIA_IPA_DATATYPE                           = 4103; //= $1007
  WIA_IPA_DATATYPE_STR                       = 'Data Type';

  WIA_IPA_DEPTH                              = 4104; //= $1008
  WIA_IPA_DEPTH_STR                          = 'Bits Per Pixel';

  WIA_IPA_PREFERRED_FORMAT                   = 4105; //= $1009
  WIA_IPA_PREFERRED_FORMAT_STR               = 'Preferred Format';

  WIA_IPA_FORMAT                             = 4106; //= $100a
  WIA_IPA_FORMAT_STR                         = 'Format';

  WIA_IPA_COMPRESSION                        = 4107; //= $100b
  WIA_IPA_COMPRESSION_STR                    = 'Compression';

  WIA_IPA_TYMED                              = 4108; //= $100c
  WIA_IPA_TYMED_STR                          = 'Media Type';

  WIA_IPA_CHANNELS_PER_PIXEL                 = 4109; //= $100d
  WIA_IPA_CHANNELS_PER_PIXEL_STR             = 'Channels Per Pixel';

  WIA_IPA_BITS_PER_CHANNEL                   = 4110; //= $100e
  WIA_IPA_BITS_PER_CHANNEL_STR               = 'Bits Per Channel';

  WIA_IPA_PLANAR                             = 4111; //= $100f
  WIA_IPA_PLANAR_STR                         = 'Planar';

  WIA_IPA_PIXELS_PER_LINE                    = 4112; //= $1010
  WIA_IPA_PIXELS_PER_LINE_STR                = 'Pixels Per Line';

  WIA_IPA_BYTES_PER_LINE                     = 4113; //= $1011
  WIA_IPA_BYTES_PER_LINE_STR                 = 'Bytes Per Line';

  WIA_IPA_NUMBER_OF_LINES                    = 4114; //= $1012
  WIA_IPA_NUMBER_OF_LINES_STR                = 'Number of Lines';

  WIA_IPA_GAMMA_CURVES                       = 4115; //= $1013
  WIA_IPA_GAMMA_CURVES_STR                   = 'Gamma Curves';

  WIA_IPA_ITEM_SIZE                          = 4116; //= $1014
  WIA_IPA_ITEM_SIZE_STR                      = 'Item Size';

  WIA_IPA_COLOR_PROFILE                      = 4117; //= $1015
  WIA_IPA_COLOR_PROFILE_STR                  = 'Color Profiles';

  WIA_IPA_MIN_BUFFER_SIZE                    = 4118; //= $1016
  WIA_IPA_MIN_BUFFER_SIZE_STR                = 'Buffer Size';

  WIA_IPA_BUFFER_SIZE                        = 4118; //= $1016
  WIA_IPA_BUFFER_SIZE_STR                    = 'Buffer Size';

  WIA_IPA_REGION_TYPE                        = 4119; //= $1017
  WIA_IPA_REGION_TYPE_STR                    = 'Region Type';

  WIA_IPA_ICM_PROFILE_NAME                   = 4120; //= $1018
  WIA_IPA_ICM_PROFILE_NAME_STR               = 'Color Profile Name';

  WIA_IPA_APP_COLOR_MAPPING                  = 4121; //= $1019
  WIA_IPA_APP_COLOR_MAPPING_STR              = 'Application Applies Color Mapping';

  WIA_IPA_PROP_STREAM_COMPAT_ID              = 4122; //= $101a
  WIA_IPA_PROP_STREAM_COMPAT_ID_STR          = 'Stream Compatibility ID';

  WIA_IPA_FILENAME_EXTENSION                 = 4123; //= $101b
  WIA_IPA_FILENAME_EXTENSION_STR             = 'Filename extension';

  WIA_IPA_SUPPRESS_PROPERTY_PAGE             = 4124; //= $101c
  WIA_IPA_SUPPRESS_PROPERTY_PAGE_STR         = 'Suppress a property page';

  WIA_IPC_THUMBNAIL                          = 5122; //= $1402
  WIA_IPC_THUMBNAIL_STR                      = 'Thumbnail Data';

  WIA_IPC_THUMB_WIDTH                        = 5123; //= $1403
  WIA_IPC_THUMB_WIDTH_STR                    = 'Thumbnail Width';

  WIA_IPC_THUMB_HEIGHT                       = 5124; //= $1404
  WIA_IPC_THUMB_HEIGHT_STR                   = 'Thumbnail Height';

  WIA_IPC_AUDIO_AVAILABLE                    = 5125; //= $1405
  WIA_IPC_AUDIO_AVAILABLE_STR                = 'Audio Available';

  WIA_IPC_AUDIO_DATA_FORMAT                  = 5126; //= $1406
  WIA_IPC_AUDIO_DATA_FORMAT_STR              = 'Audio Format';

  WIA_IPC_AUDIO_DATA                         = 5127; //= $1407
  WIA_IPC_AUDIO_DATA_STR                     = 'Audio Data';

  WIA_IPC_NUM_PICT_PER_ROW                   = 5128; //= $1408
  WIA_IPC_NUM_PICT_PER_ROW_STR               = 'Pictures per Row';

  WIA_IPC_SEQUENCE                           = 5129; //= $1409
  WIA_IPC_SEQUENCE_STR                       = 'Sequence Number';

  WIA_IPC_TIMEDELAY                          = 5130; //= $140a
  WIA_IPC_TIMEDELAY_STR                      = 'Time Delay';

  WIA_IPS_CUR_INTENT                         = 6146; //= $1802
  WIA_IPS_CUR_INTENT_STR                     = 'Current Intent';

  WIA_IPS_XRES                               = 6147; //= $1803
  WIA_IPS_XRES_STR                           = 'Horizontal Resolution';

  WIA_IPS_YRES                               = 6148; //= $1804
  WIA_IPS_YRES_STR                           = 'Vertical Resolution';

  WIA_IPS_XPOS                               = 6149; //= $1805
  WIA_IPS_XPOS_STR                           = 'Horizontal Start Position';

  WIA_IPS_YPOS                               = 6150; //= $1806
  WIA_IPS_YPOS_STR                           = 'Vertical Start Position';

  WIA_IPS_XEXTENT                            = 6151; //= $1807
  WIA_IPS_XEXTENT_STR                        = 'Horizontal Extent';

  WIA_IPS_YEXTENT                            = 6152; //= $1808
  WIA_IPS_YEXTENT_STR                        = 'Vertical Extent';

  WIA_IPS_PHOTOMETRIC_INTERP                 = 6153; //= $1809
  WIA_IPS_PHOTOMETRIC_INTERP_STR             = 'Photometric Interpretation';

  WIA_IPS_BRIGHTNESS                         = 6154; //= $180a
  WIA_IPS_BRIGHTNESS_STR                     = 'Brightness';

  WIA_IPS_CONTRAST                           = 6155; //= $180b
  WIA_IPS_CONTRAST_STR                       = 'Contrast';

  WIA_IPS_ORIENTATION                        = 6156; //= $180c
  WIA_IPS_ORIENTATION_STR                    = 'Orientation';

  WIA_IPS_ROTATION                           = 6157; //= $180d
  WIA_IPS_ROTATION_STR                       = 'Rotation';

  WIA_IPS_MIRROR                             = 6158; //= $180e
  WIA_IPS_MIRROR_STR                         = 'Mirror';

  WIA_IPS_THRESHOLD                          = 6159; //= $180f
  WIA_IPS_THRESHOLD_STR                      = 'Threshold';

  WIA_IPS_INVERT                             = 6160; //= $1810
  WIA_IPS_INVERT_STR                         = 'Invert';

  WIA_IPS_WARM_UP_TIME                       = 6161; //= $1811
  WIA_IPS_WARM_UP_TIME_STR                   = 'Lamp Warm up Time';

  //{$ifdef win32vista} //{$if _WIN32_WINNT>=$0600}
  //
  // New properties, property names and values specific to WIA 2.0 (introduced in Windows Vista):
  //

  WIA_DPS_USER_NAME                          = 3112; //= $c28
  WIA_DPS_USER_NAME_STR                      = 'User Name';

  WIA_DPS_SERVICE_ID                         = 3113; //= $c29
  WIA_DPS_SERVICE_ID_STR                     = 'Service ID';

  WIA_DPS_DEVICE_ID                          = 3114; //= $c2a
  WIA_DPS_DEVICE_ID_STR                      = 'Device ID';

  WIA_DPS_GLOBAL_IDENTITY                    = 3115; //= $c2b
  WIA_DPS_GLOBAL_IDENTITY_STR                = 'Global Identity';

  WIA_DPS_SCAN_AVAILABLE_ITEM                = 3116; //= $c2c
  WIA_DPS_SCAN_AVAILABLE_ITEM_STR            = 'Scan Available Item';

  WIA_IPS_DESKEW_X                           = 6162; //= $1812
  WIA_IPS_DESKEW_X_STR                       = 'DeskewX';

  WIA_IPS_DESKEW_Y                           = 6163; //= $1813
  WIA_IPS_DESKEW_Y_STR                       = 'DeskewY';

  WIA_IPS_SEGMENTATION                       = 6164; //= $1814
  WIA_IPS_SEGMENTATION_STR                   = 'Segmentation';

  WIA_SEGMENTATION_FILTER_STR                = 'SegmentationFilter';
  WIA_IMAGEPROC_FILTER_STR                   = 'ImageProcessingFilter';

  WIA_IPS_MAX_HORIZONTAL_SIZE                = 6165; //= $1815
  WIA_IPS_MAX_HORIZONTAL_SIZE_STR            = 'Maximum Horizontal Scan Size';

  WIA_IPS_MAX_VERTICAL_SIZE                  = 6166; //= $1816
  WIA_IPS_MAX_VERTICAL_SIZE_STR              = 'Maximum Vertical Scan Size';

  WIA_IPS_MIN_HORIZONTAL_SIZE                = 6167; //= $1817
  WIA_IPS_MIN_HORIZONTAL_SIZE_STR            = 'Minimum Horizontal Scan Size';

  WIA_IPS_MIN_VERTICAL_SIZE                  = 6168; //= $1818
  WIA_IPS_MIN_VERTICAL_SIZE_STR              = 'Minimum Vertical Scan Size';

  WIA_IPS_TRANSFER_CAPABILITIES              = 6169; //= $1819
  WIA_IPS_TRANSFER_CAPABILITIES_STR          = 'Transfer Capabilities';

  WIA_IPS_SHEET_FEEDER_REGISTRATION          = 3078; //= $c06
  WIA_IPS_SHEET_FEEDER_REGISTRATION_STR      = 'Sheet Feeder Registration';

  WIA_IPS_DOCUMENT_HANDLING_SELECT           = 3088; //= $c10
  WIA_IPS_DOCUMENT_HANDLING_SELECT_STR       = 'Document Handling Select';

  WIA_IPS_OPTICAL_XRES                       = 3090; //= $c12
  WIA_IPS_OPTICAL_XRES_STR                   = 'Horizontal Optical Resolution';

  WIA_IPS_OPTICAL_YRES                       = 3091; //= $c13
  WIA_IPS_OPTICAL_YRES_STR                   = 'Vertical Optical Resolution';

  WIA_IPS_PAGES                              = 3096; //= $c18
  WIA_IPS_PAGES_STR                          = 'Pages';

  WIA_IPS_PAGE_SIZE                          = 3097; //= $c19
  WIA_IPS_PAGE_SIZE_STR                      = 'Page Size';

  WIA_IPS_PAGE_WIDTH                         = 3098; //= $c1a
  WIA_IPS_PAGE_WIDTH_STR                     = 'Page Width';

  WIA_IPS_PAGE_HEIGHT                        = 3099; //= $c1b
  WIA_IPS_PAGE_HEIGHT_STR                    = 'Page Height';

  WIA_IPS_PREVIEW                            = 3100; //= $c1c
  WIA_IPS_PREVIEW_STR                        = 'Preview';

  WIA_IPS_SHOW_PREVIEW_CONTROL               = 3103; //= $c1f
  WIA_IPS_SHOW_PREVIEW_CONTROL_STR           = 'Show preview control';

  WIA_IPS_FILM_SCAN_MODE                     = 3104; //= $c20
  WIA_IPS_FILM_SCAN_MODE_STR                 = 'Film Scan Mode';

  WIA_IPS_LAMP                               = 3105; //= $c21
  WIA_IPS_LAMP_STR                           = 'Lamp';

  WIA_IPS_LAMP_AUTO_OFF                      = 3106; //= $c22
  WIA_IPS_LAMP_AUTO_OFF_STR                  = 'Lamp Auto Off';

  WIA_IPS_AUTO_DESKEW                        = 3107; //= $c23
  WIA_IPS_AUTO_DESKEW_STR                    = 'Automatic Deskew';

  WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION       = 3108; //= $c24
  WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION_STR   = 'Supports Child Item Creation';

  WIA_IPS_XSCALING                           = 3109; //= $c25
  WIA_IPS_XSCALING_STR                       = 'Horizontal Scaling';

  WIA_IPS_YSCALING                           = 3110; //= $c26
  WIA_IPS_YSCALING_STR                       = 'Vertical Scaling';

  WIA_IPS_PREVIEW_TYPE                       = 3111; //= $c27
  WIA_IPS_PREVIEW_TYPE_STR                   = 'Preview Type';

  WIA_IPA_ITEM_CATEGORY                      = 4125; //= $101d
  WIA_IPA_ITEM_CATEGORY_STR                  = 'Item Category';

  WIA_IPA_UPLOAD_ITEM_SIZE                   = 4126; //= $101e
  WIA_IPA_UPLOAD_ITEM_SIZE_STR               = 'Upload Item Size';

  WIA_IPA_ITEMS_STORED                       = 4127; //= $101f
  WIA_IPA_ITEMS_STORED_STR                   = 'Items Stored';

  WIA_IPA_RAW_BITS_PER_CHANNEL               = 4128; //= $1020
  WIA_IPA_RAW_BITS_PER_CHANNEL_STR           = 'Raw Bits Per Channel';

  WIA_IPS_FILM_NODE_NAME                     = 4129; //= $1021
  WIA_IPS_FILM_NODE_NAME_STR                 = 'Film Node Name';

  WIA_IPS_PRINTER_ENDORSER                   = 4130; //= $1022
  WIA_IPS_PRINTER_ENDORSER_STR               = 'Printer/Endorser';

  WIA_IPS_PRINTER_ENDORSER_ORDER             = 4131; //= $1023
  WIA_IPS_PRINTER_ENDORSER_ORDER_STR         = 'Printer/Endorser Order';

  WIA_IPS_PRINTER_ENDORSER_COUNTER           = 4132; //= $1024
  WIA_IPS_PRINTER_ENDORSER_COUNTER_STR       = 'Printer/Endorser Counter';

  WIA_IPS_PRINTER_ENDORSER_STEP              = 4133; //= $1025
  WIA_IPS_PRINTER_ENDORSER_STEP_STR          = 'Printer/Endorser Step';

  WIA_IPS_PRINTER_ENDORSER_XOFFSET           = 4134; //= $1026
  WIA_IPS_PRINTER_ENDORSER_XOFFSET_STR       = 'Printer/Endorser Horizontal Offset';

  WIA_IPS_PRINTER_ENDORSER_YOFFSET           = 4135; //= $1027
  WIA_IPS_PRINTER_ENDORSER_YOFFSET_STR       = 'Printer/Endorser Vertical Offset';

  WIA_IPS_PRINTER_ENDORSER_NUM_LINES         = 4136; //= $1028
  WIA_IPS_PRINTER_ENDORSER_NUM_LINES_STR     = 'Printer/Endorser Lines';

  WIA_IPS_PRINTER_ENDORSER_STRING            = 4137; //= $1029
  WIA_IPS_PRINTER_ENDORSER_STRING_STR        = 'Printer/Endorser String';

  WIA_IPS_PRINTER_ENDORSER_VALID_CHARACTERS  = 4138; //= $102a
  WIA_IPS_PRINTER_ENDORSER_VALID_CHARACTERS_STR        = 'Printer/Endorser Valid Characters';

  WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS = 4139; //= $102b
  WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS_STR = 'Printer/Endorser Valid Format Specifiers';

  WIA_IPS_PRINTER_ENDORSER_TEXT_UPLOAD       = 4140; //= $102c
  WIA_IPS_PRINTER_ENDORSER_TEXT_UPLOAD_STR   = 'Printer/Endorser Text Upload';

  WIA_IPS_PRINTER_ENDORSER_TEXT_DOWNLOAD     = 4141; //= $102d
  WIA_IPS_PRINTER_ENDORSER_TEXT_DOWNLOAD_STR = 'Printer/Endorser Text Download';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS          = 4142; //= $102e
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_STR      = 'Printer/Endorser Graphics';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_POSITION = 4143; //= $102f
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_POSITION_STR       = 'Printer/Endorser Graphics Position';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_WIDTH = 4144; //= $1030
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_WIDTH_STR      = 'Printer/Endorser Graphics Minimum Width';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_WIDTH = 4145; //= $1031
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_WIDTH_STR      = 'Printer/Endorser Graphics Maximum Width';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_HEIGHT = 4146; //= $1032
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_HEIGHT_STR     = 'Printer/Endorser Graphics Minimum Height';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_HEIGHT = 4147; //= $1033
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_HEIGHT_STR     = 'Printer/Endorser Graphics Maximum Height';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_UPLOAD   = 4148; //= $1034
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_UPLOAD_STR         = 'Printer/Endorser Graphics Upload';

  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_DOWNLOAD = 4149; //= $1035
  WIA_IPS_PRINTER_ENDORSER_GRAPHICS_DOWNLOAD_STR       = 'Printer/Endorser Graphics Download';

  WIA_IPS_BARCODE_READER                     = 4150; //= $1036
  WIA_IPS_BARCODE_READER_STR                 = 'Barcode Reader';

  WIA_IPS_MAXIMUM_BARCODES_PER_PAGE          = 4151; //= $1037
  WIA_IPS_MAXIMUM_BARCODES_PER_PAGE_STR      = 'Maximum Barcodes Per Page';

  WIA_IPS_BARCODE_SEARCH_DIRECTION           = 4152; //= $1038
  WIA_IPS_BARCODE_SEARCH_DIRECTION_STR       = 'Barcode Search Direction';

  WIA_IPS_MAXIMUM_BARCODE_SEARCH_RETRIES     = 4153; //= $1039
  WIA_IPS_MAXIMUM_BARCODE_SEARCH_RETRIES_STR = 'Barcode Search Retries';

  WIA_IPS_BARCODE_SEARCH_TIMEOUT             = 4154; //= $103a
  WIA_IPS_BARCODE_SEARCH_TIMEOUT_STR         = 'Barcode Search Timeout';

  WIA_IPS_SUPPORTED_BARCODE_TYPES            = 4155; //= $103b
  WIA_IPS_SUPPORTED_BARCODE_TYPES_STR        = 'Supported Barcode Types';

  WIA_IPS_ENABLED_BARCODE_TYPES              = 4156; //= $103c
  WIA_IPS_ENABLED_BARCODE_TYPES_STR          = 'Enabled Barcode Types';

  WIA_IPS_PATCH_CODE_READER                  = 4157; //= $103d
  WIA_IPS_PATCH_CODE_READER_STR              = 'Patch Code Reader';

  WIA_IPS_SUPPORTED_PATCH_CODE_TYPES         = 4162; //= $1042
  WIA_IPS_SUPPORTED_PATCH_CODE_TYPES_STR     = 'Supported Patch Code Types';

  WIA_IPS_ENABLED_PATCH_CODE_TYPES           = 4163; //= $1043
  WIA_IPS_ENABLED_PATCH_CODE_TYPES_STR       = 'Enabled Path Code Types';

  WIA_IPS_MICR_READER                        = 4164; //= $1044
  WIA_IPS_MICR_READER_STR                    = 'MICR Reader';

  WIA_IPS_JOB_SEPARATORS                     = 4165; //= $1045
  WIA_IPS_JOB_SEPARATORS_STR                 = 'Job Separators';

  WIA_IPS_LONG_DOCUMENT                      = 4166; //= $1046
  WIA_IPS_LONG_DOCUMENT_STR                  = 'Long Document';

  WIA_IPS_BLANK_PAGES                        = 4167; //= $1047
  WIA_IPS_BLANK_PAGES_STR                    = 'Blank Pages';

  WIA_IPS_MULTI_FEED                         = 4168; //= $1048
  WIA_IPS_MULTI_FEED_STR                     = 'Multi-Feed';

  WIA_IPS_MULTI_FEED_SENSITIVITY             = 4169; //= $1049
  WIA_IPS_MULTI_FEED_SENSITIVITY_STR         = 'Multi-Feed Sensitivity';

  WIA_IPS_AUTO_CROP                          = 4170; //= $104a
  WIA_IPS_AUTO_CROP_STR                      = 'Auto-Crop';

  WIA_IPS_OVER_SCAN                          = 4171; //= $104b
  WIA_IPS_OVER_SCAN_STR                      = 'Overscan';

  WIA_IPS_OVER_SCAN_LEFT                     = 4172; //= $104c
  WIA_IPS_OVER_SCAN_LEFT_STR                 = 'Overscan Left';

  WIA_IPS_OVER_SCAN_RIGHT                    = 4173; //= $104d
  WIA_IPS_OVER_SCAN_RIGHT_STR                = 'Overscan Right';

  WIA_IPS_OVER_SCAN_TOP                      = 4174; //= $104e
  WIA_IPS_OVER_SCAN_TOP_STR                  = 'Overscan Top';

  WIA_IPS_OVER_SCAN_BOTTOM                   = 4175; //= $104f
  WIA_IPS_OVER_SCAN_BOTTOM_STR               = 'Overscan Bottom';

  WIA_IPS_COLOR_DROP                         = 4176; //= $1050
  WIA_IPS_COLOR_DROP_STR                     = 'Color Drop';

  WIA_IPS_COLOR_DROP_RED                     = 4177; //= $1051
  WIA_IPS_COLOR_DROP_RED_STR                 = 'Color Drop Red';

  WIA_IPS_COLOR_DROP_GREEN                   = 4178; //= $1052
  WIA_IPS_COLOR_DROP_GREEN_STR               = 'Color Drop Green';

  WIA_IPS_COLOR_DROP_BLUE                    = 4179; //= $1053
  WIA_IPS_COLOR_DROP_BLUE_STR                = 'Color Drop Blue';

  WIA_IPS_SCAN_AHEAD                         = 4180; //= $1054
  WIA_IPS_SCAN_AHEAD_STR                     = 'Scan Ahead';

  WIA_IPS_SCAN_AHEAD_CAPACITY                = 4181; //= $1055
  WIA_IPS_SCAN_AHEAD_CAPACITY_STR            = 'Scan Ahead Capacity';

  WIA_IPS_FEEDER_CONTROL                     = 4182; //= $1056
  WIA_IPS_FEEDER_CONTROL_STR                 = 'Feeder Control';

  WIA_IPS_PRINTER_ENDORSER_PADDING           = 4183; //= $1057
  WIA_IPS_PRINTER_ENDORSER_PADDING_STR       = 'Printer/Endorser Padding';

  WIA_IPS_PRINTER_ENDORSER_FONT_TYPE         = 4184; //= $1058
  WIA_IPS_PRINTER_ENDORSER_FONT_TYPE_STR     = 'Printer/Endorser Font Type';

  WIA_IPS_ALARM                              = 4185; //= $1059
  WIA_IPS_ALARM_STR                          = 'Alarm';

  WIA_IPS_PRINTER_ENDORSER_INK               = 4186; //= $105A
  WIA_IPS_PRINTER_ENDORSER_INK_STR           = 'Printer/Endorser Ink';

  WIA_IPS_PRINTER_ENDORSER_CHARACTER_ROTATION= 4187; //= $105B
  WIA_IPS_PRINTER_ENDORSER_CHARACTER_ROTATION_STR      = 'Printer/Endorser Character Rotation';

  WIA_IPS_PRINTER_ENDORSER_MAX_CHARACTERS    = 4188; //= $105C
  WIA_IPS_PRINTER_ENDORSER_MAX_CHARACTERS_STR= 'Printer/Endorser Maximum Characters';

  WIA_IPS_PRINTER_ENDORSER_MAX_GRAPHICS      = 4189; //= $105D
  WIA_IPS_PRINTER_ENDORSER_MAX_GRAPHICS_STR  = 'Printer/Endorser Maximum Graphics';

  WIA_IPS_PRINTER_ENDORSER_COUNTER_DIGITS    = 4190; //= $105E
  WIA_IPS_PRINTER_ENDORSER_COUNTER_DIGITS_STR= 'Printer/Endorser Counter Digits';

  WIA_IPS_COLOR_DROP_MULTI                   = 4191; //= $105F
  WIA_IPS_COLOR_DROP_MULTI_STR               = 'Color Drop Multiple';

  WIA_IPS_BLANK_PAGES_SENSITIVITY            = 4192; //= $1060
  WIA_IPS_BLANK_PAGES_SENSITIVITY_STR        = 'Blank Pages Sensitivity';

  WIA_IPS_MULTI_FEED_DETECT_METHOD           = 4193; //= $1061
  WIA_IPS_MULTI_FEED_DETECT_METHOD_STR       = 'Multi-Feed Detection Method';

  //
  // WIA_IPA_ITEM_CATEGORY constants
  //

  WIA_CATEGORY_FINISHED_FILE: TGUID = '{ff2b77ca-cf84-432b-a735-3a130dde2a88}';
  WIA_CATEGORY_FLATBED: TGUID = '{fb607b1f-43f3-488b-855b-fb703ec342a6}';
  WIA_CATEGORY_FEEDER: TGUID = '{fe131934-f84c-42ad-8da4-6129cddd7288}';
  WIA_CATEGORY_FILM: TGUID = '{fcf65be7-3ce3-4473-af85-f5d37d21b68a}';
  WIA_CATEGORY_ROOT: TGUID = '{f193526f-59b8-4a26-9888-e16e4f97ce10}';
  WIA_CATEGORY_FOLDER: TGUID = '{c692a446-6f5a-481d-85bb-92e2e86fd30a}';
  WIA_CATEGORY_FEEDER_FRONT: TGUID = '{4823175c-3b28-487b-a7e6-eebc17614fd1}';
  WIA_CATEGORY_FEEDER_BACK: TGUID = '{61ca74d4-39db-42aa-89b1-8c19c9cd4c23}';
  WIA_CATEGORY_AUTO: TGUID = '{defe5fd8-6c97-4dde-b11e-cb509b270e11}';
  WIA_CATEGORY_IMPRINTER: TGUID = '{fc65016d-9202-43dd-91a7-64c2954cfb8b}';
  WIA_CATEGORY_ENDORSER: TGUID = '{47102cc3-127f-4771-adfc-991ab8ee1e97}';
  WIA_CATEGORY_BARCODE_READER: TGUID = '{36e178a0-473f-494b-af8f-6c3f6d7486fc}';
  WIA_CATEGORY_PATCH_CODE_READER: TGUID = '{8faa1a6d-9c8a-42cd-98b3-ee9700cbc74f}';
  WIA_CATEGORY_MICR_READER: TGUID = '{3b86c1ec-71bc-4645-b4d5-1b19da2be978}';

  //
  // Default Segmentation Filter GUID
  //

  CLSID_WiaDefaultSegFilter: TGUID = '{D4F4D30B-0B29-4508-8922-0C5797D42765}';

  //
  // WIA_IPS_TRANSFER_CAPABILITIES flags:
  //

  WIA_TRANSFER_CHILDREN_SINGLE_SCAN    = $00000001;

  //
  // WIA_IPS_SEGMENTATION_FILTER constants
  //

  WIA_USE_SEGMENTATION_FILTER          = 0;
  WIA_DONT_USE_SEGMENTATION_FILTER     = 1;

  //
  // WIA_IPS_FILM_SCAN_MODE constants
  //

  WIA_FILM_COLOR_SLIDE                 = 0;
  WIA_FILM_COLOR_NEGATIVE              = 1;
  WIA_FILM_BW_NEGATIVE                 = 2;

  //
  // WIA_IPS_LAMP constants
  //

  WIA_LAMP_ON                          = 0;
  WIA_LAMP_OFF                         = 1;

  //
  // WIA_IPS_AUTO_DESKEW constants
  //

  WIA_AUTO_DESKEW_ON                   = 0;
  WIA_AUTO_DESKEW_OFF                  = 1;

  //
  // WIA_IPS_PREVIEW_TYPE constants
  //

  WIA_ADVANCED_PREVIEW                 = 0;
  WIA_BASIC_PREVIEW                    = 1;

  //
  // WIA_IPS_PRINTER_ENDORSER constants
  //

  WIA_PRINTER_ENDORSER_DISABLED       = 0;
  WIA_PRINTER_ENDORSER_AUTO           = 1;
  WIA_PRINTER_ENDORSER_FLATBED        = 2;
  WIA_PRINTER_ENDORSER_FEEDER_FRONT   = 3;
  WIA_PRINTER_ENDORSER_FEEDER_BACK    = 4;
  WIA_PRINTER_ENDORSER_FEEDER_DUPLEX  = 5;
  WIA_PRINTER_ENDORSER_DIGITAL        = 6;

  //
  // WIA_IPS_PRINTER_ENDORSER_ORDER constants
  //

  WIA_PRINTER_ENDORSER_BEFORE_SCAN    = 0;
  WIA_PRINTER_ENDORSER_AFTER_SCAN     = 1;

  //
  // WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS constants
  //

  WIA_PRINT_DATE                      = 0; //= '$DATE$';
  WIA_PRINT_YEAR                      = 1; //= '$YEAR$';
  WIA_PRINT_MONTH                     = 2; //= '$MONTH$';
  WIA_PRINT_DAY                       = 3; //= '$DAY$';
  WIA_PRINT_WEEK_DAY                  = 4; //= '$WEEK_DAY$';
  WIA_PRINT_TIME_24H                  = 5; //= '$TIME$';
  WIA_PRINT_TIME_12H                  = 6; //= '$TIME_12H$';
  WIA_PRINT_HOUR_24H                  = 7; //= '$HOUR_24H$';
  WIA_PRINT_HOUR_12H                  = 8; //= '$HOUR_12H$';
  WIA_PRINT_AM_PM                     = 9; //= '$AM_PM$';
  WIA_PRINT_MINUTE                    = 10; //= '$MINUTE$';
  WIA_PRINT_SECOND                    = 11; //= '$SECOND$';
  WIA_PRINT_PAGE_COUNT                = 12; //= '$PAGE_COUNT$';
  WIA_PRINT_IMAGE                     = 13; //= '$IMAGE$';
  WIA_PRINT_MILLISECOND               = 14; //= '$MSECOND$';
  WIA_PRINT_MONTH_NAME                = 15; //= '$MONTH_NAME$';
  WIA_PRINT_MONTH_SHORT               = 16; //= '$MONTH_SHORT$';
  WIA_PRINT_WEEK_DAY_SHORT            = 17; //= '$WEEK_DAY_SHORT$';

  //
  // WIA_IPS_PRINTER_ENDORSER_GRAPHICS_POSITION constants
  //

  WIA_PRINTER_ENDORSER_GRAPHICS_LEFT   = 0;
  WIA_PRINTER_ENDORSER_GRAPHICS_RIGHT  = 1;
  WIA_PRINTER_ENDORSER_GRAPHICS_TOP    = 2;
  WIA_PRINTER_ENDORSER_GRAPHICS_BOTTOM = 3;
  WIA_PRINTER_ENDORSER_GRAPHICS_TOP_LEFT = 4;
  WIA_PRINTER_ENDORSER_GRAPHICS_TOP_RIGHT = 5;
  WIA_PRINTER_ENDORSER_GRAPHICS_BOTTOM_LEFT = 6;
  WIA_PRINTER_ENDORSER_GRAPHICS_BOTTOM_RIGHT = 7;
  WIA_PRINTER_ENDORSER_GRAPHICS_BACKGROUND = 8;
  WIA_PRINTER_ENDORSER_GRAPHICS_DEVICE_DEFAULT = 9;

  //
  // WIA_IPS_BARCODE_READER constants
  //

  WIA_BARCODE_READER_DISABLED          = 0;
  WIA_BARCODE_READER_AUTO              = 1;
  WIA_BARCODE_READER_FLATBED           = 2;
  WIA_BARCODE_READER_FEEDER_FRONT      = 3;
  WIA_BARCODE_READER_FEEDER_BACK       = 4;
  WIA_BARCODE_READER_FEEDER_DUPLEX     = 5;

  //
  // The WIA_IPS_BARCODE_SEARCH_DIRECTION constants
  //

  WIA_BARCODE_HORIZONTAL_SEARCH          = 0;
  WIA_BARCODE_VERTICAL_SEARCH            = 1;
  WIA_BARCODE_HORIZONTAL_VERTICAL_SEARCH = 2;
  WIA_BARCODE_VERTICAL_HORIZONTAL_SEARCH = 3;
  WIA_BARCODE_AUTO_SEARCH                = 4;

  //
  // WIA_IPS_SUPPORTED_BARCODE_TYPES constants
  //

  WIA_BARCODE_UPCA                = 0;
  WIA_BARCODE_UPCE                = 1;
  WIA_BARCODE_CODABAR             = 2;
  WIA_BARCODE_NONINTERLEAVED_2OF5 = 3;
  WIA_BARCODE_INTERLEAVED_2OF5    = 4;
  WIA_BARCODE_CODE39              = 5;
  WIA_BARCODE_CODE39_MOD43        = 6;
  WIA_BARCODE_CODE39_FULLASCII    = 7;
  WIA_BARCODE_CODE93              = 8;
  WIA_BARCODE_CODE128             = 9;
  WIA_BARCODE_CODE128A            = 10;
  WIA_BARCODE_CODE128B            = 11;
  WIA_BARCODE_CODE128C            = 12;
  WIA_BARCODE_GS1128              = 13;
  WIA_BARCODE_GS1DATABAR          = 14;
  WIA_BARCODE_ITF14               = 15;
  WIA_BARCODE_EAN8                = 16;
  WIA_BARCODE_EAN13               = 17;
  WIA_BARCODE_POSTNETA            = 18;
  WIA_BARCODE_POSTNETB            = 19;
  WIA_BARCODE_POSTNETC            = 20;
  WIA_BARCODE_POSTNET_DPBC        = 21;
  WIA_BARCODE_PLANET              = 22;
  WIA_BARCODE_INTELLIGENT_MAIL    = 23;
  WIA_BARCODE_POSTBAR             = 24;
  WIA_BARCODE_RM4SCC              = 25;
  WIA_BARCODE_HIGH_CAPACITY_COLOR = 26;
  WIA_BARCODE_MAXICODE            = 27;
  WIA_BARCODE_PDF417              = 28;
  WIA_BARCODE_CPCBINARY           = 29;
  WIA_BARCODE_FIM                 = 30;
  WIA_BARCODE_PHARMACODE          = 31;
  WIA_BARCODE_PLESSEY             = 32;
  WIA_BARCODE_MSI                 = 33;
  WIA_BARCODE_JAN                 = 34;
  WIA_BARCODE_TELEPEN             = 35;
  WIA_BARCODE_AZTEC               = 36;
  WIA_BARCODE_SMALLAZTEC          = 37;
  WIA_BARCODE_DATAMATRIX          = 38;
  WIA_BARCODE_DATASTRIP           = 39;
  WIA_BARCODE_EZCODE              = 40;
  WIA_BARCODE_QRCODE              = 41;
  WIA_BARCODE_SHOTCODE            = 42;
  WIA_BARCODE_SPARQCODE           = 43;
  WIA_BARCODE_CUSTOMBASE          = $8000;

  //
  // WIA_IPS_PATH_CODE_READER constants
  //

  WIA_PATCH_CODE_READER_DISABLED       = 0;
  WIA_PATCH_CODE_READER_AUTO           = 1;
  WIA_PATCH_CODE_READER_FLATBED        = 2;
  WIA_PATCH_CODE_READER_FEEDER_FRONT   = 3;
  WIA_PATCH_CODE_READER_FEEDER_BACK    = 4;
  WIA_PATCH_CODE_READER_FEEDER_DUPLEX  = 5;

  //
  // WIA_IPS_SUPPORTED_PATCH_CODE_TYPES constants
  //

  WIA_PATCH_CODE_UNKNOWN     = 0;
  WIA_PATCH_CODE_1           = 1;
  WIA_PATCH_CODE_2           = 2;
  WIA_PATCH_CODE_3           = 3;
  WIA_PATCH_CODE_4           = 4;
  WIA_PATCH_CODE_T           = 5;
  WIA_PATCH_CODE_6           = 6;
  WIA_PATCH_CODE_7           = 7;
  WIA_PATCH_CODE_8           = 8;
  WIA_PATCH_CODE_9           = 9;
  WIA_PATCH_CODE_10          = 10;
  WIA_PATCH_CODE_11          = 11;
  WIA_PATCH_CODE_12          = 12;
  WIA_PATCH_CODE_13          = 13;
  WIA_PATCH_CODE_14          = 14;
  WIA_PATCH_CODE_CUSTOM_BASE = $8000;

  //
  // WIA_IPS_MICR_READER constants
  //

  WIA_MICR_READER_DISABLED             = 0;
  WIA_MICR_READER_AUTO                 = 1;
  WIA_MICR_READER_FLATBED              = 2;
  WIA_MICR_READER_FEEDER_FRONT         = 3;
  WIA_MICR_READER_FEEDER_BACK          = 4;
  WIA_MICR_READER_FEEDER_DUPLEX        = 5;

  //
  // WIA_IPS_JOB_SEPARATORS constants
  //

  WIA_SEPARATOR_DISABLED               = 0;
  WIA_SEPARATOR_DETECT_SCAN_CONTINUE   = 1;
  WIA_SEPARATOR_DETECT_SCAN_STOP       = 2;
  WIA_SEPARATOR_DETECT_NOSCAN_CONTINUE = 3;
  WIA_SEPARATOR_DETECT_NOSCAN_STOP     = 4;

  //
  // WIA_IPS_LONG_DOCUMENT constants
  //

  WIA_LONG_DOCUMENT_DISABLED           = 0;
  WIA_LONG_DOCUMENT_ENABLED            = 1;
  WIA_LONG_DOCUMENT_SPLIT              = 2;

  //
  // WIA_IPS_BLANK_PAGES constants
  //

  WIA_BLANK_PAGE_DETECTION_DISABLED    = 0;
  WIA_BLANK_PAGE_DISCARD               = 1;
  WIA_BLANK_PAGE_JOB_SEPARATOR         = 2;

  //
  // WIA_IPS_MULTI_FEED constants
  //

  WIA_MULTI_FEED_DETECT_DISABLED       = 0;
  WIA_MULTI_FEED_DETECT_STOP_ERROR     = 1;
  WIA_MULTI_FEED_DETECT_STOP_SUCCESS   = 2;
  WIA_MULTI_FEED_DETECT_CONTINUE       = 3;

  //
  // WIA_IPS_MULTI_FEED_DETECT_METHOD constants
  //

  WIA_MULTI_FEED_DETECT_METHOD_LENGTH  = 0;
  WIA_MULTI_FEED_DETECT_METHOD_OVERLAP = 1;

  //
  // WIA_IPS_AUTO_CROP constants
  //

  WIA_AUTO_CROP_DISABLED               = 0;
  WIA_AUTO_CROP_SINGLE                 = 1;
  WIA_AUTO_CROP_MULTI                  = 2;

  //
  // WIA_IPS_OVER_SCAN constants
  //

  WIA_OVER_SCAN_DISABLED               = 0;
  WIA_OVER_SCAN_TOP_BOTTOM             = 1;
  WIA_OVER_SCAN_LEFT_RIGHT             = 2;
  WIA_OVER_SCAN_ALL                    = 3;

  //
  // WIA_IPS_COLOR_DROP constants
  //

  WIA_COLOR_DROP_DISABLED              = 0;
  WIA_COLOR_DROP_RED                   = 1;
  WIA_COLOR_DROP_GREEN                 = 2;
  WIA_COLOR_DROP_BLUE                  = 3;
  WIA_COLOR_DROP_RGB                   = 4;

  //
  // WIA_IPS_SCAN_AHEAD constants
  //

  WIA_SCAN_AHEAD_DISABLED              = 0;
  WIA_SCAN_AHEAD_ENABLED               = 1;

  //
  // WIA_IPS_FEEDER_CONTROL constants
  //

  WIA_FEEDER_CONTROL_AUTO              = 0;
  WIA_FEEDER_CONTROL_MANUAL            = 1;

  //
  // WIA_IPS_PRINTER_ENDORSER_PADDING constants
  //

  WIA_PRINT_PADDING_NONE               = 0;
  WIA_PRINT_PADDING_ZERO               = 1;
  WIA_PRINT_PADDING_BLANK              = 2;

  //
  // WIA_IPS_PRINTER_ENDORSER_FONT_TYPE constants
  //

  WIA_PRINT_FONT_NORMAL                  = 0;
  WIA_PRINT_FONT_BOLD                    = 1;
  WIA_PRINT_FONT_EXTRA_BOLD              = 2;
  WIA_PRINT_FONT_ITALIC_BOLD             = 3;
  WIA_PRINT_FONT_ITALIC_EXTRA_BOLD       = 4;
  WIA_PRINT_FONT_ITALIC                  = 5;
  WIA_PRINT_FONT_SMALL                   = 6;
  WIA_PRINT_FONT_SMALL_BOLD              = 7;
  WIA_PRINT_FONT_SMALL_EXTRA_BOLD        = 8;
  WIA_PRINT_FONT_SMALL_ITALIC_BOLD       = 9;
  WIA_PRINT_FONT_SMALL_ITALIC_EXTRA_BOLD = 10;
  WIA_PRINT_FONT_SMALL_ITALIC            = 11;
  WIA_PRINT_FONT_LARGE                   = 12;
  WIA_PRINT_FONT_LARGE_BOLD              = 13;
  WIA_PRINT_FONT_LARGE_EXTRA_BOLD        = 14;
  WIA_PRINT_FONT_LARGE_ITALIC_BOLD       = 15;
  WIA_PRINT_FONT_LARGE_ITALIC_EXTRA_BOLD = 16;
  WIA_PRINT_FONT_LARGE_ITALIC            = 17;

  //
  // WIA_IPS_ALARM constants
  //

  WIA_ALARM_NONE   = 0;
  WIA_ALARM_BEEP1  = 1;
  WIA_ALARM_BEEP2  = 2;
  WIA_ALARM_BEEP3  = 3;
  WIA_ALARM_BEEP4  = 4;
  WIA_ALARM_BEEP5  = 5;
  WIA_ALARM_BEEP6  = 6;
  WIA_ALARM_BEEP7  = 7;
  WIA_ALARM_BEEP8  = 8;
  WIA_ALARM_BEEP9  = 9;
  WIA_ALARM_BEEP10 = 10;

type
  //
  // WIA Raw Format header structure:
  //
  _WIA_RAW_HEADER = packed record
      Tag : DWORD;
      Version : DWORD;
      HeaderSize : DWORD;
      XRes : DWORD;
      YRes : DWORD;
      XExtent : DWORD;
      YExtent : DWORD;
      BytesPerLine : DWORD;
      BitsPerPixel : DWORD;
      ChannelsPerPixel : DWORD;
      DataType : DWORD;
      BitsPerChannel : array[0..7] of BYTE;
      Compression : DWORD;
      PhotometricInterp : DWORD;
      LineOrder : DWORD;
      RawDataOffset : DWORD;
      RawDataSize : DWORD;
      PaletteOffset : DWORD;
      PaletteSize : DWORD;
  end;
  WIA_RAW_HEADER = _WIA_RAW_HEADER;
  PWIA_RAW_HEADER = ^_WIA_RAW_HEADER;

  //
  // Barcode Metadata Raw Format structures:
  //
  _WIA_BARCODE_INFO = packed record
      Size : DWORD;
      _Type : DWORD;
      Page : DWORD;
      Confidence : DWORD;
      XOffset : DWORD;
      YOffset : DWORD;
      Rotation : DWORD;
      Length : DWORD;
      Text : array[0..0] of WCHAR;
  end;
  WIA_BARCODE_INFO = _WIA_BARCODE_INFO;

  _WIA_BARCODES = packed record
      Tag : DWORD;
      Version : DWORD;
      Size : DWORD;
      Count : DWORD;
      Barcodes : array[0..0] of WIA_BARCODE_INFO;
  end;
  WIA_BARCODES = _WIA_BARCODES;

  //
  // Patch Code Metadata Raw Format structures:
  //
  _WIA_PATCH_CODE_INFO = packed record
      _Type : DWORD;
   end;
   WIA_PATCH_CODE_INFO = _WIA_PATCH_CODE_INFO;

   _WIA_PATCH_CODES = packed record
       Tag : DWORD;
       Version : DWORD;
       Size : DWORD;
       Count : DWORD;
       PatchCodes : array[0..0] of WIA_PATCH_CODE_INFO;
   end;
   WIA_PATCH_CODES = _WIA_PATCH_CODES;

   //
   // MICR Metadata Raw Format structures:
   //
   _WIA_MICR_INFO = packed record
       Size : DWORD;
       Page : DWORD;
       Length : DWORD;
       Text : array[0..0] of WCHAR;
   end;
   WIA_MICR_INFO = _WIA_MICR_INFO;

   _WIA_MICR = packed record
       Tag : DWORD;
       Version : DWORD;
       Size : DWORD;
       Placeholder : WCHAR;
       Reserved : WORD;
       Count : DWORD;
       Micr : array[0..0] of WIA_MICR_INFO;
   end;
   WIA_MICR = _WIA_MICR;
//{$endif} //(_WIN32_WINNT >= 0x0600)

const
  //
  // Use the WIA property offsets to define private WIA properties.
  //
  // Example: (Creating a private WIA property)
  //
  // WIA_THE_PROP         (WIA_PRIVATE_DEVPROP + 1)
  // WIA_THE_PROP_STR     = 'Property Name';)
  //

  //
  // Private property offset constants
  //

  WIA_PRIVATE_DEVPROP = 38914; // offset for private device (root) item properties
  WIA_PRIVATE_ITEMPROP = 71682; // offset for private item properties

  //
  // WIA image format constants
  //

  WiaImgFmt_UNDEFINED: TGUID = '{b96b3ca9-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_RAWRGB: TGUID = '{bca48b55-f272-4371-b0f1-4a150d057bb4}';
  WiaImgFmt_MEMORYBMP: TGUID = '{b96b3caa-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_BMP: TGUID = '{b96b3cab-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_EMF: TGUID = '{b96b3cac-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_WMF: TGUID = '{b96b3cad-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_JPEG: TGUID = '{b96b3cae-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_PNG: TGUID = '{b96b3caf-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_GIF: TGUID = '{b96b3cb0-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_TIFF: TGUID = '{b96b3cb1-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_EXIF: TGUID = '{b96b3cb2-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_PHOTOCD: TGUID = '{b96b3cb3-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_FLASHPIX: TGUID = '{b96b3cb4-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_ICO: TGUID = '{b96b3cb5-0728-11d3-9d7b-0000f81ef32e}';
  WiaImgFmt_CIFF: TGUID = '{9821a8ab-3a7e-4215-94e0-d27a460c03b2}';
  WiaImgFmt_PICT: TGUID = '{a6bc85d8-6b3e-40ee-a95c-25d482e41adc}';
  WiaImgFmt_JPEG2K: TGUID = '{344ee2b2-39db-4dde-8173-c4b75f8f1e49}';
  WiaImgFmt_JPEG2KX: TGUID = '{43e14614-c80a-4850-baf3-4b152dc8da27}';
  //#if (_WIN32_WINNT >= 0x0600)
  WiaImgFmt_RAW: TGUID = '{6f120719-f1a8-4e07-9ade-9b64c63a3dcc}';
  WiaImgFmt_JBIG: TGUID = '{41e8dd92-2f0a-43d4-8636-f1614ba11e46}';
  WiaImgFmt_JBIG2: TGUID = '{bb8e7e67-283c-4235-9e59-0b9bf94ca687}';
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA document format constants
  //

  WiaImgFmt_RTF: TGUID = '{573dd6a3-4834-432d-a9b5-e198dd9e890d}';
  WiaImgFmt_XML: TGUID = '{b9171457-dac8-4884-b393-15b471d5f07e}';
  WiaImgFmt_HTML: TGUID = '{c99a4e62-99de-4a94-acca-71956ac2977d}';
  WiaImgFmt_TXT: TGUID = '{fafd4d82-723f-421f-9318-30501ac44b59}';
  //#if (_WIN32_WINNT >= 0x0600)
  WiaImgFmt_PDFA: TGUID = '{9980bd5b-3463-43c7-bdca-3caa146f229f}';
  WiaImgFmt_XPS: TGUID = '{700b4a0f-2011-411c-b430-d1e0b2e10b28}';
  WiaImgFmt_OXPS: TGUID = '{2c7b1240-c14d-4109-9755-04b89025153a}';
  WiaImgFmt_CSV: TGUID = '{355bda24-5a9f-4494-80dc-be752cecbc8c}';
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA video format constants
  //

  WiaImgFmt_MPG: TGUID = '{ecd757e4-d2ec-4f57-955d-bcf8a97c4e52}';
  WiaImgFmt_AVI: TGUID = '{32f8ca14-087c-4908-b7c4-6757fe7e90ab}';

  //
  // WIA audio format constants
  //

  WiaAudFmt_WAV: TGUID = '{f818e146-07af-40ff-ae55-be8f2c065dbe}';
  WiaAudFmt_MP3: TGUID = '{0fbc71fb-43bf-49f2-9190-e6fecff37e54}';
  WiaAudFmt_AIFF: TGUID = '{66e2bf4f-b6fc-443f-94c8-2f33c8a65aaf}';
  WiaAudFmt_WMA: TGUID = '{d61d6413-8bc2-438f-93ad-21bd484db6a1}';

  //
  // WIA misc format constants
  //

  WiaImgFmt_ASF: TGUID = '{8d948ee9-d0aa-4a12-9d9a-9cc5de36199b}';
  WiaImgFmt_SCRIPT: TGUID = '{fe7d6c53-2dac-446a-b0bd-d73e21e924c9}';
  WiaImgFmt_EXEC: TGUID = '{485da097-141e-4aa5-bb3b-a5618d95d02b}';
  WiaImgFmt_UNICODE16: TGUID = '{1b7639b6-6357-47d1-9a07-12452dc073e9}';
  WiaImgFmt_DPOF: TGUID = '{369eeeab-a0e8-45ca-86a6-a83ce5697e28}';

  //
  // WIA meta-data format constants
  //

  WiaImgFmt_XMLBAR: TGUID = '{6235701c-3a98-484c-b2a8-fdffd87e6b16}';
  WiaImgFmt_RAWBAR: TGUID = '{da63f833-d26e-451e-90d2-ea55a1365d62}';
  WiaImgFmt_XMLPAT: TGUID = '{f8986f55-f052-460d-9523-3a7dfedbb33c}';
  WiaImgFmt_RAWPAT: TGUID = '{7760507c-5064-400c-9a17-575624d8824b}';
  WiaImgFmt_XMLMIC: TGUID = '{2d164c61-b9ae-4b23-8973-c7067e1fbd31}';
  WiaImgFmt_RAWMIC: TGUID = '{22c4f058-0d88-409c-ac1c-eec12b0ea680}';

  //
  // WIA event constants
  //

  WIA_EVENT_DEVICE_DISCONNECTED: TGUID = '{143e4e83-6497-11d2-a231-00c04fa31809}';
  WIA_EVENT_DEVICE_CONNECTED: TGUID = '{a28bbade-64b6-11d2-a231-00c04fa31809}';
  WIA_EVENT_ITEM_DELETED: TGUID = '{1d22a559-e14f-11d2-b326-00c04f68ce61}';
  WIA_EVENT_ITEM_CREATED: TGUID = '{4c8f4ef5-e14f-11d2-b326-00c04f68ce61}';
  WIA_EVENT_TREE_UPDATED: TGUID = '{c9859b91-4ab2-4cd6-a1fc-582eec55e585}';
  WIA_EVENT_VOLUME_INSERT: TGUID = '{9638bbfd-d1bd-11d2-b31f-00c04f68ce61}';
  WIA_EVENT_SCAN_IMAGE: TGUID = '{a6c5a715-8c6e-11d2-977a-0000f87a926f}';
  WIA_EVENT_SCAN_PRINT_IMAGE: TGUID = '{b441f425-8c6e-11d2-977a-0000f87a926f}';
  WIA_EVENT_SCAN_FAX_IMAGE: TGUID = '{c00eb793-8c6e-11d2-977a-0000f87a926f}';
  WIA_EVENT_SCAN_OCR_IMAGE: TGUID = '{9d095b89-37d6-4877-afed-62a297dc6dbe}';
  WIA_EVENT_SCAN_EMAIL_IMAGE: TGUID = '{c686dcee-54f2-419e-9a27-2fc7f2e98f9e}';
  WIA_EVENT_SCAN_FILM_IMAGE: TGUID = '{9b2b662c-6185-438c-b68b-e39ee25e71cb}';
  WIA_EVENT_SCAN_IMAGE2: TGUID = '{fc4767c1-c8b3-48a2-9cfa-2e90cb3d3590}';
  WIA_EVENT_SCAN_IMAGE3: TGUID = '{154e27be-b617-4653-acc5-0fd7bd4c65ce}';
  WIA_EVENT_SCAN_IMAGE4: TGUID = '{a65b704a-7f3c-4447-a75d-8a26dfca1fdf}';
  WIA_EVENT_STORAGE_CREATED: TGUID = '{353308b2-fe73-46c8-895e-fa4551ccc85a}';
  WIA_EVENT_STORAGE_DELETED: TGUID = '{5e41e75e-9390-44c5-9a51-e47019e390cf}';
  WIA_EVENT_STI_PROXY: TGUID = '{d711f81f-1f0d-422d-8641-927d1b93e5e5}';
  WIA_EVENT_CANCEL_IO: TGUID = '{c860f7b8-9ccd-41ea-bbbf-4dd09c5b1795}';

  //
  // Power management event GUIDs, sent by the WIA service to drivers
  //

  WIA_EVENT_POWER_SUSPEND: TGUID = '{a0922ff9-c3b4-411c-9e29-03a66993d2be}';
  WIA_EVENT_POWER_RESUME: TGUID = '{618f153e-f686-4350-9634-4115a304830c}';

  //
  // No action handler and prompt handler
  //

  WIA_EVENT_HANDLER_NO_ACTION: TGUID = '{e0372b7d-e115-4525-bc55-b629e68c745a}';
  WIA_EVENT_HANDLER_PROMPT: TGUID = '{5f4baad0-4d59-4fcd-b213-783ce7a92f22}';

  //
  // Device status change events (actionable)
  //

  WIA_EVENT_DEVICE_NOT_READY: TGUID = '{d8962d7e-e4dc-4b4d-ba29-668a87f42e6f}';
  WIA_EVENT_DEVICE_READY: TGUID = '{7523ec6c-988b-419e-9a0a-425ac31b37dc}';
  WIA_EVENT_FLATBED_LID_OPEN: TGUID = '{ba0a0623-437d-4f03-a97d-7793b123113c}';
  WIA_EVENT_FLATBED_LID_CLOSED: TGUID = '{f879af0f-9b29-4283-ad95-d412164d39a9}';
  WIA_EVENT_FEEDER_LOADED: TGUID = '{cc8d701e-9aba-481d-bf74-78f763dc342a}';
  WIA_EVENT_FEEDER_EMPTIED: TGUID = '{e70b4b82-6dda-46bb-8ff9-53ceb1a03e35}';
  WIA_EVENT_COVER_OPEN: TGUID = '{19a12136-fa1c-4f66-900f-8f914ec74ec9}';
  WIA_EVENT_COVER_CLOSED: TGUID = '{6714a1e6-e285-468c-9b8c-da7dc4cbaa05}';

  //
  // WIA command constants
  //

  WIA_CMD_SYNCHRONIZE: TGUID = '{9b26b7b2-acad-11d2-a093-00c04f72dc3c}';
  WIA_CMD_TAKE_PICTURE: TGUID = '{af933cac-acad-11d2-a093-00c04f72dc3c}';
  WIA_CMD_DELETE_ALL_ITEMS: TGUID = '{e208c170-acad-11d2-a093-00c04f72dc3c}';
  WIA_CMD_CHANGE_DOCUMENT: TGUID = '{04e725b0-acae-11d2-a093-00c04f72dc3c}';
  WIA_CMD_UNLOAD_DOCUMENT: TGUID = '{1f3b3d8e-acae-11d2-a093-00c04f72dc3c}';
  WIA_CMD_DIAGNOSTIC: TGUID = '{10ff52f5-de04-4cf0-a5ad-691f8dce0141}';
  WIA_CMD_FORMAT: TGUID = '{c3a693aa-f788-4d34-a5b0-be7190759a24}';

  //
  // WIA command constants used for debugging only
  //

  WIA_CMD_DELETE_DEVICE_TREE: TGUID = '{73815942-dbea-11d2-8416-00c04fa36145}';
  WIA_CMD_BUILD_DEVICE_TREE: TGUID = '{9cba5ce0-dbea-11d2-8416-00c04fa36145}';

  //
  // WIA command constants for feeder motor control
  //

  //#if (_WIN32_WINNT >= 0x0600)
  WIA_CMD_START_FEEDER: TGUID = '{5a9df6c9-5f2d-4a39-9d6c-00456d047f00}';
  WIA_CMD_STOP_FEEDER: TGUID = '{d847b06d-3905-459c-9509-9b29cdb691e7}';
  WIA_CMD_PAUSE_FEEDER: TGUID = '{50985e4d-a5b2-4b71-9c95-6d7d7c469a43}';
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  BASE_VAL_WIA_ERROR = $00000000;
  BASE_VAL_WIA_SUCCESS = $00000000;

  FACILITY_WIA = 33;

  WIA_ERROR_GENERAL_ERROR                    = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or 1;
  WIA_ERROR_PAPER_JAM                        = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  2;
  WIA_ERROR_PAPER_EMPTY                      = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  3;
  WIA_ERROR_PAPER_PROBLEM                    = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  4;
  WIA_ERROR_OFFLINE                          = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  5;
  WIA_ERROR_BUSY                             = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  6;
  WIA_ERROR_WARMING_UP                       = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  7;
  WIA_ERROR_USER_INTERVENTION                = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  8;
  WIA_ERROR_ITEM_DELETED                     = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  9;
  WIA_ERROR_DEVICE_COMMUNICATION             = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  10;
  WIA_ERROR_INVALID_COMMAND                  = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  11;
  WIA_ERROR_INCORRECT_HARDWARE_SETTING       = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  12;
  WIA_ERROR_DEVICE_LOCKED                    = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  13;
  WIA_ERROR_EXCEPTION_IN_DRIVER              = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  14;
  WIA_ERROR_INVALID_DRIVER_RESPONSE          = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  15;
  WIA_ERROR_COVER_OPEN                       = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  16;
  WIA_ERROR_LAMP_OFF                         = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  17;
  WIA_ERROR_DESTINATION                      = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  18;
  WIA_ERROR_NETWORK_RESERVATION_FAILED       = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  19;
  //#if (_WIN32_WINNT >= 0x0600)
  WIA_ERROR_MULTI_FEED                       = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  20;
  WIA_ERROR_MAXIMUM_PRINTER_ENDORSER_COUNTER = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  21;
  //#endif //#if (_WIN32_WINNT >= 0x0600)
  WIA_STATUS_END_OF_MEDIA                    = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  1;

  //
  // Definitions for errors and status codes passed to IWiaDataTransfer::BandedDataCallback as the lReason parameter.
  // These codes are in addition to the errors defined above; in some cases the SEVERITY_SUCCESS version of
  // an error is meant to replace the SEVERITY_ERROR version listed above.
  //

  WIA_STATUS_WARMING_UP                        = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  2;
  WIA_STATUS_CALIBRATING                       = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  3;
  WIA_STATUS_RESERVING_NETWORK_DEVICE          = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  6;
  WIA_STATUS_NETWORK_DEVICE_RESERVED           = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  7;
  WIA_STATUS_CLEAR                             = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  8;
  WIA_STATUS_SKIP_ITEM                         = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  9;
  WIA_STATUS_NOT_HANDLED                       = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  10;

  //
  // The value is returned by Scansetting.dll when the user chooses to change the scanner in scandialog
  //

  WIA_S_CHANGE_DEVICE                          = (SEVERITY_SUCCESS shl 31) or (FACILITY_WIA shl 16) or  11;

  //
  // SelectDeviceDlg and SelectDeviceDlgID status code when there are no devices available
  //

  WIA_S_NO_DEVICE_AVAILABLE                    = (SEVERITY_ERROR shl 31) or (FACILITY_WIA shl 16) or  21;

  //
  // SelectDeviceDlg and GetImageDlg flag constants
  //

  WIA_SELECT_DEVICE_NODEFAULT = $00000001;

  //
  // DeviceDlg and GetImageDlg flags constants
  //

  WIA_DEVICE_DIALOG_SINGLE_IMAGE  = $00000002;  // Only allow one image to be selected
  WIA_DEVICE_DIALOG_USE_COMMON_UI = $00000004;  // Give preference to the system-provided UI, if available

  //
  // RegisterEventCallbackInterface and RegisterEventCallbackCLSID flag constants
  //

   WIA_REGISTER_EVENT_CALLBACK   = $00000001;
   WIA_UNREGISTER_EVENT_CALLBACK = $00000002;
   WIA_SET_DEFAULT_HANDLER       = $00000004;

  //
  // WIA event type constants
  //

   WIA_NOTIFICATION_EVENT = $00000001;
   WIA_ACTION_EVENT       = $00000002;

  //
  // Additional WIA raw format constants
  //

   WIA_LINE_ORDER_TOP_TO_BOTTOM = $00000001;
   WIA_LINE_ORDER_BOTTOM_TO_TOP = $00000002;

  //
  // WIA event persistent handler flag constants
  //

   WIA_IS_DEFAULT_HANDLER = $00000001;

  //
  // WIA connected and disconnected event description strings
  //

  WIA_EVENT_DEVICE_DISCONNECTED_STR = 'Device Disconnected';
  WIA_EVENT_DEVICE_CONNECTED_STR    = 'Device Connected';

  //
  // WIA event and command icon resource identifier constants
  //
  // Events   : -1000 to -1499 (Standard), -1500 to -1999 (Custom)
  // Commands : -2000 to -2499 (Standard), -2500 to -2999 (Custom)
  //

  WIA_ICON_DEVICE_DISCONNECTED = 'sti.dll,-1001';
  WIA_ICON_DEVICE_CONNECTED    = 'sti.dll,-1001';
  WIA_ICON_ITEM_DELETED        = 'sti.dll,-1001';
  WIA_ICON_ITEM_CREATED        = 'sti.dll,-1001';
  WIA_ICON_TREE_UPDATED        = 'sti.dll,-1001';
  WIA_ICON_VOLUME_INSERT       = 'sti.dll,-1001';
  WIA_ICON_SCAN_BUTTON_PRESS   = 'sti.dll,-1001';
  WIA_ICON_DEVICE_NOT_READY    = 'sti.dll,-1001';
  WIA_ICON_DEVICE_READY        = 'sti.dll,-1001';
  WIA_ICON_FLATBED_LID_OPEN    = 'sti.dll,-1001';
  WIA_ICON_FLATBED_LID_CLOSED  = 'sti.dll,-1001';
  WIA_ICON_FEEDER_LOADED       = 'sti.dll,-1001';
  WIA_ICON_FEEDER_EMPTIED      = 'sti.dll,-1001';
  WIA_ICON_COVER_OPEN          = 'sti.dll,-1001';
  WIA_ICON_COVER_CLOSED        = 'sti.dll,-1001';
  WIA_ICON_START_FEEDER        = 'sti.dll,-1001';
  WIA_ICON_STOP_FEEDER         = 'sti.dll,-1001';
  WIA_ICON_SYNCHRONIZE         = 'sti.dll,-2000';
  WIA_ICON_TAKE_PICTURE        = 'sti.dll,-2001';
  WIA_ICON_DELETE_ALL_ITEMS    = 'sti.dll,-2002';
  WIA_ICON_CHANGE_DOCUMENT     = 'sti.dll,-2003';
  WIA_ICON_UNLOAD_DOCUMENT     = 'sti.dll,-2004';
  WIA_ICON_DELETE_DEVICE_TREE  = 'sti.dll,-2005';
  WIA_ICON_BUILD_DEVICE_TREE   = 'sti.dll,-2006';

  //
  // WIA TYMED constants
  //

  TYMED_CALLBACK           = 128;
  TYMED_MULTIPAGE_FILE     = 256;
  TYMED_MULTIPAGE_CALLBACK = 512;

  //
  // IWiaDataCallback and IWiaMiniDrvCallBack message ID constants
  //

  IT_MSG_DATA_HEADER              = $0001;
  IT_MSG_DATA                     = $0002;
  IT_MSG_STATUS                   = $0003;
  IT_MSG_TERMINATION              = $0004;
  IT_MSG_NEW_PAGE                 = $0005;
  IT_MSG_FILE_PREVIEW_DATA        = $0006;
  IT_MSG_FILE_PREVIEW_DATA_HEADER = $0007;

  //
  // IWiaDataCallback and IWiaMiniDrvCallBack status flag constants
  //

  IT_STATUS_TRANSFER_FROM_DEVICE = $0001;
  IT_STATUS_PROCESSING_DATA      = $0002;
  IT_STATUS_TRANSFER_TO_CLIENT   = $0004;
  IT_STATUS_MASK                 = $0007; // any status value that doesn't fit the mask is an HRESULT
  //
  // IWiaTransfer flags
  //

  WIA_TRANSFER_ACQUIRE_CHILDREN = $0001;

  //
  // IWiaTransferCallback Message types
  //

  WIA_TRANSFER_MSG_STATUS          = $00001;
  WIA_TRANSFER_MSG_END_OF_STREAM   = $00002;
  WIA_TRANSFER_MSG_END_OF_TRANSFER = $00003;
  WIA_TRANSFER_MSG_DEVICE_STATUS   = $00005;
  WIA_TRANSFER_MSG_NEW_PAGE        = $00006;

  //
  // IWiaEventCallback code constants
  //

  WIA_MAJOR_EVENT_DEVICE_CONNECT    = $01;
  WIA_MAJOR_EVENT_DEVICE_DISCONNECT = $02;
  WIA_MAJOR_EVENT_PICTURE_TAKEN     = $03;
  WIA_MAJOR_EVENT_PICTURE_DELETED   = $04;

  //
  // WIA device connection status constants
  //

   WIA_DEVICE_NOT_CONNECTED = 0;
   WIA_DEVICE_CONNECTED     = 1;

  //
  // EnumDeviceCapabilities and drvGetCapabilities flags
  //

  WIA_DEVICE_COMMANDS     = 1;
  WIA_DEVICE_EVENTS       = 2;

  //
  // EnumDeviceInfo Flags
  //

  WIA_DEVINFO_ENUM_ALL    = $0000000F;
  WIA_DEVINFO_ENUM_LOCAL  = $00000010;


  //
  // WIA item type constants
  //

  WiaItemTypeFree           = $00000000;
  WiaItemTypeImage          = $00000001;
  WiaItemTypeFile           = $00000002;
  WiaItemTypeFolder         = $00000004;
  WiaItemTypeRoot           = $00000008;
  WiaItemTypeAnalyze        = $00000010;
  WiaItemTypeAudio          = $00000020;
  WiaItemTypeDevice         = $00000040;
  WiaItemTypeDeleted        = $00000080;
  WiaItemTypeDisconnected   = $00000100;
  WiaItemTypeHPanorama      = $00000200;
  WiaItemTypeVPanorama      = $00000400;
  WiaItemTypeBurst          = $00000800;
  WiaItemTypeStorage        = $00001000;
  WiaItemTypeTransfer       = $00002000;
  WiaItemTypeGenerated      = $00004000;
  WiaItemTypeHasAttachments = $00008000;
  WiaItemTypeVideo          = $00010000;
  WiaItemTypeRemoved        = $80000000;
  //
  // 0x00020000 is reserved for the TWAIN-WIA Compatiblity Layer
  //
  //#if (_WIN32_WINNT >= 0x0600)
  WiaItemTypeDocument               = $00040000;
  WiaItemTypeProgrammableDataSource = $00080000;
  {$ifdef win32_xp}
  WiaItemTypeMask         = $8003FFFF;
  {$else}
  WiaItemTypeMask         = $800FFFFF;
  {$endif}

  //
  // Maximum device specific item context
  //

  WIA_MAX_CTX_SIZE        = $01000000;

  //
  // WIA property access flag constants
  //

  WIA_PROP_READ          = $01;
  WIA_PROP_WRITE         = $02;
  WIA_PROP_RW            = (WIA_PROP_READ or WIA_PROP_WRITE);
  WIA_PROP_SYNC_REQUIRED = $04;

  WIA_PROP_NONE  = $08;
  WIA_PROP_RANGE = $10;
  WIA_PROP_LIST  = $20;
  WIA_PROP_FLAG  = $40;

  WIA_PROP_CACHEABLE = $10000;

  //
  // IWiaItem2 CreateChildItem flag constants
  //

  COPY_PARENT_PROPERTY_VALUES = $40000000;

  //
  // WIA item access flag constants
  //

  WIA_ITEM_CAN_BE_DELETED = $80;
  WIA_ITEM_READ           = WIA_PROP_READ;
  WIA_ITEM_WRITE          = WIA_PROP_WRITE;
  WIA_ITEM_RD             = (WIA_ITEM_READ or WIA_ITEM_CAN_BE_DELETED);
  WIA_ITEM_RWD            = (WIA_ITEM_READ or WIA_ITEM_WRITE or WIA_ITEM_CAN_BE_DELETED);

  //
  // WIA property container constants
  //

   WIA_RANGE_MIN       = 0;
   WIA_RANGE_NOM       = 1;
   WIA_RANGE_MAX       = 2;
   WIA_RANGE_STEP      = 3;
   WIA_RANGE_NUM_ELEMS = 4;

   WIA_LIST_COUNT      = 0;
   WIA_LIST_NOM        = 1;
   WIA_LIST_VALUES     = 2;
   WIA_LIST_NUM_ELEMS  = 2;

   WIA_FLAG_NOM        = 0;
   WIA_FLAG_VALUES     = 1;
   WIA_FLAG_NUM_ELEMS  = 2;

  //
  // WIA property LIST container MACROS
  //

 // TO-DO MaxM
 // WIA_PROP_LIST_COUNT(ppv) (((PROPVARIANT*)ppv)->cal.cElems - WIA_LIST_VALUES)

 // WIA_PROP_LIST_VALUE(ppv, index)                              \\
 //      ((index > ((PROPVARIANT*) ppv)->cal.cElems - WIA_LIST_VALUES) || (index < -WIA_LIST_NOM)) ?\\
 //      NULL :                                                          \\
 //      (((PROPVARIANT*) ppv)->vt == VT_UI1) ?                          \\
 //      ((PROPVARIANT*) ppv)->caub.pElems[WIA_LIST_VALUES + index] :    \\
 //      (((PROPVARIANT*) ppv)->vt == VT_UI2) ?                          \\
 //      ((PROPVARIANT*) ppv)->caui.pElems[WIA_LIST_VALUES + index] :    \\
 //      (((PROPVARIANT*) ppv)->vt == VT_UI4) ?                          \\
 //      ((PROPVARIANT*) ppv)->caul.pElems[WIA_LIST_VALUES + index] :    \\
 //      (((PROPVARIANT*) ppv)->vt == VT_I2) ?                           \\
 //      ((PROPVARIANT*) ppv)->cai.pElems[WIA_LIST_VALUES + index] :     \\
 //      (((PROPVARIANT*) ppv)->vt == VT_I4) ?                           \\
 //      ((PROPVARIANT*) ppv)->cal.pElems[WIA_LIST_VALUES + index] :     \\
  //     (((PROPVARIANT*) ppv)->vt == VT_R4) ?                           \\
  //     ((PROPVARIANT*) ppv)->caflt.pElems[WIA_LIST_VALUES + index] :   \\
  //     (((PROPVARIANT*) ppv)->vt == VT_R8) ?                           \\
  //     ((PROPVARIANT*) ppv)->cadbl.pElems[WIA_LIST_VALUES + index] :   \\
  //     (((PROPVARIANT*) ppv)->vt == VT_BSTR) ?                         \\
  //     (LONG)(((PROPVARIANT*) ppv)->cabstr.pElems[WIA_LIST_VALUES + index]) : \\
  //     NULL

  //
  // Microsoft defined WIA property offset constants
  //

  WIA_DIP_FIRST              = 2;
  WIA_IPA_FIRST              = 4098;
  WIA_DPF_FIRST              = 3330;
  WIA_IPS_FIRST              = 6146;
  WIA_DPS_FIRST              = 3074;
  WIA_IPC_FIRST              = 5122;
  WIA_NUM_IPC                = 9;
  WIA_RESERVED_FOR_NEW_PROPS = 1024;

  //
  // WIA_DPC_WHITE_BALANCE constants
  //

  WHITEBALANCE_MANUAL       = 1;
  WHITEBALANCE_AUTO         = 2;
  WHITEBALANCE_ONEPUSH_AUTO = 3;
  WHITEBALANCE_DAYLIGHT     = 4;
  WHITEBALANCE_FLORESCENT   = 5;
  WHITEBALANCE_TUNGSTEN     = 6;
  WHITEBALANCE_FLASH        = 7;

  //
  // WIA_DPC_FOCUS_MODE constants
  //

  FOCUSMODE_MANUAL    = 1;
  FOCUSMODE_AUTO      = 2;
  FOCUSMODE_MACROAUTO = 3;

  //
  // WIA_DPC_EXPOSURE_METERING_MODE constants
  //

  EXPOSUREMETERING_AVERAGE      = 1;
  EXPOSUREMETERING_CENTERWEIGHT = 2;
  EXPOSUREMETERING_MULTISPOT    = 3;
  EXPOSUREMETERING_CENTERSPOT   = 4;

  //
  // WIA_DPC_FLASH_MODE constants
  //

  FLASHMODE_AUTO         = 1;
  FLASHMODE_OFF          = 2;
  FLASHMODE_FILL         = 3;
  FLASHMODE_REDEYE_AUTO  = 4;
  FLASHMODE_REDEYE_FILL  = 5;
  FLASHMODE_EXTERNALSYNC = 6;

  //
  // WIA_DPC_EXPOSURE_MODE constants
  //

  EXPOSUREMODE_MANUAL            = 1;
  EXPOSUREMODE_AUTO              = 2;
  EXPOSUREMODE_APERTURE_PRIORITY = 3;
  EXPOSUREMODE_SHUTTER_PRIORITY  = 4;
  EXPOSUREMODE_PROGRAM_CREATIVE  = 5;
  EXPOSUREMODE_PROGRAM_ACTION    = 6;
  EXPOSUREMODE_PORTRAIT          = 7;

  //
  // WIA_DPC_CAPTURE_MODE constants
  //

  CAPTUREMODE_NORMAL    = 1;
  CAPTUREMODE_BURST     = 2;
  CAPTUREMODE_TIMELAPSE = 3;

  //
  // WIA_DPC_EFFECT_MODE constants
  //

  EFFECTMODE_STANDARD = 1;
  EFFECTMODE_BW       = 2;
  EFFECTMODE_SEPIA    = 3;

  //
  // WIA_DPC_FOCUS_METERING_MODE constants
  //

  FOCUSMETERING_CENTERSPOT = 1;
  FOCUSMETERING_MULTISPOT  = 2;

  //
  // WIA_DPC_POWER_MODE constants
  //

  POWERMODE_LINE    = 1;
  POWERMODE_BATTERY = 2;

  //
  // WIA_DPS_SHEET_FEEDER_REGISTRATION and
  // WIA_DPS_HORIZONTAL_BED_REGISTRATION constants
  //

   LEFT_JUSTIFIED  = 0;
   CENTERED        = 1;
   RIGHT_JUSTIFIED = 2;

  //
  // WIA_DPS_VERTICAL_BED_REGISTRATION constants
  //

   TOP_JUSTIFIED    = 0;
   //CENTERED         = 1;
   BOTTOM_JUSTIFIED = 2;

  //
  // WIA_IPS_ORIENTATION and WIA_IPS_ROTATION constants
  //

   PORTRAIT  = 0;
   LANSCAPE  = 1;
  //#if (_WIN32_WINNT >= 0x0600)
   LANDSCAPE = LANSCAPE;
  //#endif
   ROT180    = 2;
   ROT270    = 3;


  //
  // WIA_IPS_MIRROR flags
  //

   MIRRORED = $01;

  //
  // WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES flags
  //

   FEED              = $01;
   FLAT              = $02;
   DUP               = $04;
   DETECT_FLAT       = $08;
   DETECT_SCAN       = $10;
   DETECT_FEED       = $20;
   DETECT_DUP        = $40;
   DETECT_FEED_AVAIL = $80;
   DETECT_DUP_AVAIL  = $100;
  //#if (_WIN32_WINNT >= 0x0600)
   FILM_TPA          = $200;
   DETECT_FILM_TPA   = $400;
   STOR              = $800;
   DETECT_STOR       = $1000;
   ADVANCED_DUP      = $2000;
   AUTO_SOURCE       = $8000;
   IMPRINTER         = $10000;
   ENDORSER          = $20000;
   BARCODE_READER    = $40000;
   PATCH_CODE_READER = $80000;
   MICR_READER       = $100000;
  //#endif

  //
  // WIA_DPS_DOCUMENT_HANDLING_STATUS flags
  //

   FEED_READY              = $01;
   FLAT_READY              = $02;
   DUP_READY               = $04;
   FLAT_COVER_UP           = $08;
   PATH_COVER_UP           = $10;
   PAPER_JAM               = $20;
  //#if (_WIN32_WINNT >= 0x0600)
   FILM_TPA_READY          = $40;
   STORAGE_READY           = $80;
   STORAGE_FULL            = $100;
   MULTIPLE_FEED           = $200;
   DEVICE_ATTENTION        = $400;
   LAMP_ERR                = $800;
   IMPRINTER_READY         = $1000;
   ENDORSER_READY          = $2000;
   BARCODE_READER_READY    = $4000;
   PATCH_CODE_READER_READY = $8000;
   MICR_READER_READY       = $10000;
  //#endif

  //
  // WIA_DPS_DOCUMENT_HANDLING_SELECT flags
  //

   FEEDER          = $001;
   FLATBED         = $002;
   DUPLEX          = $004;
   FRONT_FIRST     = $008;
   BACK_FIRST      = $010;
   FRONT_ONLY      = $020;
   BACK_ONLY       = $040;
   NEXT_PAGE       = $080;
   PREFEED         = $100;
   AUTO_ADVANCE    = $200;
  //#if (_WIN32_WINNT >= 0x0600)
  //
  // New WIA_IPS_DOCUMENT_HANDLING_SELECT flag
  //
   ADVANCED_DUPLEX = $400;
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA_DPS_TRANSPARENCY / WIA_DPS_TRANSPARENCY_STATUS flags
  //

   LIGHT_SOURCE_PRESENT_DETECT = $01;
   LIGHT_SOURCE_PRESENT        = $02;
   LIGHT_SOURCE_DETECT_READY   = $04;
   LIGHT_SOURCE_READY          = $08;

  //
  // WIA_DPS_TRANSPARENCY_CAPABILITIES
  //

  TRANSPARENCY_DYNAMIC_FRAME_SUPPORT = $01;
  TRANSPARENCY_STATIC_FRAME_SUPPORT  = $02;

  //
  // WIA_DPS_TRANSPARENCY_SELECT flags
  //

   LIGHT_SOURCE_SELECT   = $001; // currently not used
   LIGHT_SOURCE_POSITIVE = $002;
   LIGHT_SOURCE_NEGATIVE = $004;

  //
  // WIA_DPS_SCAN_AHEAD_PAGES constants
  //
  // WIA_DPS_SCAN_AHEAD_PAGES is superseded in WIA 2.0 by WIA_IPS_SCAN_AHEAD
  //

   WIA_SCAN_AHEAD_ALL = 0;

  //
  // WIA_DPS_PAGES constants
  //

   ALL_PAGES = 0;

  //
  // WIA_DPS_PREVIEW constants
  //

  WIA_FINAL_SCAN   = 0;
  WIA_PREVIEW_SCAN = 1;

  //
  // WIA_DPS_SHOW_PREVIEW_CONTROL constants
  //

  WIA_SHOW_PREVIEW_CONTROL      = 0;
  WIA_DONT_SHOW_PREVIEW_CONTROL = 1;

  //
  // Predefined strings for WIA_DPS_ENDORSER_STRING
  //
  // WIA_DPS_ENDORSER_STRING is superseded in WIA 2.0 by WIA_IPS_PRINTER_ENDORSER_STRING and these
  // constant values are replaced by the WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS constants
  //

  WIA_ENDORSER_TOK_DATE = '$DATE$';
  WIA_ENDORSER_TOK_TIME = '$TIME$';
  WIA_ENDORSER_TOK_PAGE_COUNT = '$PAGE_COUNT$';
  WIA_ENDORSER_TOK_DAY = '$DAY$';
  WIA_ENDORSER_TOK_MONTH = '$MONTH$';
  WIA_ENDORSER_TOK_YEAR = '$YEAR$';

  //
  // WIA_DPS_PAGE_SIZE/WIA_IPS_PAGE_SIZE constants
  // Dimensions are defined as (WIDTH x HEIGHT) in 1/1000ths of an inch
  //

  WIA_PAGE_A4           = 0; //  8267 x 11692
  WIA_PAGE_LETTER       = 1; //  8500 x 11000
  WIA_PAGE_CUSTOM       = 2; // (current extent settings)

  WIA_PAGE_USLEGAL      = 3; //  8500 x 14000
  WIA_PAGE_USLETTER     = WIA_PAGE_LETTER;
  WIA_PAGE_USLEDGER     = 4; // 11000 x 17000
  WIA_PAGE_USSTATEMENT  = 5; //  5500 x  8500
  WIA_PAGE_BUSINESSCARD = 6; //  3543 x  2165

  //
  // ISO A page size constants
  //

  WIA_PAGE_ISO_A0 = 7; // 33110 x 46811
  WIA_PAGE_ISO_A1 = 8; // 23385 x 33110
  WIA_PAGE_ISO_A2 = 9; // 16535 x 23385
  WIA_PAGE_ISO_A3  = 10; // 11692 x 16535
  WIA_PAGE_ISO_A4  = WIA_PAGE_A4;
  WIA_PAGE_ISO_A5  = 11; //  5826 x  8267
  WIA_PAGE_ISO_A6  = 12; //  4133 x  5826
  WIA_PAGE_ISO_A7  = 13; //  2913 x  4133
  WIA_PAGE_ISO_A8  = 14; //  2047 x  2913
  WIA_PAGE_ISO_A9  = 15; //  1456 x  2047
  WIA_PAGE_ISO_A10 = 16; //  1023 x  1456

  //
  // ISO B page size constants
  //

  WIA_PAGE_ISO_B0  = 17; //  39370 x 55669
  WIA_PAGE_ISO_B1  = 18; //  27834 x 39370
  WIA_PAGE_ISO_B2  = 19; //  19685 x 27834
  WIA_PAGE_ISO_B3  = 20; //  13897 x 19685
  WIA_PAGE_ISO_B4  = 21; //   9842 x 13897
  WIA_PAGE_ISO_B5  = 22; //   6929 x  9842
  WIA_PAGE_ISO_B6  = 23; //   4921 x  6929
  WIA_PAGE_ISO_B7  = 24; //   3464 x  4921
  WIA_PAGE_ISO_B8  = 25; //   2440 x  3464
  WIA_PAGE_ISO_B9  = 26; //   1732 x  2440
  WIA_PAGE_ISO_B10 = 27; //   1220 x  1732

  //
  // ISO C page size constants
  //

  WIA_PAGE_ISO_C0  = 28; //  36102 x 51062
  WIA_PAGE_ISO_C1  = 29; //  25511 x 36102
  WIA_PAGE_ISO_C2  = 30; //  18031 x 25511
  WIA_PAGE_ISO_C3  = 31; //  12755 x 18031
  WIA_PAGE_ISO_C4  = 32; //   9015 x 12755 (unfolded)
  WIA_PAGE_ISO_C5  = 33; //   6377 x  9015 (folded once)
  WIA_PAGE_ISO_C6  = 34; //   4488 x  6377 (folded twice)
  WIA_PAGE_ISO_C7  = 35; //   3188 x  4488
  WIA_PAGE_ISO_C8  = 36; //   2244 x  3188
  WIA_PAGE_ISO_C9  = 37; //   1574 x  2244
  WIA_PAGE_ISO_C10 = 38; //   1102 x  1574

  //
  // JIS B page size constants
  //

  WIA_PAGE_JIS_B0  = 39; //  40551 x 57322
  WIA_PAGE_JIS_B1  = 40; //  28661 x 40551
  WIA_PAGE_JIS_B2  = 41; //  20275 x 28661
  WIA_PAGE_JIS_B3  = 42; //  14330 x 20275
  WIA_PAGE_JIS_B4  = 43; //  10118 x 14330
  WIA_PAGE_JIS_B5  = 44; //   7165 x 10118
  WIA_PAGE_JIS_B6  = 45; //   5039 x  7165
  WIA_PAGE_JIS_B7  = 46; //   3582 x  5039
  WIA_PAGE_JIS_B8  = 47; //   2519 x  3582
  WIA_PAGE_JIS_B9  = 48; //   1771 x  2519
  WIA_PAGE_JIS_B10 = 49; //   1259 x  1771

  //
  // JIS A page size constants
  //

  WIA_PAGE_JIS_2A = 50; //  46811 x 66220
  WIA_PAGE_JIS_4A = 51; //  66220 x  93622

  //
  // DIN B page size constants
  //

  WIA_PAGE_DIN_2B = 52; //  55669 x 78740
  WIA_PAGE_DIN_4B = 53; //  78740 x 111338

  //#if (_WIN32_WINNT >= 0x0600)
  //
  // Additional WIA_IPS_PAGE_SIZE constants:
  //
  WIA_PAGE_AUTO = 100;
  WIA_PAGE_CUSTOM_BASE = $8000;
  //#endif //#if (_WIN32_WINNT >= 0x0600)


  //
  // WIA_IPA_COMPRESSION constants
  //

  WIA_COMPRESSION_NONE    = 0;
  WIA_COMPRESSION_BI_RLE4 = 1;
  WIA_COMPRESSION_BI_RLE8 = 2;
  WIA_COMPRESSION_G3      = 3;
  WIA_COMPRESSION_G4      = 4;
  WIA_COMPRESSION_JPEG    = 5;
  //#if (_WIN32_WINNT >= 0x0600)
  WIA_COMPRESSION_JBIG    = 6;
  WIA_COMPRESSION_JPEG2K  = 7;
  WIA_COMPRESSION_PNG     = 8;
  WIA_COMPRESSION_AUTO    = 100;
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA_IPA_PLANAR constants
  //

  WIA_PACKED_PIXEL = 0;
  WIA_PLANAR       = 1;

  //
  // WIA_IPA_DATATYPE constants
  //

  WIA_DATA_THRESHOLD       = 0;
  WIA_DATA_DITHER          = 1;
  WIA_DATA_GRAYSCALE       = 2;
  WIA_DATA_COLOR           = 3;
  WIA_DATA_COLOR_THRESHOLD = 4;
  WIA_DATA_COLOR_DITHER    = 5;
  //#if (_WIN32_WINNT >= 0x0600)
  WIA_DATA_RAW_RGB         = 6;
  WIA_DATA_RAW_BGR         = 7;
  WIA_DATA_RAW_YUV         = 8;
  WIA_DATA_RAW_YUVK        = 9;
  WIA_DATA_RAW_CMY         = 10;
  WIA_DATA_RAW_CMYK        = 11;
  WIA_DATA_AUTO            = 100;
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA_IPA_DEPTH constant
  //
  //#if (_WIN32_WINNT >= 0x0600)
  WIA_DEPTH_AUTO = 0;
  //#endif //#if (_WIN32_WINNT >= 0x0600)

  //
  // WIA_IPS_PHOTOMETRIC_INTERP constants
  //

  WIA_PHOTO_WHITE_1 = 0; // white is 1, black is 0
  WIA_PHOTO_WHITE_0 = 1; // white is 0, black is 1

  //
  // WIA_IPA_SUPPRESS_PROPERTY_PAGE flags
  //

  WIA_PROPPAGE_SCANNER_ITEM_GENERAL = $00000001;
  WIA_PROPPAGE_CAMERA_ITEM_GENERAL  = $00000002;
  WIA_PROPPAGE_DEVICE_GENERAL       = $00000004;

  //
  // WIA_IPS_CUR_INTENT flags
  //
  WIA_INTENT_NONE                 = $00000000;
  WIA_INTENT_IMAGE_TYPE_COLOR     = $00000001;
  WIA_INTENT_IMAGE_TYPE_GRAYSCALE = $00000002;
  WIA_INTENT_IMAGE_TYPE_TEXT      = $00000004;
  WIA_INTENT_IMAGE_TYPE_MASK      = $0000000F;
  WIA_INTENT_MINIMIZE_SIZE        = $00010000;
  WIA_INTENT_MAXIMIZE_QUALITY     = $00020000;
  WIA_INTENT_BEST_PREVIEW         = $00040000;
  WIA_INTENT_SIZE_MASK            = $000F0000;

  //
  // Global WIA device information property arrays
  //

  WIA_NUM_DIP = 16;

  {$ifdef WIA_DECLARE_DEVINFO_PROP_ARRAY}
  g_psDeviceInfo : array[0..WIA_NUM_DIP-1] of PROPSPEC =
  (
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_DEV_ID),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_VEND_DESC),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_DEV_DESC),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_DEV_TYPE),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_PORT_NAME),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_DEV_NAME),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_SERVER_NAME),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_REMOTE_DEV_ID),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_UI_CLSID),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_HW_CONFIG),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_BAUDRATE),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_STI_GEN_CAPABILITIES),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_WIA_VERSION),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_DRIVER_VERSION),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_PNP_ID),
      (ulKind: PRSPEC_PROPID; propid: WIA_DIP_STI_DRIVER_VERSION)
  );

  g_piDeviceInfo : array[0..WIA_NUM_DIP-1] of PROPID =
  (
      WIA_DIP_DEV_ID,
      WIA_DIP_VEND_DESC,
      WIA_DIP_DEV_DESC,
      WIA_DIP_DEV_TYPE,
      WIA_DIP_PORT_NAME,
      WIA_DIP_DEV_NAME,
      WIA_DIP_SERVER_NAME,
      WIA_DIP_REMOTE_DEV_ID,
      WIA_DIP_UI_CLSID,
      WIA_DIP_HW_CONFIG,
      WIA_DIP_BAUDRATE,
      WIA_DIP_STI_GEN_CAPABILITIES,
      WIA_DIP_WIA_VERSION,
      WIA_DIP_DRIVER_VERSION,
      WIA_DIP_PNP_ID,
      WIA_DIP_STI_DRIVER_VERSION
  );

  g_pszDeviceInfo : array[0..WIA_NUM_DIP-1] of PCWSTR =
  (
      WIA_DIP_DEV_ID_STR,
      WIA_DIP_VEND_DESC_STR,
      WIA_DIP_DEV_DESC_STR,
      WIA_DIP_DEV_TYPE_STR,
      WIA_DIP_PORT_NAME_STR,
      WIA_DIP_DEV_NAME_STR,
      WIA_DIP_SERVER_NAME_STR,
      WIA_DIP_REMOTE_DEV_ID_STR,
      WIA_DIP_UI_CLSID_STR,
      WIA_DIP_HW_CONFIG_STR,
      WIA_DIP_BAUDRATE_STR,
      WIA_DIP_STI_GEN_CAPABILITIES_STR,
      WIA_DIP_WIA_VERSION_STR,
      WIA_DIP_DRIVER_VERSION_STR,
      WIA_DIP_PNP_ID_STR,
      WIA_DIP_STI_DRIVER_VERSION_STR
  );

  {$else}
  var
     g_psDeviceInfo : array[0..WIA_NUM_DIP-1] of PROPSPEC; //cvar; external;
     g_piDeviceInfo : array[0..WIA_NUM_DIP-1] of PROPID;   //cvar;external;
     g_pszDeviceInfo : array[0..WIA_NUM_DIP-1] of LPOLESTR;//cvar;external;
  {$endif}

  //
  // Global WIA property ID to property name array
  //

  {$ifdef DEFINE_WIA_PROPID_TO_NAME}
  const
    g_wiaPropIdToName : array [0..214] of WIA_PROPID_TO_NAME =
    (
      (propid: WIA_DIP_DEV_ID;                                   pszName: WIA_DIP_DEV_ID_STR),
      (propid: WIA_DIP_VEND_DESC;                                pszName: WIA_DIP_VEND_DESC_STR),
      (propid: WIA_DIP_DEV_DESC;                                 pszName: WIA_DIP_DEV_DESC_STR),
      (propid: WIA_DIP_DEV_TYPE;                                 pszName: WIA_DIP_DEV_TYPE_STR),
      (propid: WIA_DIP_PORT_NAME;                                pszName: WIA_DIP_PORT_NAME_STR),
      (propid: WIA_DIP_DEV_NAME;                                 pszName: WIA_DIP_DEV_NAME_STR),
      (propid: WIA_DIP_SERVER_NAME;                              pszName: WIA_DIP_SERVER_NAME_STR),
      (propid: WIA_DIP_REMOTE_DEV_ID;                            pszName: WIA_DIP_REMOTE_DEV_ID_STR),
      (propid: WIA_DIP_UI_CLSID;                                 pszName: WIA_DIP_UI_CLSID_STR),
      (propid: WIA_DIP_HW_CONFIG;                                pszName: WIA_DIP_HW_CONFIG_STR),
      (propid: WIA_DIP_BAUDRATE;                                 pszName: WIA_DIP_BAUDRATE_STR),
      (propid: WIA_DIP_STI_GEN_CAPABILITIES;                     pszName: WIA_DIP_STI_GEN_CAPABILITIES_STR),
      (propid: WIA_DIP_WIA_VERSION;                              pszName: WIA_DIP_WIA_VERSION_STR),
      (propid: WIA_DIP_DRIVER_VERSION;                           pszName: WIA_DIP_DRIVER_VERSION_STR),
      (propid: WIA_DIP_PNP_ID;                                   pszName: WIA_DIP_PNP_ID_STR),
      (propid: WIA_DIP_STI_DRIVER_VERSION;                       pszName: WIA_DIP_STI_DRIVER_VERSION_STR),
      (propid: WIA_DPA_FIRMWARE_VERSION;                         pszName: WIA_DPA_FIRMWARE_VERSION_STR),
      (propid: WIA_DPA_CONNECT_STATUS;                           pszName: WIA_DPA_CONNECT_STATUS_STR),
      (propid: WIA_DPA_DEVICE_TIME;                              pszName: WIA_DPA_DEVICE_TIME_STR),
      (propid: WIA_DPC_PICTURES_TAKEN;                           pszName: WIA_DPC_PICTURES_TAKEN_STR),
      (propid: WIA_DPC_PICTURES_REMAINING;                       pszName: WIA_DPC_PICTURES_REMAINING_STR),
      (propid: WIA_DPC_EXPOSURE_MODE;                            pszName: WIA_DPC_EXPOSURE_MODE_STR),
      (propid: WIA_DPC_EXPOSURE_COMP;                            pszName: WIA_DPC_EXPOSURE_COMP_STR),
      (propid: WIA_DPC_EXPOSURE_TIME;                            pszName: WIA_DPC_EXPOSURE_TIME_STR),
      (propid: WIA_DPC_FNUMBER;                                  pszName: WIA_DPC_FNUMBER_STR),
      (propid: WIA_DPC_FLASH_MODE;                               pszName: WIA_DPC_FLASH_MODE_STR),
      (propid: WIA_DPC_FOCUS_MODE;                               pszName: WIA_DPC_FOCUS_MODE_STR),
      (propid: WIA_DPC_FOCUS_MANUAL_DIST;                        pszName: WIA_DPC_FOCUS_MANUAL_DIST_STR),
      (propid: WIA_DPC_ZOOM_POSITION;                            pszName: WIA_DPC_ZOOM_POSITION_STR),
      (propid: WIA_DPC_PAN_POSITION;                             pszName: WIA_DPC_PAN_POSITION_STR),
      (propid: WIA_DPC_TILT_POSITION;                            pszName: WIA_DPC_TILT_POSITION_STR),
      (propid: WIA_DPC_TIMER_MODE;                               pszName: WIA_DPC_TIMER_MODE_STR),
      (propid: WIA_DPC_TIMER_VALUE;                              pszName: WIA_DPC_TIMER_VALUE_STR),
      (propid: WIA_DPC_POWER_MODE;                               pszName: WIA_DPC_POWER_MODE_STR),
      (propid: WIA_DPC_BATTERY_STATUS;                           pszName: WIA_DPC_BATTERY_STATUS_STR),
      (propid: WIA_DPC_DIMENSION;                                pszName: WIA_DPC_DIMENSION_STR),
      (propid: WIA_DPS_HORIZONTAL_BED_SIZE;                      pszName: WIA_DPS_HORIZONTAL_BED_SIZE_STR),
      (propid: WIA_DPS_VERTICAL_BED_SIZE;                        pszName: WIA_DPS_VERTICAL_BED_SIZE_STR),
      (propid: WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE;               pszName: WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE_STR),
      (propid: WIA_DPS_VERTICAL_SHEET_FEED_SIZE;                 pszName: WIA_DPS_VERTICAL_SHEET_FEED_SIZE_STR),
      (propid: WIA_DPS_SHEET_FEEDER_REGISTRATION;                pszName: WIA_DPS_SHEET_FEEDER_REGISTRATION_STR),
      (propid: WIA_DPS_HORIZONTAL_BED_REGISTRATION;              pszName: WIA_DPS_HORIZONTAL_BED_REGISTRATION_STR),
      (propid: WIA_DPS_VERTICAL_BED_REGISTRATION;                pszName: WIA_DPS_VERTICAL_BED_REGISTRATION_STR),
      (propid: WIA_DPS_PLATEN_COLOR;                             pszName: WIA_DPS_PLATEN_COLOR_STR),
      (propid: WIA_DPS_PAD_COLOR;                                pszName: WIA_DPS_PAD_COLOR_STR),
      (propid: WIA_DPS_FILTER_SELECT;                            pszName: WIA_DPS_FILTER_SELECT_STR),
      (propid: WIA_DPS_DITHER_SELECT;                            pszName: WIA_DPS_DITHER_SELECT_STR),
      (propid: WIA_DPS_DITHER_PATTERN_DATA;                      pszName: WIA_DPS_DITHER_PATTERN_DATA_STR),
      (propid: WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES;           pszName: WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES_STR),
      (propid: WIA_DPS_DOCUMENT_HANDLING_STATUS;                 pszName: WIA_DPS_DOCUMENT_HANDLING_STATUS_STR),
      (propid: WIA_DPS_DOCUMENT_HANDLING_SELECT;                 pszName: WIA_DPS_DOCUMENT_HANDLING_SELECT_STR),
      (propid: WIA_DPS_DOCUMENT_HANDLING_CAPACITY;               pszName: WIA_DPS_DOCUMENT_HANDLING_CAPACITY_STR),
      (propid: WIA_DPS_OPTICAL_XRES;                             pszName: WIA_DPS_OPTICAL_XRES_STR),
      (propid: WIA_DPS_OPTICAL_YRES;                             pszName: WIA_DPS_OPTICAL_YRES_STR),
      (propid: WIA_DPS_ENDORSER_CHARACTERS;                      pszName: WIA_DPS_ENDORSER_CHARACTERS_STR),
      (propid: WIA_DPS_ENDORSER_STRING;                          pszName: WIA_DPS_ENDORSER_STRING_STR),
      (propid: WIA_DPS_SCAN_AHEAD_PAGES;                         pszName: WIA_DPS_SCAN_AHEAD_PAGES_STR),
      (propid: WIA_DPS_MAX_SCAN_TIME;                            pszName: WIA_DPS_MAX_SCAN_TIME_STR),
      (propid: WIA_DPS_PAGES;                                    pszName: WIA_DPS_PAGES_STR),
      (propid: WIA_DPS_PAGE_SIZE;                                pszName: WIA_DPS_PAGE_SIZE_STR),
      (propid: WIA_DPS_PAGE_WIDTH;                               pszName: WIA_DPS_PAGE_WIDTH_STR),
      (propid: WIA_DPS_PAGE_HEIGHT;                              pszName: WIA_DPS_PAGE_HEIGHT_STR),
      (propid: WIA_DPS_PREVIEW;                                  pszName: WIA_DPS_PREVIEW_STR),
      (propid: WIA_DPS_TRANSPARENCY;                             pszName: WIA_DPS_TRANSPARENCY_STR),
      (propid: WIA_DPS_TRANSPARENCY_SELECT;                      pszName: WIA_DPS_TRANSPARENCY_SELECT_STR),
      (propid: WIA_DPS_SHOW_PREVIEW_CONTROL;                     pszName: WIA_DPS_SHOW_PREVIEW_CONTROL_STR),
      (propid: WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE;           pszName: WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE_STR),
      (propid: WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE;             pszName: WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE_STR),
      (propid: WIA_DPS_USER_NAME;                                pszName: WIA_DPS_USER_NAME_STR),
      (propid: WIA_DPV_LAST_PICTURE_TAKEN;                       pszName: WIA_DPV_LAST_PICTURE_TAKEN_STR),
      (propid: WIA_DPV_IMAGES_DIRECTORY;                         pszName: WIA_DPV_IMAGES_DIRECTORY_STR),
      (propid: WIA_DPV_DSHOW_DEVICE_PATH;                        pszName: WIA_DPV_DSHOW_DEVICE_PATH_STR),
      (propid: WIA_DPF_MOUNT_POINT;                              pszName: WIA_DPF_MOUNT_POINT_STR),
      (propid: WIA_IPA_ITEM_NAME;                                pszName: WIA_IPA_ITEM_NAME_STR),
      (propid: WIA_IPA_FULL_ITEM_NAME;                           pszName: WIA_IPA_FULL_ITEM_NAME_STR),
      (propid: WIA_IPA_ITEM_TIME;                                pszName: WIA_IPA_ITEM_TIME_STR),
      (propid: WIA_IPA_ITEM_FLAGS;                               pszName: WIA_IPA_ITEM_FLAGS_STR),
      (propid: WIA_IPA_ACCESS_RIGHTS;                            pszName: WIA_IPA_ACCESS_RIGHTS_STR),
      (propid: WIA_IPA_DATATYPE;                                 pszName: WIA_IPA_DATATYPE_STR),
      (propid: WIA_IPA_DEPTH;                                    pszName: WIA_IPA_DEPTH_STR),
      (propid: WIA_IPA_PREFERRED_FORMAT;                         pszName: WIA_IPA_PREFERRED_FORMAT_STR),
      (propid: WIA_IPA_FORMAT;                                   pszName: WIA_IPA_FORMAT_STR),
      (propid: WIA_IPA_COMPRESSION;                              pszName: WIA_IPA_COMPRESSION_STR),
      (propid: WIA_IPA_TYMED;                                    pszName: WIA_IPA_TYMED_STR),
      (propid: WIA_IPA_CHANNELS_PER_PIXEL;                       pszName: WIA_IPA_CHANNELS_PER_PIXEL_STR),
      (propid: WIA_IPA_BITS_PER_CHANNEL;                         pszName: WIA_IPA_BITS_PER_CHANNEL_STR),
      (propid: WIA_IPA_PLANAR;                                   pszName: WIA_IPA_PLANAR_STR),
      (propid: WIA_IPA_PIXELS_PER_LINE;                          pszName: WIA_IPA_PIXELS_PER_LINE_STR),
      (propid: WIA_IPA_BYTES_PER_LINE;                           pszName: WIA_IPA_BYTES_PER_LINE_STR),
      (propid: WIA_IPA_NUMBER_OF_LINES;                          pszName: WIA_IPA_NUMBER_OF_LINES_STR),
      (propid: WIA_IPA_GAMMA_CURVES;                             pszName: WIA_IPA_GAMMA_CURVES_STR),
      (propid: WIA_IPA_ITEM_SIZE;                                pszName: WIA_IPA_ITEM_SIZE_STR),
      (propid: WIA_IPA_COLOR_PROFILE;                            pszName: WIA_IPA_COLOR_PROFILE_STR),
      (propid: WIA_IPA_MIN_BUFFER_SIZE;                          pszName: WIA_IPA_MIN_BUFFER_SIZE_STR),
      (propid: WIA_IPA_REGION_TYPE;                              pszName: WIA_IPA_REGION_TYPE_STR),
      (propid: WIA_IPA_ICM_PROFILE_NAME;                         pszName: WIA_IPA_ICM_PROFILE_NAME_STR),
      (propid: WIA_IPA_APP_COLOR_MAPPING;                        pszName: WIA_IPA_APP_COLOR_MAPPING_STR),
      (propid: WIA_IPA_PROP_STREAM_COMPAT_ID;                    pszName: WIA_IPA_PROP_STREAM_COMPAT_ID_STR),
      (propid: WIA_IPA_FILENAME_EXTENSION;                       pszName: WIA_IPA_FILENAME_EXTENSION_STR),
      (propid: WIA_IPA_SUPPRESS_PROPERTY_PAGE;                   pszName: WIA_IPA_SUPPRESS_PROPERTY_PAGE_STR),
      (propid: WIA_IPC_THUMBNAIL;                                pszName: WIA_IPC_THUMBNAIL_STR),
      (propid: WIA_IPC_THUMB_WIDTH;                              pszName: WIA_IPC_THUMB_WIDTH_STR),
      (propid: WIA_IPC_THUMB_HEIGHT;                             pszName: WIA_IPC_THUMB_HEIGHT_STR),
      (propid: WIA_IPC_AUDIO_AVAILABLE;                          pszName: WIA_IPC_AUDIO_AVAILABLE_STR),
      (propid: WIA_IPC_AUDIO_DATA_FORMAT;                        pszName: WIA_IPC_AUDIO_DATA_FORMAT_STR),
      (propid: WIA_IPC_AUDIO_DATA;                               pszName: WIA_IPC_AUDIO_DATA_STR),
      (propid: WIA_IPC_NUM_PICT_PER_ROW;                         pszName: WIA_IPC_NUM_PICT_PER_ROW_STR),
      (propid: WIA_IPC_SEQUENCE;                                 pszName: WIA_IPC_SEQUENCE_STR),
      (propid: WIA_IPC_TIMEDELAY;                                pszName: WIA_IPC_TIMEDELAY_STR),
      (propid: WIA_IPS_CUR_INTENT;                               pszName: WIA_IPS_CUR_INTENT_STR),
      (propid: WIA_IPS_XRES;                                     pszName: WIA_IPS_XRES_STR),
      (propid: WIA_IPS_YRES;                                     pszName: WIA_IPS_YRES_STR),
      (propid: WIA_IPS_XPOS;                                     pszName: WIA_IPS_XPOS_STR),
      (propid: WIA_IPS_YPOS;                                     pszName: WIA_IPS_YPOS_STR),
      (propid: WIA_IPS_XEXTENT;                                  pszName: WIA_IPS_XEXTENT_STR),
      (propid: WIA_IPS_YEXTENT;                                  pszName: WIA_IPS_YEXTENT_STR),
      (propid: WIA_IPS_PHOTOMETRIC_INTERP;                       pszName: WIA_IPS_PHOTOMETRIC_INTERP_STR),
      (propid: WIA_IPS_BRIGHTNESS;                               pszName: WIA_IPS_BRIGHTNESS_STR),
      (propid: WIA_IPS_CONTRAST;                                 pszName: WIA_IPS_CONTRAST_STR),
      (propid: WIA_IPS_ORIENTATION;                              pszName: WIA_IPS_ORIENTATION_STR),
      (propid: WIA_IPS_ROTATION;                                 pszName: WIA_IPS_ROTATION_STR),
      (propid: WIA_IPS_MIRROR;                                   pszName: WIA_IPS_MIRROR_STR),
      (propid: WIA_IPS_THRESHOLD;                                pszName: WIA_IPS_THRESHOLD_STR),
      (propid: WIA_IPS_INVERT;                                   pszName: WIA_IPS_INVERT_STR),
      (propid: WIA_IPS_WARM_UP_TIME;                             pszName: WIA_IPS_WARM_UP_TIME_STR),
  //#if (_WIN32_WINNT >= 0x0600)
      (propid: WIA_IPA_ITEM_CATEGORY;                            pszName: WIA_IPA_ITEM_CATEGORY_STR),
      (propid: WIA_IPA_RAW_BITS_PER_CHANNEL;                     pszName: WIA_IPA_RAW_BITS_PER_CHANNEL_STR),
      (propid: WIA_IPS_DESKEW_X;                                 pszName: WIA_IPS_DESKEW_X_STR),
      (propid: WIA_IPS_DESKEW_Y;                                 pszName: WIA_IPS_DESKEW_Y_STR),
      (propid: WIA_IPS_SEGMENTATION;                             pszName: WIA_IPS_SEGMENTATION_STR),
      (propid: WIA_IPS_MAX_HORIZONTAL_SIZE;                      pszName: WIA_IPS_MAX_HORIZONTAL_SIZE_STR),
      (propid: WIA_IPS_MAX_VERTICAL_SIZE;                        pszName: WIA_IPS_MAX_VERTICAL_SIZE_STR),
      (propid: WIA_IPS_MIN_HORIZONTAL_SIZE;                      pszName: WIA_IPS_MIN_HORIZONTAL_SIZE_STR),
      (propid: WIA_IPS_MIN_VERTICAL_SIZE;                        pszName: WIA_IPS_MIN_VERTICAL_SIZE_STR),
      (propid: WIA_IPS_SHEET_FEEDER_REGISTRATION;                pszName: WIA_IPS_SHEET_FEEDER_REGISTRATION_STR),
      (propid: WIA_IPS_DOCUMENT_HANDLING_SELECT;                 pszName: WIA_IPS_DOCUMENT_HANDLING_SELECT_STR),
      (propid: WIA_IPS_OPTICAL_XRES;                             pszName: WIA_IPS_OPTICAL_XRES_STR),
      (propid: WIA_IPS_OPTICAL_YRES;                             pszName: WIA_IPS_OPTICAL_YRES_STR),
      (propid: WIA_IPS_PAGES;                                    pszName: WIA_IPS_PAGES_STR),
      (propid: WIA_IPS_PAGE_SIZE;                                pszName: WIA_IPS_PAGE_SIZE_STR),
      (propid: WIA_IPS_PAGE_WIDTH;                               pszName: WIA_IPS_PAGE_WIDTH_STR),
      (propid: WIA_IPS_PAGE_HEIGHT;                              pszName: WIA_IPS_PAGE_HEIGHT_STR),
      (propid: WIA_IPS_PREVIEW;                                  pszName: WIA_IPS_PREVIEW_STR),
      (propid: WIA_IPS_SHOW_PREVIEW_CONTROL;                     pszName: WIA_IPS_SHOW_PREVIEW_CONTROL_STR),
      (propid: WIA_IPS_TRANSFER_CAPABILITIES;                    pszName: WIA_IPS_TRANSFER_CAPABILITIES_STR),
      (propid: WIA_IPS_FILM_SCAN_MODE;                           pszName: WIA_IPS_FILM_SCAN_MODE_STR),
      (propid: WIA_IPS_LAMP;                                     pszName: WIA_IPS_LAMP_STR),
      (propid: WIA_IPS_LAMP_AUTO_OFF;                            pszName: WIA_IPS_LAMP_AUTO_OFF_STR),
      (propid: WIA_IPS_AUTO_DESKEW;                              pszName: WIA_IPS_AUTO_DESKEW_STR),
      (propid: WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION;             pszName: WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION_STR),
      (propid: WIA_IPS_PREVIEW_TYPE;                             pszName: WIA_IPS_PREVIEW_TYPE_STR),
      (propid: WIA_IPS_XSCALING;                                 pszName: WIA_IPS_XSCALING_STR),
      (propid: WIA_IPS_YSCALING;                                 pszName: WIA_IPS_YSCALING_STR),
      (propid: WIA_IPA_UPLOAD_ITEM_SIZE;                         pszName: WIA_IPA_UPLOAD_ITEM_SIZE_STR),
      (propid: WIA_IPA_ITEMS_STORED;                             pszName: WIA_IPA_ITEMS_STORED_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER;                         pszName: WIA_IPS_PRINTER_ENDORSER_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_ORDER;                   pszName: WIA_IPS_PRINTER_ENDORSER_ORDER_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_COUNTER;                 pszName: WIA_IPS_PRINTER_ENDORSER_COUNTER_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_STEP;                    pszName: WIA_IPS_PRINTER_ENDORSER_STEP_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_XOFFSET;                 pszName: WIA_IPS_PRINTER_ENDORSER_XOFFSET_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_YOFFSET;                 pszName: WIA_IPS_PRINTER_ENDORSER_YOFFSET_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_NUM_LINES;               pszName: WIA_IPS_PRINTER_ENDORSER_NUM_LINES_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_STRING;                  pszName: WIA_IPS_PRINTER_ENDORSER_STRING_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_VALID_CHARACTERS;        pszName: WIA_IPS_PRINTER_ENDORSER_VALID_CHARACTERS_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS; pszName: WIA_IPS_PRINTER_ENDORSER_VALID_FORMAT_SPECIFIERS_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_TEXT_UPLOAD;             pszName: WIA_IPS_PRINTER_ENDORSER_TEXT_UPLOAD_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_TEXT_DOWNLOAD;           pszName: WIA_IPS_PRINTER_ENDORSER_TEXT_DOWNLOAD_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS;                pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_POSITION;       pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_POSITION_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_WIDTH;      pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_WIDTH_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_WIDTH;      pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_WIDTH_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_HEIGHT;     pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MIN_HEIGHT_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_HEIGHT;     pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_MAX_HEIGHT_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_UPLOAD;         pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_UPLOAD_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_DOWNLOAD;       pszName: WIA_IPS_PRINTER_ENDORSER_GRAPHICS_DOWNLOAD_STR),
      (propid: WIA_IPS_BARCODE_READER;                           pszName: WIA_IPS_BARCODE_READER_STR),
      (propid: WIA_IPS_MAXIMUM_BARCODES_PER_PAGE;                pszName: WIA_IPS_MAXIMUM_BARCODES_PER_PAGE_STR),
      (propid: WIA_IPS_BARCODE_SEARCH_DIRECTION;                 pszName: WIA_IPS_BARCODE_SEARCH_DIRECTION_STR),
      (propid: WIA_IPS_MAXIMUM_BARCODE_SEARCH_RETRIES;           pszName: WIA_IPS_MAXIMUM_BARCODE_SEARCH_RETRIES_STR),
      (propid: WIA_IPS_BARCODE_SEARCH_TIMEOUT;                   pszName: WIA_IPS_BARCODE_SEARCH_TIMEOUT_STR),
      (propid: WIA_IPS_SUPPORTED_BARCODE_TYPES;                  pszName: WIA_IPS_SUPPORTED_BARCODE_TYPES_STR),
      (propid: WIA_IPS_ENABLED_BARCODE_TYPES;                    pszName: WIA_IPS_ENABLED_BARCODE_TYPES_STR),
      (propid: WIA_IPS_PATCH_CODE_READER;                        pszName: WIA_IPS_PATCH_CODE_READER_STR),
      (propid: WIA_IPS_SUPPORTED_PATCH_CODE_TYPES;               pszName: WIA_IPS_SUPPORTED_PATCH_CODE_TYPES_STR),
      (propid: WIA_IPS_ENABLED_PATCH_CODE_TYPES;                 pszName: WIA_IPS_ENABLED_PATCH_CODE_TYPES_STR),
      (propid: WIA_IPS_MICR_READER;                              pszName: WIA_IPS_MICR_READER_STR),
      (propid: WIA_IPS_JOB_SEPARATORS;                           pszName: WIA_IPS_JOB_SEPARATORS_STR),
      (propid: WIA_IPS_LONG_DOCUMENT;                            pszName: WIA_IPS_LONG_DOCUMENT_STR),
      (propid: WIA_IPS_BLANK_PAGES;                              pszName: WIA_IPS_BLANK_PAGES_STR),
      (propid: WIA_IPS_MULTI_FEED;                               pszName: WIA_IPS_MULTI_FEED_STR),
      (propid: WIA_IPS_MULTI_FEED_SENSITIVITY;                   pszName: WIA_IPS_MULTI_FEED_SENSITIVITY_STR),
      (propid: WIA_IPS_AUTO_CROP;                                pszName: WIA_IPS_AUTO_CROP_STR),
      (propid: WIA_IPS_OVER_SCAN;                                pszName: WIA_IPS_OVER_SCAN_STR),
      (propid: WIA_IPS_OVER_SCAN_LEFT;                           pszName: WIA_IPS_OVER_SCAN_LEFT_STR),
      (propid: WIA_IPS_OVER_SCAN_RIGHT;                          pszName: WIA_IPS_OVER_SCAN_RIGHT_STR),
      (propid: WIA_IPS_OVER_SCAN_TOP;                            pszName: WIA_IPS_OVER_SCAN_TOP_STR),
      (propid: WIA_IPS_OVER_SCAN_BOTTOM;                         pszName: WIA_IPS_OVER_SCAN_BOTTOM_STR),
      (propid: WIA_IPS_COLOR_DROP;                               pszName: WIA_IPS_COLOR_DROP_STR),
      (propid: WIA_IPS_COLOR_DROP_RED;                           pszName: WIA_IPS_COLOR_DROP_RED_STR),
      (propid: WIA_IPS_COLOR_DROP_GREEN;                         pszName: WIA_IPS_COLOR_DROP_GREEN_STR),
      (propid: WIA_IPS_COLOR_DROP_BLUE;                          pszName: WIA_IPS_COLOR_DROP_BLUE_STR),
      (propid: WIA_IPS_SCAN_AHEAD;                               pszName: WIA_IPS_SCAN_AHEAD_STR),
      (propid: WIA_IPS_SCAN_AHEAD_CAPACITY;                      pszName: WIA_IPS_SCAN_AHEAD_CAPACITY_STR),
      (propid: WIA_IPS_FEEDER_CONTROL;                           pszName: WIA_IPS_FEEDER_CONTROL_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_PADDING;                 pszName: WIA_IPS_PRINTER_ENDORSER_PADDING_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_FONT_TYPE;               pszName: WIA_IPS_PRINTER_ENDORSER_FONT_TYPE_STR),
      (propid: WIA_IPS_ALARM;                                    pszName: WIA_IPS_ALARM_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_INK;                     pszName: WIA_IPS_PRINTER_ENDORSER_INK_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_CHARACTER_ROTATION;      pszName: WIA_IPS_PRINTER_ENDORSER_CHARACTER_ROTATION_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_MAX_CHARACTERS;          pszName: WIA_IPS_PRINTER_ENDORSER_MAX_CHARACTERS_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_MAX_GRAPHICS;            pszName: WIA_IPS_PRINTER_ENDORSER_MAX_GRAPHICS_STR),
      (propid: WIA_IPS_PRINTER_ENDORSER_COUNTER_DIGITS;          pszName: WIA_IPS_PRINTER_ENDORSER_COUNTER_DIGITS_STR),
      (propid: WIA_IPS_COLOR_DROP_MULTI;                         pszName: WIA_IPS_COLOR_DROP_MULTI_STR),
      (propid: WIA_IPS_BLANK_PAGES_SENSITIVITY;                  pszName: WIA_IPS_BLANK_PAGES_SENSITIVITY_STR),
  //#endif
      (propid: 0; pszName: 'Not a standard WIA property')
  );

  {$else}
  var
     g_wiaPropIdToName : array[0..0] of WIA_PROPID_TO_NAME; //cvar;external;
  {$endif}


implementation

end.

