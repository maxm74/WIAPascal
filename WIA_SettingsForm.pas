(******************************************************************************
*                FreePascal \ Delphi WIA Implementation                       *
*                                                                             *
*  FILE: WIA_SettingsForm.pas                                                 *
*                                                                             *
*  VERSION:     0.0.2                                                         *
*                                                                             *
*  DESCRIPTION:                                                               *
*    WIA Property Settings for Device Dialog.                                 *
*                                                                             *
*******************************************************************************
*                                                                             *
*  (c) 2024 Massimo Magnano                                                   *
*                                                                             *
*  See changelog.txt for Change Log                                           *
*                                                                             *
*******************************************************************************)

unit WIA_SettingsForm;

{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  ComCtrls, StdCtrls, Spin,
  WIA,
  ImgList {$ifndef fpc}, ImageList{$endif};

type
  TInitialItemValues = (initDefault, initParams, initCurrent);

  { TWIASettingsSource }
  TWIASettingsSource = class(TForm)
    btCancel: TBitBtn;
    btOrientation: TSpeedButton;
    btRefreshUndo: TBitBtn;
    btOk: TBitBtn;
    btRefreshCurrent: TBitBtn;
    btRefreshDefault: TBitBtn;
    cbBitDepth: TComboBox;
    cbSourceItem: TComboBox;
    cbPaperSize: TComboBox;
    cbPixelType: TComboBox;
    cbResolution: TComboBox;
    cbUseNativeUI: TCheckBox;
    edBrightness: TSpinEdit;
    edContrast: TSpinEdit;
    imgList: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    PageSourceTypes: TPageControl;
    Panel1: TPanel;
    panelButtons: TPanel;
    tbSource_Scanner: TTabSheet;
    trBrightness: TTrackBar;
    trContrast: TTrackBar;
    procedure edBrightnessChange(Sender: TObject);
    procedure edContrastChange(Sender: TObject);
    procedure trBrightnessChange(Sender: TObject);
    procedure trContrastChange(Sender: TObject);
  private
    WIASource: TWIADevice;
    WIASelectedItemIndex: Integer;
    WIACap: TWIAParamsCapabilities;
    WIAParams: TWIAParams;

  public
     class function Execute(AWIASource: TWIADevice;
                            var ASelectedItemIndex: Integer; { #note -oMaxM : Eventualmente Filtri per quali Item Mostrare }
                            initItemValues: TInitialItemValues; var AParams: TWIAParams): Boolean;
  end;

var
  WIASettingsSource: TWIASettingsSource=nil;

implementation

{$ifdef FPC}
  {$R *.lfm}
{$else}
  {$R *.dfm}
{$endif}

uses WIA_PaperSizes;

{ TWIASettingsSource }

procedure TWIASettingsSource.trBrightnessChange(Sender: TObject);
begin
  edBrightness.Value:=trBrightness.Position;
end;

procedure TWIASettingsSource.edBrightnessChange(Sender: TObject);
begin
  trBrightness.Position:=edBrightness.Value;
end;

procedure TWIASettingsSource.edContrastChange(Sender: TObject);
begin
  trContrast.Position:=edContrast.Value;
end;

procedure TWIASettingsSource.trContrastChange(Sender: TObject);
begin
  edContrast.Value:=trContrast.Position;
end;

class function TWIASettingsSource.Execute(AWIASource: TWIADevice; var ASelectedItemIndex: Integer;
  initItemValues: TInitialItemValues; var AParams: TWIAParams): Boolean;
var
  paperI: TWIAPaperSize;
//  pixelI:TTwainPixelType;
  i, cbSelected: Integer;
  (*  TwainSource:TTwainSource;
    bitCurrent: Integer;
    paperCurrent: TTwainPaperSize;
    pixelCurrent:TTwainPixelType;
    resolutionCurrent:Single;
  *)
begin
  if (WIASettingsSource=nil)
  then WIASettingsSource :=TWIASettingsSource.Create(nil);

  with WIASettingsSource do
  begin
    WIASource:= AWIASource;
    WIASelectedItemIndex:= ASelectedItemIndex;
    WIAParams:= AParams;

    { #note -oMaxM : Eventualmente Muovere nella Classe TWIADevice }
    //Get current Selected Item Default Values
    with WIACap do
    begin
      WIASource.GetPaperSizeSet(PaperSizeCurrent, PaperSizeDefault, PaperSizeSet);
    end;

(*    cbSourceItem.Clear;
    if (pfFlatbed in TwainCap.PaperFeedingSet) then cbSourceItem.Items.AddObject('Flatbed', TObject(PtrUInt(pfFlatbed)));
    if (pfFeeder in TwainCap.PaperFeedingSet) then cbSourceItem.Items.AddObject('Feeder', TObject(PtrUInt(pfFeeder)));
    cbSourceItem.ItemIndex:=cbSourceItem.Items.IndexOfObject(TObject(PtrUInt(AParams.PaperFeed)));
*)
    //Fill List of Papers
    cbPaperSize.Clear;
    cbSelected :=0;
    cbPaperSize.Items.AddObject('Full size', TObject(PtrUInt(wpsMAX)));
    for paperI in WIACap.PaperSizeSet do
    begin
      Case paperI of
      wpsCUSTOM: cbPaperSize.Items.AddObject('Custom size', TObject(PtrUInt(wpsCUSTOM)));
      wpsAUTO: cbPaperSize.Items.AddObject('Auto size', TObject(PtrUInt(wpsAUTO)));
      else cbPaperSize.Items.AddObject(PaperSizesWIA[paperI].name+
             ' ('+FloatToStrF(PaperSizesWIA[paperI].w, ffFixed, 15, 2)+' x '+
                  FloatToStrF(PaperSizesWIA[paperI].h, ffFixed, 15, 2)+')',
             TObject(PtrUInt(paperI)));
      end;

      case initItemValues of
      initDefault: begin if (paperI = WIACap.PaperSizeDefault) then cbSelected :=cbPaperSize.Items.Count-1; end;
      initParams:  begin if (paperI = AParams.PaperSize) then cbSelected :=cbPaperSize.Items.Count-1; end;
      initCurrent: begin if (paperI = WIACap.PaperSizeCurrent) then cbSelected :=cbPaperSize.Items.Count-1; end;
      end;
    end;
    cbPaperSize.ItemIndex:=cbSelected;
(*
    //Fill List of Bit Depth
    cbBitDepth.Clear;
    cbSelected :=0;
    for i:=0 to TwainCap.BitDepthArraySize-1 do
    begin
      cbBitDepth.Items.AddObject(IntToStr(TwainCap.BitDepthArray[i])+' Bit', TObject(PtrUInt(TwainCap.BitDepthArray[i])));

      if useDeviceDefault
      then begin if (TwainCap.BitDepthArray[i] = TwainCap.BitDepthDefault) then cbSelected :=cbBitDepth.Items.Count-1; end
      else begin if (TwainCap.BitDepthArray[i] = AParams.BitDepth) then cbSelected :=cbBitDepth.Items.Count-1; end;
    end;
    cbBitDepth.ItemIndex:=cbSelected;

    //Fill List of Pixel Type
    cbPixelType.Clear;
    cbSelected :=0;
    for pixelI in TwainCap.PixelType do
    begin
      cbPixelType.Items.AddObject(TwainPixelTypes[pixelI], TObject(PtrUInt(pixelI)));

      if useDeviceDefault
      then begin if (pixelI = TwainCap.PixelTypeDefault) then cbSelected :=cbPixelType.Items.Count-1; end
      else begin if (pixelI = AParams.PixelType) then cbSelected :=cbPixelType.Items.Count-1; end;
    end;
    cbPixelType.ItemIndex:=cbSelected;

    //Fill List of Resolution (Y Resolution=X Resolution)
    cbResolution.Clear;
    cbSelected :=0;
    for i:=0 to TwainCap.ResolutionArraySize-1 do
    begin
      cbResolution.Items.AddObject(FloatToStr(TwainCap.ResolutionArray[i]), TObject(PtrUInt(i)));

      if useDeviceDefault
      then begin if (TwainCap.ResolutionArray[i] = TwainCap.ResolutionDefault) then cbSelected :=cbResolution.Items.Count-1; end
      else begin if (TwainCap.ResolutionArray[i] = AParams.Resolution) then cbSelected :=cbResolution.Items.Count-1; end;
    end;
    cbResolution.ItemIndex:=cbSelected;

    { #todo -oMaxM 2 : is an Single or an Integer? }
    trContrast.Position:=Trunc(AParams.Contrast);
    edContrast.Value:=Trunc(AParams.Contrast);
    trBrightness.Position:=Trunc(AParams.Brightness);
    edBrightness.Value:=Trunc(AParams.Brightness);
*)
    Result := (ShowModal=mrOk);

    if Result then
    begin
(*
      //Fill AParams with new values
      if (cbSourceItem.ItemIndex>-1)
      then AParams.PaperFeed:=TTwainPaperFeeding(PtrUInt(cbSourceItem.Items.Objects[cbSourceItem.ItemIndex]));
      { #todo -oMaxM : else Predefined Value }
*)
      if (cbPaperSize.ItemIndex>-1)
      then AParams.PaperSize:=TWIAPaperSize(PtrUInt(cbPaperSize.Items.Objects[cbPaperSize.ItemIndex]));
(*
      if (cbBitDepth.ItemIndex>-1)
      then AParams.BitDepth:=PtrUInt(cbBitDepth.Items.Objects[cbBitDepth.ItemIndex]);

      if (cbPixelType.ItemIndex>-1)
      then AParams.PixelType:=TTwainPixelType(PtrUInt(cbPixelType.Items.Objects[cbPixelType.ItemIndex]));

      if (cbResolution.ItemIndex>-1)
      then AParams.Resolution:=TwainCap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

      AParams.Contrast:=edContrast.Value;
      AParams.Brightness:=edBrightness.Value;
*)
    end;
  end;
end;

end.

