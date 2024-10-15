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
  WIA, WIA_PaperSizes,
  ImgList {$ifndef fpc}, ImageList{$endif};

const
  //False to display then measurement in Inchs
  WIASettings_Unit_cm: Boolean = True; //fucks inch

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
    edResolution: TSpinEdit;
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
    trResolution: TTrackBar;
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
                            var ASelectedItemIndex: Integer; { #todo -oMaxM : Possibly Filters for which Items Kinds to Show }
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

uses WiaDef;

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
  pFlags: TWIAPropertyFlags;
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

  with WIASettingsSource, WIACap do
  begin
    WIASource:= AWIASource;
    WIASelectedItemIndex:= ASelectedItemIndex;
    WIAParams:= AParams;

    { #note -oMaxM : Eventualmente Muovere nella Classe TWIADevice }
    //Get current Selected Item Default Values
    Result:= WIASource.GetParamsCapabilities(WIACap);
    if not(Result) then raise Exception.Create('Cannot Get Capabilities for Device'#13#10+WIASource.Name);

(*    cbSourceItem.Clear;
    if (pfFlatbed in TwainCap.PaperFeedingSet) then cbSourceItem.Items.AddObject('Flatbed', TObject(PtrUInt(pfFlatbed)));
    if (pfFeeder in TwainCap.PaperFeedingSet) then cbSourceItem.Items.AddObject('Feeder', TObject(PtrUInt(pfFeeder)));
    cbSourceItem.ItemIndex:=cbSourceItem.Items.IndexOfObject(TObject(PtrUInt(AParams.PaperFeed)));
*)
    //Fill List of Papers
    cbPaperSize.Clear;
    cbSelected :=0;
    cbPaperSize.Items.AddObject('Full size', TObject(PtrUInt(wpsMAX)));
    for paperI in PaperSizeSet do
    begin
      Case paperI of
      wpsMAX:begin end;
      wpsCUSTOM: cbPaperSize.Items.AddObject('Custom size', TObject(PtrUInt(wpsCUSTOM)));
      wpsAUTO: cbPaperSize.Items.AddObject('Auto size', TObject(PtrUInt(wpsAUTO)));
      else begin
             if WIASettings_Unit_cm
             then cbPaperSize.Items.AddObject(PaperSizesWIA[paperI].name+' ('+
                                              THInchToCmStr(PaperSizesWIA[paperI].w)+' x '+
                                              THInchToCmStr(PaperSizesWIA[paperI].h)+' cm)',
                                              TObject(PtrUInt(paperI)))
             else cbPaperSize.Items.AddObject(PaperSizesWIA[paperI].name+' ('+
                                              THInchToInchStr(PaperSizesWIA[paperI].w)+' x '+
                                              THInchToInchStr(PaperSizesWIA[paperI].h)+' in)',
                                              TObject(PtrUInt(paperI)))
           end;
      end;

      case initItemValues of
      initDefault: begin if (paperI = PaperSizeDefault) then cbSelected :=cbPaperSize.Items.Count-1; end;
      initParams:  begin if (paperI = AParams.PaperSize) then cbSelected :=cbPaperSize.Items.Count-1; end;
      initCurrent: begin if (paperI = PaperSizeCurrent) then cbSelected :=cbPaperSize.Items.Count-1; end;
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
*)
    if ResolutionRange
    then begin
           trResolution.Visible:= True; edResolution.Visible:= True;
           cbResolution.Visible:= False;

           trResolution.Min:= ResolutionArray[WIA_RANGE_MIN];
           trResolution.Max:= ResolutionArray[WIA_RANGE_MAX];
           trResolution.LineSize:= ResolutionArray[WIA_RANGE_STEP];
           case initItemValues of
           initDefault: trResolution.Position:= ResolutionDefault;
           initParams:  trResolution.Position:= AParams.Resolution;
           initCurrent: trResolution.Position:= ResolutionCurrent;
           end;
           edResolution.MinValue:= ResolutionArray[WIA_RANGE_MIN];
           edResolution.MaxValue:= ResolutionArray[WIA_RANGE_MAX];
           edResolution.Increment:= ResolutionArray[WIA_RANGE_STEP];
           edResolution.Value:= trResolution.Position;
         end
    else begin
           trResolution.Visible:= False; edResolution.Visible:= False;
           cbResolution.Visible:= True;

           //Fill List of Resolution (Y Resolution=X Resolution)
           cbResolution.Clear;
           cbSelected :=0;
           for i:=0 to Length(ResolutionArray)-1 do
           begin
             cbResolution.Items.AddObject(IntToStr(ResolutionArray[i]), TObject(PtrUInt(i)));

             case initItemValues of
             initDefault: begin if (ResolutionArray[i] = ResolutionDefault) then cbSelected :=cbResolution.Items.Count-1; end;
             initParams:  begin if (ResolutionArray[i] = AParams.Resolution) then cbSelected :=cbResolution.Items.Count-1; end;
             initCurrent: begin if (ResolutionArray[i] = ResolutionCurrent) then cbSelected :=cbResolution.Items.Count-1; end;
             end;
           end;
           cbResolution.ItemIndex:=cbSelected;
         end;

    //Brightness
    trBrightness.Min:= BrightnessMin;
    trBrightness.Max:= BrightnessMax;
    trBrightness.LineSize:= BrightnessStep;
    case initItemValues of
    initDefault: trBrightness.Position:= BrightnessDefault;
    initParams:  trBrightness.Position:= AParams.Brightness;
    initCurrent: trBrightness.Position:= BrightnessCurrent;
    end;
    edBrightness.MinValue:= BrightnessMin;
    edBrightness.MaxValue:= BrightnessMax;
    edBrightness.Increment:= BrightnessStep;
    edBrightness.Value:= trBrightness.Position;

    //Contrast
    trContrast.Min:= ContrastMin;
    trContrast.Max:= ContrastMax;
    trContrast.LineSize:= ContrastStep;
    case initItemValues of
    initDefault: trContrast.Position:= ContrastDefault;
    initParams:  trContrast.Position:= AParams.Contrast;
    initCurrent: trContrast.Position:= ContrastCurrent;
    end;
    edContrast.MinValue:= ContrastMin;
    edContrast.MaxValue:= ContrastMax;
    edContrast.Increment:= ContrastStep;
    edContrast.Value:= trContrast.Position;

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
*)
      if (cbResolution.ItemIndex>-1)
      then AParams.Resolution:=WIACap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

      AParams.Contrast:=edContrast.Value;
      AParams.Brightness:=edBrightness.Value;
    end;
  end;
end;

end.

