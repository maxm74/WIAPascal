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
    cbDataType: TComboBox;
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
  paperI: TWIAPaperSize;
  dataI: TWIADataType;
  i, cbSelected: Integer;

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

    { #todo 10 -oMaxM : In theory the selectable BitDepths depend on ImageType,
                        but WIA does not give me an error when I set for example
                        "Black and White" and BitDepht to 24 bit but I get a damaged Image.
                        Change automatically from the Form? }

    //Fill List of Image Bit Depth
    cbBitDepth.Clear;
    cbSelected :=0;
    for i:=0 to Length(BitDepthArray)-1 do
    begin
      if (BitDepthArray[i] = 0)
      then cbBitDepth.Items.AddObject('Auto', nil)
      else cbBitDepth.Items.AddObject(IntToStr(BitDepthArray[i])+' Bit', TObject(PtrUInt(BitDepthArray[i])));

      case initItemValues of
      initDefault: begin if (BitDepthArray[i] = BitDepthDefault) then cbSelected :=cbBitDepth.Items.Count-1; end;
      initParams:  begin if (BitDepthArray[i] = AParams.BitDepth) then cbSelected :=cbBitDepth.Items.Count-1; end;
      initCurrent: begin if (BitDepthArray[i] = BitDepthCurrent) then cbSelected :=cbBitDepth.Items.Count-1; end;
      end;
    end;
    cbBitDepth.ItemIndex:=cbSelected;

    //Fill List of Image Data Type
    cbDataType.Clear;
    cbSelected :=0;
    for dataI in DataTypeSet do
    begin
      Case dataI of
      wdtAUTO: cbDataType.Items.AddObject('Auto type', TObject(PtrUInt(wdtAUTO)));
      wdtDITHER, wdtCOLOR_DITHER: begin end;
      else begin
             cbDataType.Items.AddObject(WIADataTypeStr[dataI], TObject(PtrUInt(dataI)));

             case initItemValues of
             initDefault: begin if (dataI = DataTypeDefault) then cbSelected :=cbDataType.Items.Count-1; end;
             initParams:  begin if (dataI = AParams.DataType) then cbSelected :=cbDataType.Items.Count-1; end;
             initCurrent: begin if (dataI = DataTypeCurrent) then cbSelected :=cbDataType.Items.Count-1; end;
             end;
           end;
      end;
    end;
    cbDataType.ItemIndex:=cbSelected;

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

      if (cbBitDepth.ItemIndex>-1)
      then AParams.BitDepth:=PtrUInt(cbBitDepth.Items.Objects[cbBitDepth.ItemIndex]);

      if (cbDataType.ItemIndex>-1)
      then AParams.DataType:=TWIADataType(PtrUInt(cbDataType.Items.Objects[cbDataType.ItemIndex]));

      if (cbResolution.ItemIndex>-1)
      then AParams.Resolution:=WIACap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

      AParams.Contrast:=edContrast.Value;
      AParams.Brightness:=edBrightness.Value;
    end;
  end;
end;

end.

