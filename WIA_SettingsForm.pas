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
    btContrast0: TSpeedButton;
    btContrastD: TSpeedButton;
    btResolutionD: TSpeedButton;
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
    btBrightness0: TSpeedButton;
    btBrightnessD: TSpeedButton;
    tbSource_Scanner: TTabSheet;
    trBrightness: TTrackBar;
    trResolution: TTrackBar;
    trContrast: TTrackBar;
    procedure bt0Click(Sender: TObject);
    procedure btDClick(Sender: TObject);
    procedure btOrientationClick(Sender: TObject);
    procedure cbSourceItemChange(Sender: TObject);
    procedure edBrightnessChange(Sender: TObject);
    procedure edContrastChange(Sender: TObject);
    procedure trBrightnessChange(Sender: TObject);
    procedure trContrastChange(Sender: TObject);
  private
    WIASource: TWIADevice;
    WIASelectedItemIndex: Integer;
    WIACap: TWIAParamsCapabilities;
    WIAParams: TWIAParams;
    WIAParamsArray: TArrayWIAParams;
    initItemValues,
    newInitItemValues: TInitialItemValues;

    procedure SelectCurrentItem(AIndex: Integer);
    procedure StoreCurrentItemParams;

  public
     class function Execute(AWIASource: TWIADevice;
                            var ASelectedItemIndex: Integer;
                            { #todo -oMaxM : Possibly Filters for which Items Kinds to Show? How manage AParams without Indexes? }
                            AInitItemValues: TInitialItemValues;
                            var AParamsArray: TArrayWIAParams;
                            NewParamsInitValues: TInitialItemValues=initDefault): Boolean;
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

procedure TWIASettingsSource.cbSourceItemChange(Sender: TObject);
begin
  Case MessageDlg('Apply Changes to Item '+WIASource.Items[WIASelectedItemIndex]^.Name+' of '+WIASource.Name,
                 mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
    mrYes: begin
            StoreCurrentItemParams;
            SelectCurrentItem(cbSourceItem.ItemIndex);
    end;
    mrNo: SelectCurrentItem(cbSourceItem.ItemIndex);
  end;
end;

procedure TWIASettingsSource.bt0Click(Sender: TObject);
begin
  Case (TSpeedButton(Sender).Tag) of
  1: edBrightness.Value:= 0;
  2: edContrast.Value:= 0;
  end;
end;

procedure TWIASettingsSource.btDClick(Sender: TObject);
var
   i: Integer;

begin
  with WIACap do
  Case (TSpeedButton(Sender).Tag) of
  1: edBrightness.Value:= BrightnessDefault;
  2: edContrast.Value:= ContrastDefault;
  3: begin
       if ResolutionRange
       then edResolution.Value:= ResolutionDefault
       else begin
             i:= cbResolution.Items.IndexOf(IntToStr(ResolutionDefault));
             if (i > -1) then cbResolution.ItemIndex:= i;
            end;
     end;
  end;
end;

procedure TWIASettingsSource.btOrientationClick(Sender: TObject);
begin
  if btOrientation.Down
  then begin btOrientation.ImageIndex:= 5; btOrientation.Hint:= 'Landscape'; end
  else begin btOrientation.ImageIndex:= 4; btOrientation.Hint:= 'Portrait'; end;
end;

procedure TWIASettingsSource.edContrastChange(Sender: TObject);
begin
  trContrast.Position:=edContrast.Value;
end;

procedure TWIASettingsSource.trContrastChange(Sender: TObject);
begin
  edContrast.Value:=trContrast.Position;
end;

procedure TWIASettingsSource.SelectCurrentItem(AIndex: Integer);
var
   AParams: TWIAParams;
   capRet: Boolean;
   paperI: TWIAPaperSize;
   dataI: TWIADataType;
   i, cbSelected: Integer;

begin
  AParams:= WiaParamsArray[AIndex];
  WiaSource.SelectedItemIndex:= AIndex;

  { #note 10 -oMaxM : Get Capabilities for all the Items in method Execute. WIACapArray}

  //Get current Selected Item Default Values
  capRet:= WIASource.GetParamsCapabilities(WIACap);
  if not(capRet) then raise Exception.Create('Cannot Get Capabilities for Source Item');

  with WIACap do
  begin
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

  //Set Landscape/Portrait Button
  case initItemValues of
  initDefault: btOrientation.Down:= (RotationDefault in [wrLandscape, wrRot270]); //270 is basically landscape
  initParams:  btOrientation.Down:= (AParams.Rotation in [wrLandscape, wrRot270]);
  initCurrent: btOrientation.Down:= (RotationCurrent in [wrLandscape, wrRot270]);
  end;
  if btOrientation.Down
  then begin btOrientation.ImageIndex:= 5; btOrientation.Hint:= 'Landscape'; end
  else begin btOrientation.ImageIndex:= 4; btOrientation.Hint:= 'Portrait'; end;

(*
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
*)
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
  end;

  WIAParams:= AParams;
  WIASelectedItemIndex:= AIndex;
end;

procedure TWIASettingsSource.StoreCurrentItemParams;
begin
  if (cbPaperSize.ItemIndex>-1)
  then WIAParams.PaperSize:=TWIAPaperSize(PtrUInt(cbPaperSize.Items.Objects[cbPaperSize.ItemIndex]));

  if btOrientation.Down
  then WIAParams.Rotation:= wrLandscape
  else WIAParams.Rotation:= wrPortrait;

  (*
  if (cbBitDepth.ItemIndex>-1)
  then AParams.BitDepth:=PtrUInt(cbBitDepth.Items.Objects[cbBitDepth.ItemIndex]);
  *)

  if (cbDataType.ItemIndex>-1)
  then WIAParams.DataType:=TWIADataType(PtrUInt(cbDataType.Items.Objects[cbDataType.ItemIndex]));

  if (cbResolution.ItemIndex>-1)
  then WIAParams.Resolution:=WIACap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

  WIAParams.Contrast:=edContrast.Value;
  WIAParams.Brightness:=edBrightness.Value;

  WIAParamsArray[WIASelectedItemIndex]:= WIAParams;
end;

class function TWIASettingsSource.Execute(AWIASource: TWIADevice; var ASelectedItemIndex: Integer;
                                          AInitItemValues: TInitialItemValues; var AParamsArray: TArrayWIAParams;
                                          NewParamsInitValues: TInitialItemValues): Boolean;
var
  i, itemCount: Integer;

begin
  Result:= False;
  if (AWIASource = nil) then exit;

  if (WIASettingsSource=nil)
  then WIASettingsSource :=TWIASettingsSource.Create(nil);

  with WIASettingsSource do
  try
    WIASource:= AWIASource;
    itemCount:= WIASource.ItemCount;
    initItemValues:= AInitItemValues;
    newInitItemValues:= NewParamsInitValues;

    //Do A Copy of Params Array so if the user cancels the Dialog we don't modify the starting array
    WIAParamsArray:= Copy(AParamsArray);

    { #note 10 -oMaxM : Get Capabilities for all the Items, WIACapArray}

    //If WiaSource Item Count is greater then our Array enlarge it
    if (Length(WIAParamsArray) < itemCount) then
    begin
      SetLength(WIAParamsArray, itemCount);
    end;

    //Fill List of Source Items
    cbSourceItem.Clear;
    for i:=0 to WIASource.ItemCount-1 do
      cbSourceItem.Items.AddObject(WIASource.Items[i]^.Name, TObject(PtrUInt(i)));

    try
      //Select the Initial Item to ASelectedItemIndex
      try
         WIASource.SelectedItemIndex:= ASelectedItemIndex;
      except
         WIASource.SelectedItemIndex:= 0;
      end;
      WIASelectedItemIndex:= WIASource.SelectedItemIndex;

      PageSourceTypes.Enabled:= (WIASource.SelectedItem <> nil);
      if (PageSourceTypes.Enabled)
      then SelectCurrentItem(WIASelectedItemIndex)
      else MessageDlg('Error Selecting Source '+WIASource.Items[WIASelectedItemIndex]^.Name+' of '+WIASource.Name+
                #13#10'Try to Select another Source Item', mtError, [mbOk], 0);
    except
       on E: Exception do
       MessageDlg('Error Selecting Source ['+IntToStr(WIASelectedItemIndex)+'] of '+WIASource.Name+
                 #13#10+E.Message+#13#10'Try to Select another Source Item', mtError, [mbOk], 0);
    end;
    cbSourceItem.ItemIndex:= WIASelectedItemIndex;

    Result := (ShowModal=mrOk);

    if Result then
    begin
      ASelectedItemIndex:= WIASelectedItemIndex;

      StoreCurrentItemParams;

      //Do A Copy of WIAParamsArray to AParams Array and stretch it if needed
      if (Length(AParamsArray) < Length(WIAParamsArray)) then SetLength(AParamsArray, Length(WIAParamsArray));
      //Move(WIAParamsArray, AParamsArray, Length(AParamsArray));
      //AParams:= Copy(WIAParamsArray);
      for i:=0 to Length(AParamsArray)-1 do AParamsArray[i]:= WIAParamsArray[i];
    end;

  finally
    WIAParamsArray:= nil;
  end;
end;

end.

