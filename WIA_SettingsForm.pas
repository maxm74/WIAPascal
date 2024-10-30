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
  DelphiCompatibility, Classes, SysUtils, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, Buttons, ComCtrls, StdCtrls, Spin,
  WIA, WIA_PaperSizes,
  ImgList {$ifndef fpc}, ImageList, NumberBox{$endif};

const
  //False to display then measurement in Inchs
  WIASettings_Unit_cm: Boolean = True; //fucks inch

type
  TInitialItemValues = (initDefault, initParams, initCurrent);

  { TWIASettingsSource }
  TInitDefaultValuesEvent = procedure (var ACap: TWIAParamsCapabilities) of object;

  TWIASettingsSource = class(TForm)
    btContrast0: TSpeedButton;
    btContrastD: TSpeedButton;
    btResolutionD: TSpeedButton;
    btCancel: TBitBtn;
    btPaperOrientation: TSpeedButton;
    btRefreshUndo: TBitBtn;
    btOk: TBitBtn;
    btRefreshCurrent: TBitBtn;
    btRefreshDefault: TBitBtn;
    cbBitDepth: TComboBox;
    cbSourceItem: TComboBox;
    cbPaperType: TComboBox;
    cbDataType: TComboBox;
    cbResolution: TComboBox;
    cbUseNativeUI: TCheckBox;
    edBrightness: TSpinEdit;
    {$ifdef fpc}
    edPaperH: TFloatSpinEdit;
    edPaperW: TFloatSpinEdit;
    {$else}
    edPaperH: TNumberBox;
    edPaperW: TNumberBox;
    {$endif}
    edResolution: TSpinEdit;
    edContrast: TSpinEdit;
    gbPaperAlign: TGroupBox;
    gbPaperSize: TGroupBox;
    imgAlign: TImage;
    imgListAlign: TImageList;
    imgList: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    PageSourceTypes: TPageControl;
    Panel1: TPanel;
    panelButtons: TPanel;
    btBrightness0: TSpeedButton;
    btBrightnessD: TSpeedButton;
    tbSource_Scanner: TTabSheet;
    trBrightness: TTrackBar;
    trHAlign: TTrackBar;
    trResolution: TTrackBar;
    trContrast: TTrackBar;
    trVAlign: TTrackBar;
    procedure bt0Click(Sender: TObject);
    procedure btDClick(Sender: TObject);
    procedure btPaperOrientationClick(Sender: TObject);
    procedure cbPaperTypeChange(Sender: TObject);
    procedure cbSourceItemChange(Sender: TObject);
    procedure cbUseNativeUIChange(Sender: TObject);
    procedure edBrightnessChange(Sender: TObject);
    procedure edContrastChange(Sender: TObject);
    procedure trBrightnessChange(Sender: TObject);
    procedure trContrastChange(Sender: TObject);
    procedure trHAlignChange(Sender: TObject);
  private
    WIASource: TWIADevice;
    WIAPaperMaxWidth,
    WIAPaperMaxHeight,
    WIASelectedItemIndex: Integer;
    WIACaps: TArrayWIAParamsCapabilities;
    WIAParams: TArrayWIAParams;
    curCap: TWIAParamsCapabilities;
    curParams: TWIAParams;
    initItemValues: TInitialItemValues;

    OnInitDefaultValues: TInitDefaultValuesEvent;

    procedure SelectCurrentItem(AIndex: Integer);
    procedure StoreCurrentItemParams;

  public
     class function Execute(AWIASource: TWIADevice;
                            var ASelectedItemIndex: Integer;
                            { #todo -oMaxM : Possibly Filters for which Items Kinds to Show? How manage AParams without Indexes? }
                            AInitItemValues: TInitialItemValues;
                            var AParams: TArrayWIAParams;
                            AOnInitDefaultValues: TInitDefaultValuesEvent=nil): Boolean;
  end;

var
  WIASettingsSource: TWIASettingsSource=nil;

implementation

{$ifdef fpc}
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

procedure TWIASettingsSource.cbUseNativeUIChange(Sender: TObject);
begin
  PageSourceTypes.Enabled:= not(cbUseNativeUI.Checked);
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
  with curCap do
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

procedure TWIASettingsSource.btPaperOrientationClick(Sender: TObject);
begin
  if btPaperOrientation.Down
  then begin btPaperOrientation.ImageIndex:= 1; btPaperOrientation.Hint:= 'Landscape'; end
  else begin btPaperOrientation.ImageIndex:= 0; btPaperOrientation.Hint:= 'Portrait'; end;
end;

procedure TWIASettingsSource.cbPaperTypeChange(Sender: TObject);
var
   selPaperSize: TWIAPaperType;

begin
  selPaperSize:= TWIAPaperType(PtrUInt(cbPaperType.Items.Objects[cbPaperType.ItemIndex]));
  Case selPaperSize of
    wptMAX, wptAUTO: begin
       gbPaperAlign.Enabled:= False;
       gbPaperSize.Visible:= True;
       gbPaperSize.Enabled:= False;
       btPaperOrientation.Enabled:= False;
    end;
    wptCUSTOM: begin
       gbPaperAlign.Enabled:= True;
       gbPaperSize.Visible:= True;
       gbPaperSize.Enabled:= True;
       btPaperOrientation.Enabled:= True;
    end;
    else begin
       gbPaperAlign.Enabled:= True;
       gbPaperSize.Visible:= False;
       btPaperOrientation.Enabled:= True;
    end;
  end;
end;

procedure TWIASettingsSource.edContrastChange(Sender: TObject);
begin
  trContrast.Position:=edContrast.Value;
end;

procedure TWIASettingsSource.trContrastChange(Sender: TObject);
begin
  edContrast.Value:=trContrast.Position;
end;

procedure TWIASettingsSource.trHAlignChange(Sender: TObject);
begin
  {$ifdef fpc}
  imgAlign.ImageIndex:= (trVAlign.Position*3)+trHAlign.Position;
  {$else}
  imgListAlign.GetIcon((trVAlign.Position*3)+trHAlign.Position, imgAlign.Picture.Icon);
  {$endif}
end;

procedure TWIASettingsSource.SelectCurrentItem(AIndex: Integer);
var
   AParams: TWIAParams;
   ACap: TWIAParamsCapabilities;
   capRet: Boolean;
   paperI: TWIAPaperType;
   dataI: TWIADataType;
   i, cbSelected: Integer;

begin
  AParams:= WIAParams[AIndex];
  ACap:= WIACaps[AIndex];
  //WiaSource.SelectedItemIndex:= AIndex; No Need we have ACap already Filled

  if (initItemValues = initParams)
  then cbUseNativeUI.Checked:= AParams.NativeUI
  else cbUseNativeUI.Checked:= False;

  PageSourceTypes.Enabled:= not(cbUseNativeUI.Checked);

  with ACap do
  begin
  //Fill List of Papers
  cbPaperType.Clear;
  cbSelected :=0;
  cbPaperType.Items.AddObject('Full size', TObject(PtrUInt(wptMAX)));
  for paperI in PaperTypeSet do
  begin
    Case paperI of
    wptMAX:begin end;
    wptCUSTOM: cbPaperType.Items.AddObject('Custom size', TObject(PtrUInt(wptCUSTOM)));
    wptAUTO: cbPaperType.Items.AddObject('Auto size', TObject(PtrUInt(wptAUTO)));
    else begin
           if WIASettings_Unit_cm
           then cbPaperType.Items.AddObject(PaperSizesWIA[paperI].name+' ('+
                                            THInchToCmStr(PaperSizesWIA[paperI].w)+' x '+
                                            THInchToCmStr(PaperSizesWIA[paperI].h)+' cm)',
                                            TObject(PtrUInt(paperI)))
           else cbPaperType.Items.AddObject(PaperSizesWIA[paperI].name+' ('+
                                            THInchToInchStr(PaperSizesWIA[paperI].w)+' x '+
                                            THInchToInchStr(PaperSizesWIA[paperI].h)+' in)',
                                            TObject(PtrUInt(paperI)))
         end;
    end;

    case initItemValues of
    initDefault: begin if (paperI = PaperTypeDefault) then cbSelected :=cbPaperType.Items.Count-1; end;
    initParams:  begin if (paperI = AParams.PaperType) then cbSelected :=cbPaperType.Items.Count-1; end;
    initCurrent: begin if (paperI = PaperTypeCurrent) then cbSelected :=cbPaperType.Items.Count-1; end;
    end;
  end;
  cbPaperType.ItemIndex:=cbSelected;

  //Set Landscape/Portrait Button
  case initItemValues of
  initDefault: btPaperOrientation.Down:= (RotationDefault in [wrLandscape, wrRot270]); //270 is basically landscape
  initParams:  btPaperOrientation.Down:= (AParams.Rotation in [wrLandscape, wrRot270]);
  initCurrent: btPaperOrientation.Down:= (RotationCurrent in [wrLandscape, wrRot270]);
  end;
  if btPaperOrientation.Down
  then begin btPaperOrientation.ImageIndex:= 1; btPaperOrientation.Hint:= 'Landscape'; end
  else begin btPaperOrientation.ImageIndex:= 0; btPaperOrientation.Hint:= 'Portrait'; end;

  //Paper Align
  if (initItemValues = initParams)
  then begin
         trHAlign.Position:= Integer(AParams.HAlign);
         trVAlign.Position:= Integer(AParams.VAlign);
       end
  else begin
         trHAlign.Position:= 0;
         trVAlign.Position:= 0;
       end;
  trHAlignChange(nil);

  //Set Max,Current Values for Custom Paper Size
  edPaperW.Value:= 0;
  edPaperH.Value:= 0;
  edPaperW.MaxValue:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxWidth);
  edPaperH.MaxValue:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxHeight);

  if (TWIAPaperType(PtrUInt(cbPaperType.Items.Objects[cbPaperType.ItemIndex])) = wptCUSTOM) and
     (initItemValues = initParams)
  then begin
         edPaperW.Value:= THInchToSize(WIASettings_Unit_cm, AParams.PaperW);
         edPaperH.Value:= THInchToSize(WIASettings_Unit_cm, AParams.PaperH);
       end
  else begin
         edPaperW.Value:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxWidth);
         edPaperH.Value:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxHeight);
       end;

  cbPaperTypeChange(nil);

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
           cbDataType.Items.AddObject(WIADataTypeDescr[dataI], TObject(PtrUInt(dataI)));

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

  curParams:= AParams;
  curCap:= ACap;
  WIASelectedItemIndex:= AIndex;
end;

procedure TWIASettingsSource.StoreCurrentItemParams;
begin
  with curParams do
  begin
    NativeUI:= cbUseNativeUI.Checked;
    if not(NativeUI) then
    begin
      if (cbPaperType.ItemIndex>-1)
      then PaperType:= TWIAPaperType(PtrUInt(cbPaperType.Items.Objects[cbPaperType.ItemIndex]));

      if (PaperType = wptCUSTOM) then
      begin
        PaperW:= SizeToTHInch(WIASettings_Unit_cm, edPaperW.Value);
        PaperH:= SizeToTHInch(WIASettings_Unit_cm, edPaperH.Value);
      end;

      if btPaperOrientation.Down
      then Rotation:= wrLandscape
      else Rotation:= wrPortrait;

      HAlign:= TWIAAlignHorizontal(trHAlign.Position);
      VAlign:= TWIAAlignVertical(trVAlign.Position);

      (*
      if (cbBitDepth.ItemIndex>-1)
      then BitDepth:=PtrUInt(cbBitDepth.Items.Objects[cbBitDepth.ItemIndex]);
      *)

      if (cbDataType.ItemIndex>-1)
      then DataType:=TWIADataType(PtrUInt(cbDataType.Items.Objects[cbDataType.ItemIndex]));

      if (cbResolution.ItemIndex>-1)
      then Resolution:=curCap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

      Contrast:=edContrast.Value;
      Brightness:=edBrightness.Value;
   end;
  end;

  WIAParams[WIASelectedItemIndex]:= curParams;
end;

class function TWIASettingsSource.Execute(AWIASource: TWIADevice; var ASelectedItemIndex: Integer;
                                          AInitItemValues: TInitialItemValues; var AParams: TArrayWIAParams;
                                          AOnInitDefaultValues: TInitDefaultValuesEvent): Boolean;
var
  i,
  itemCount,
  lenAParams: Integer;

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
    OnInitDefaultValues:= AOnInitDefaultValues;

    //Do A Copy of Params Array so if the user cancels the Dialog we don't modify the starting array
    WIAParams:= Copy(AParams);

    //If WiaSource Item Count is greater then our Array enlarge it
    lenAParams:= Length(AParams);
    if (lenAParams < itemCount)
    then SetLength(WIAParams, itemCount);

    SetLength(WiaCaps, itemCount);

    //Get Capabilities and Fill ComboBox of Source Items
    cbSourceItem.Clear;
    for i:=0 to itemCount-1 do
    begin
      //Get Item[i] Default Values
      WIASource.SelectedItemIndex:= i;
      if not(WIASource.GetParamsCapabilities(WiaCaps[i]))
      then raise Exception.Create('Cannot Get Capabilities for Source Item');

      //if is a new Item then Assign Default or Current Values
      if (i >= lenAParams)
      then if (initItemValues = initCurrent)
           then WIAParams[i]:= WIACopyCurrentValues(WiaCaps[i])
           else WIAParams[i]:= WIACopyDefaultValues(WiaCaps[i]);

      cbSourceItem.Items.AddObject(WIASource.Items[i]^.Name, TObject(PtrUInt(i)));
    end;

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

      //Do A Copy of WIAParams to AParams Array and stretch it if needed
      if (Length(AParams) < Length(WIAParams)) then SetLength(AParams, Length(WIAParams));
      //Move(WIAParams, AParams, Length(AParams));
      //AParams:= Copy(WIAParams);
      for i:=0 to Length(AParams)-1 do AParams[i]:= WIAParams[i];
    end;

  finally
    WIAParams:= nil;
    WIACaps:= nil;
  end;
end;

end.

