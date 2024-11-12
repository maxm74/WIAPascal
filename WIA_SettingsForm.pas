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

//{$define UI_Tests}

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
    cbPaperType: TComboBox;
    cbDataType: TComboBox;
    cbResolution: TComboBox;
    cbBackFirst: TCheckBox;
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
    gbFeeder: TGroupBox;
    ImgListSource: TImageList;
    imgAlign: TImage;
    imgListAlign: TImageList;
    imgList: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    LabelPaperSize: TLabel;
    lvSourceItems: TListView;
    PageSourceTypes: TPageControl;
    Panel1: TPanel;
    panelCenter: TPanel;
    panelButtons: TPanel;
    btBrightness0: TSpeedButton;
    btBrightnessD: TSpeedButton;
    rbFrontBack: TRadioButton;
    rbFrontOnly: TRadioButton;
    rbBackOnly: TRadioButton;
    tbSource_Scanner: TTabSheet;
    trBrightness: TTrackBar;
    trHAlign: TTrackBar;
    trResolution: TTrackBar;
    trContrast: TTrackBar;
    trVAlign: TTrackBar;
    procedure bt0Click(Sender: TObject);
    procedure btDClick(Sender: TObject);
    procedure btPaperOrientationClick(Sender: TObject);
    procedure cbBackFirstClick(Sender: TObject);
    procedure cbPaperTypeChange(Sender: TObject);
    procedure cbUseNativeUIChange(Sender: TObject);
    procedure edBrightnessChange(Sender: TObject);
    procedure edContrastChange(Sender: TObject);
    procedure edResolutionChange(Sender: TObject);
    procedure lvSourceItemsSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure trBrightnessChange(Sender: TObject);
    procedure trContrastChange(Sender: TObject);
    procedure trHAlignChange(Sender: TObject);
    procedure trResolutionChange(Sender: TObject);
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

procedure TWIASettingsSource.cbBackFirstClick(Sender: TObject);
begin
  if rbFrontBack.Enabled
  then if cbBackFirst.Checked
       then rbFrontBack.Checked:= True;
end;

procedure TWIASettingsSource.cbPaperTypeChange(Sender: TObject);
var
   selPaperSize: TWIAPaperType;

begin
  selPaperSize:= TWIAPaperType(PtrUInt(cbPaperType.Items.Objects[cbPaperType.ItemIndex]));
  Case selPaperSize of
    wptMAX, wptAUTO: begin
       gbPaperAlign.Enabled:= False;
       btPaperOrientation.Enabled:= False;
       gbPaperSize.Visible:= True;
       gbPaperSize.Enabled:= False;
       edPaperW.Value:= edPaperW.MaxValue;
       edPaperH.Value:= edPaperW.MaxValue;
    end;
    wptCUSTOM: begin
       gbPaperAlign.Enabled:= True;
       btPaperOrientation.Enabled:= True;
       gbPaperSize.Visible:= True;
       edPaperW.Visible:= True;
       edPaperH.Visible:= True;
       gbPaperSize.Enabled:= True;
       edPaperW.Value:= THInchToSize(WIASettings_Unit_cm, WIAParams[WIASelectedItemIndex].PaperW);
       edPaperH.Value:= THInchToSize(WIASettings_Unit_cm, WIAParams[WIASelectedItemIndex].PaperH);
    end;
    else begin
       gbPaperAlign.Enabled:= True;
       btPaperOrientation.Enabled:= True;
       gbPaperSize.Visible:= False;
       edPaperW.Value:= THInchToSize(WIASettings_Unit_cm, WIAParams[WIASelectedItemIndex].PaperW);
       edPaperH.Value:= THInchToSize(WIASettings_Unit_cm, WIAParams[WIASelectedItemIndex].PaperH);
    end;
  end;
end;

procedure TWIASettingsSource.edContrastChange(Sender: TObject);
begin
  trContrast.Position:=edContrast.Value;
end;

procedure TWIASettingsSource.edResolutionChange(Sender: TObject);
begin
  trResolution.Position:= edResolution.Value;
end;

procedure TWIASettingsSource.lvSourceItemsSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  if Selected and (Item.Index <> WIASelectedItemIndex) then
  Case MessageDlg('Apply Changes to Item '+WIASource.Items[WIASelectedItemIndex]^.Name+' of '+WIASource.Name,
                 mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
    mrYes: begin
            StoreCurrentItemParams;
            SelectCurrentItem(lvSourceItems.ItemIndex);
    end;
    mrNo: SelectCurrentItem(lvSourceItems.ItemIndex);
  end;
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

procedure TWIASettingsSource.trResolutionChange(Sender: TObject);
begin
  edResolution.Value:=trResolution.Position;
end;

procedure TWIASettingsSource.SelectCurrentItem(AIndex: Integer);
var
   AParams: TWIAParams;
   ACap: TWIAParamsCapabilities;
   capRet: Boolean;
   paperI: TWIAPaperType;
   dataI: TWIADataType;
   i, cbSelected: Integer;
   curItem: PWIAItem;

begin
  AParams:= WIAParams[AIndex];
  ACap:= WIACaps[AIndex];
  curItem:= WIASource.Items[AIndex];

  if (curItem <> nil) then
  with ACap do
  begin
    cbUseNativeUI.Checked:= AParams.NativeUI;
    PageSourceTypes.Enabled:= not(cbUseNativeUI.Checked) and (curItem^.ItemCategory <> wicAUTO);

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

      if (paperI = AParams.PaperType) then cbSelected :=cbPaperType.Items.Count-1;
    end;
    cbPaperType.ItemIndex:=cbSelected;

    //Set Landscape/Portrait Button
    btPaperOrientation.Down:= (AParams.Rotation in [wrLandscape, wrRot270]);
    if btPaperOrientation.Down
    then begin btPaperOrientation.ImageIndex:= 1; btPaperOrientation.Hint:= 'Landscape'; end
    else begin btPaperOrientation.ImageIndex:= 0; btPaperOrientation.Hint:= 'Portrait'; end;

    //Paper Align
    trHAlign.Position:= Integer(AParams.HAlign);
    trVAlign.Position:= Integer(AParams.VAlign);
    trHAlignChange(nil);

    //Set Max,Current Values for Custom Paper Size
    edPaperW.MaxValue:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxWidth);
    edPaperH.MaxValue:= THInchToSize(WIASettings_Unit_cm, PaperSizeMaxHeight);

    if (curItem^.ItemCategory = wicFEEDER) then
    begin
      { #todo 5 -oMaxM : Must be tested in a Duplex Scanner }
      gbFeeder.Visible:= True;
      if (wdhAdvanced_Duplex in DocHandlingSet)
      then begin
             //WIA 2 Item structure
             rbFrontOnly.Enabled:= True;
             rbFrontBack.Enabled:= True;
             rbBackOnly.Enabled:= True;
             cbBackFirst.Enabled:= True;
           end
      else begin
             rbFrontOnly.Enabled:= (wdhFront_Only in DocHandlingSet);
             rbFrontBack.Enabled:= (wdhDuplex in DocHandlingSet);
             rbBackOnly.Enabled:= (wdhBack_Only in DocHandlingSet);
             cbBackFirst.Enabled:= (wdhBack_First in DocHandlingSet);
           end;
      {$ifdef UI_Tests}
        rbFrontOnly.Enabled:= True;
        rbFrontBack.Enabled:= True;
        rbBackOnly.Enabled:= True;
        cbBackFirst.Enabled:= True;
      {$endif}

      rbFrontOnly.Checked:= (wdhFront_Only in AParams.DocHandling);
      rbFrontBack.Checked:= (wdhDuplex in AParams.DocHandling) or (wdhAdvanced_Duplex in AParams.DocHandling);
      rbBackOnly.Enabled:= (wdhBack_Only in AParams.DocHandling);
      cbBackFirst.Enabled:= (wdhBack_First in AParams.DocHandling);
     end
     else gbFeeder.Visible:= False;

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
      Case dataI of
      wdtAUTO: cbDataType.Items.AddObject('Auto type', TObject(PtrUInt(wdtAUTO)));
      wdtDITHER, wdtCOLOR_DITHER: begin end;
      else begin
             cbDataType.Items.AddObject(WIADataTypeDescr[dataI], TObject(PtrUInt(dataI)));

             if (dataI = AParams.DataType) then cbSelected :=cbDataType.Items.Count-1;
           end;
      end;
    cbDataType.ItemIndex:=cbSelected;

    {$ifdef UI_Tests}
      ResolutionRange:= True;
    {$endif}

    if ResolutionRange
    then begin
           trResolution.Visible:= True; edResolution.Visible:= True;
           cbResolution.Visible:= False;

           //Set Resolution Range Limit
           {$ifdef UI_Tests}
             trResolution.Min:= ResolutionArray[0];
             trResolution.Max:= ResolutionArray[Length(ResolutionArray)-1];
             trResolution.LineSize:= 1;

             edResolution.MinValue:= ResolutionArray[0];
             edResolution.MaxValue:= ResolutionArray[Length(ResolutionArray)-1];
             edResolution.Increment:= 1;
           {$else}
             trResolution.Min:= ResolutionArray[WIA_RANGE_MIN];
             trResolution.Max:= ResolutionArray[WIA_RANGE_MAX];
             trResolution.LineSize:= ResolutionArray[WIA_RANGE_STEP];

             edResolution.MinValue:= ResolutionArray[WIA_RANGE_MIN];
             edResolution.MaxValue:= ResolutionArray[WIA_RANGE_MAX];
             edResolution.Increment:= ResolutionArray[WIA_RANGE_STEP];
           {$endif}

           trResolution.Position:= AParams.Resolution;
           edResolution.Value:= AParams.Resolution;
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

             if (ResolutionArray[i] = AParams.Resolution) then cbSelected :=cbResolution.Items.Count-1;
           end;
           cbResolution.ItemIndex:=cbSelected;
         end;

    //Brightness
    trBrightness.Min:= BrightnessMin;
    trBrightness.Max:= BrightnessMax;
    trBrightness.LineSize:= BrightnessStep;
    trBrightness.Position:= AParams.Brightness;
    edBrightness.MinValue:= BrightnessMin;
    edBrightness.MaxValue:= BrightnessMax;
    edBrightness.Increment:= BrightnessStep;
    edBrightness.Value:= trBrightness.Position;

    //Contrast
    trContrast.Min:= ContrastMin;
    trContrast.Max:= ContrastMax;
    trContrast.LineSize:= ContrastStep;
    trContrast.Position:= AParams.Contrast;
    edContrast.MinValue:= ContrastMin;
    edContrast.MaxValue:= ContrastMax;
    edContrast.Increment:= ContrastStep;
    edContrast.Value:= trContrast.Position;
  end;

  curParams:= AParams;
  curCap:= ACap;
  WIASelectedItemIndex:= AIndex;
  cbPaperTypeChange(nil);
end;

procedure TWIASettingsSource.StoreCurrentItemParams;
var
   curItem: PWIAItem;

begin
  curItem:= WIASource.Items[WIASelectedItemIndex];

  if (curItem <> nil) then
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

      if (curItem^.ItemCategory = wicFEEDER) then
      begin
        { #todo 5 -oMaxM : Must be tested in a Duplex Scanner }
        DocHandling:= [];
        if rbFrontOnly.Checked
        then DocHandling:= DocHandling+[wdhFront_Only]
        else
        if rbFrontBack.Checked
        then begin
               //if is A WIA2 structure add wdhAdvanced_Duplex so TWIADevice.Download understand how to work
               if (wdhAdvanced_Duplex in WIACaps[WIASelectedItemIndex].DocHandlingSet)
               then DocHandling:= DocHandling+[wdhAdvanced_Duplex]
               else DocHandling:= DocHandling+[wdhDuplex];

               if cbBackFirst.Checked then DocHandling:= DocHandling+[wdhBack_First];
             end
        else
        if rbBackOnly.Checked
        then DocHandling:= DocHandling+[wdhBack_Only];
      end;

      (*
      if (cbBitDepth.ItemIndex>-1)
      then BitDepth:=PtrUInt(cbBitDepth.Items.Objects[cbBitDepth.ItemIndex]);
      *)

      if (cbDataType.ItemIndex>-1)
      then DataType:=TWIADataType(PtrUInt(cbDataType.Items.Objects[cbDataType.ItemIndex]));

      if WIACaps[WIASelectedItemIndex].ResolutionRange
      then Resolution:= edResolution.Value
      else if (cbResolution.ItemIndex > -1)
           then Resolution:= curCap.ResolutionArray[PtrUInt(cbResolution.Items.Objects[cbResolution.ItemIndex])];

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
  curListItem: TListItem;
  curItem: PWIAItem;

begin
  Result:= False;
  if (AWIASource = nil) then exit;

  if (WIASettingsSource=nil)
  then WIASettingsSource :=TWIASettingsSource.Create(nil);

  if (WIASettingsSource <> nil) then
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

    //Get Capabilities and Fill ListView of Source Items
    lvSourceItems.Clear;
    for i:=0 to itemCount-1 do
    begin
      //Get Item[i] Default Values
      WIASource.SelectedItemIndex:= i;
      curItem:= WIASource.Items[i];

      if (curItem = nil)
      then raise Exception.Create('Cannot Get Source Item '+IntToStr(i));

      if not(WIASource.GetParamsCapabilities(WiaCaps[i]))
      then raise Exception.Create('Cannot Get Capabilities for Source Item '+IntToStr(i));

      Case initItemValues of
      initDefault: WIAParams[i]:= WIACopyDefaultValues(WiaCaps[i]);
      initParams : if (i >= lenAParams) //if is a new Item then Assign Default Values
                   then WIAParams[i]:= WIACopyDefaultValues(WiaCaps[i]);
      initCurrent: WIAParams[i]:= WIACopyCurrentValues(WiaCaps[i]);
      end;

      curListItem:= lvSourceItems.Items.Add;
      curListItem.Caption :=WIASource.Items[i]^.Name;
      curListItem.Data:= Pointer(i);

      Case curItem^.ItemCategory of
      wicFLATBED: curListItem.ImageIndex:= 1;
      wicFEEDER,
      wicFEEDER_FRONT,
      wicFEEDER_BACK:  curListItem.ImageIndex:= 2;
      wicFILM:  curListItem.ImageIndex:= 3;
      wicROOT,
      wicFOLDER:  curListItem.ImageIndex:= 5;
      wicBARCODE_READER: curListItem.ImageIndex:= 6;
      else  curListItem.ImageIndex:= 0;
      end;
    end;

    try
      //Select the Initial Item to ASelectedItemIndex
      if (ASelectedItemIndex < 0)
      then WIASelectedItemIndex:= 0
      else WIASelectedItemIndex:= ASelectedItemIndex;
      try
         WIASource.SelectedItemIndex:= WIASelectedItemIndex;
      except
         WIASource.SelectedItemIndex:= 0;
      end;
      WIASelectedItemIndex:= WIASource.SelectedItemIndex;

      PageSourceTypes.Enabled:= (WIASource.SelectedItemIntf <> nil);
      if (PageSourceTypes.Enabled)
      then SelectCurrentItem(WIASelectedItemIndex)
      else MessageDlg('Error Selecting Source '+WIASource.Items[WIASelectedItemIndex]^.Name+' of '+WIASource.Name+
                #13#10'Try to Select another Source Item', mtError, [mbOk], 0);
    except
       on E: Exception do
       MessageDlg('Error Selecting Source ['+IntToStr(WIASelectedItemIndex)+'] of '+WIASource.Name+
                 #13#10+E.Message+#13#10'Try to Select another Source Item', mtError, [mbOk], 0);
    end;
    //cbSourceItem.ItemIndex:= WIASelectedItemIndex;
    lvSourceItems.ItemIndex:= WIASelectedItemIndex;
    SelectCurrentItem(WIASelectedItemIndex);

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
    WIASettingsSource.Free; WIASettingsSource:= nil;
  end;
end;

end.

