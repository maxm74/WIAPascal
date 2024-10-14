unit WIA_Demo_Form;

{$ifdef fpc}
  {$mode delphi}
{$endif}

{$H+}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, ComCtrls,
  WIA, WIA_LH, WiaDef, WIA_SettingsForm;

type
  { TFormWIADemo }

  TFormWIADemo = class(TForm)
    btAcquire: TButton;
    btSelect: TButton;
    ImageHolder: TImage;
    Panel1: TPanel;
    progressBar: TProgressBar;
    procedure btSelectClick(Sender: TObject);
    procedure btAcquireClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    (*
    rTwain:TDelphiTwain;
    capRet: TCapabilityRet;

    function getTwain: TDelphiTwain;

    property Twain: TDelphiTwain read getTwain;

    procedure TwainAcquireNative(Sender: TObject; const Index: Integer;
                                nativeHandle: TW_UINT32; var Cancel: Boolean);

    procedure TwainAcquire(Sender: TObject; const Index: Integer;
                           Image: TBitmap; var Cancel: Boolean);
    *)

    lres: HResult;
    FWia: TWIAManager;
    WIACap: TWIAParamsCapabilities;
    WIAParams: TWIAParams;

    function DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                                 lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean;
  end;

var
  FormWIADemo: TFormWIADemo;

implementation


{$ifdef fpc}
  {$R *.lfm}
{$else}
  {$R *.dfm}
{$endif}


{ TFormWIADemo }

procedure TFormWIADemo.FormCreate(Sender: TObject);
begin
    FWia :=TWIAManager.Create;
    FWia.OnAfterDeviceTransfer:= DeviceTransferEvent;
end;

procedure TFormWIADemo.FormDestroy(Sender: TObject);
begin
  if (FWIA<>nil) then FWia.Free;
end;

function TFormWIADemo.DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                                          lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean;
begin
  Result:= True;

  if (pWiaTransferParams <> nil) then
  Case pWiaTransferParams^.lMessage of
  WIA_TRANSFER_MSG_STATUS: begin
    progressBar.Position:= pWiaTransferParams^.lPercentComplete;
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_STATUS : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
  end;
  WIA_TRANSFER_MSG_END_OF_STREAM: begin
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_END_OF_STREAM : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
  end;
  WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_END_OF_TRANSFER : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
//    Memo2.Lines.Add('   Downloaded='+BoolToStr(AWiaDevice.Downloaded, True));
  end;
  WIA_TRANSFER_MSG_DEVICE_STATUS: begin
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_DEVICE_STATUS : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
  end;
  WIA_TRANSFER_MSG_NEW_PAGE: begin
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_NEW_PAGE : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
  end
  else begin
//    Memo2.Lines.Add('WIA_TRANSFER_MSG_'+IntToHex(pWiaTransferParams^.lMessage)+' : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
  end;
end;
end;

procedure TFormWIADemo.btSelectClick(Sender: TObject);
var
(*  TwainSource:TTwainSource;
  bitCurrent: Integer;
  paperCurrent: TTwainPaperSize;
  pixelCurrent:TTwainPixelType;
  resolutionCurrent:Single;
*)
  WIASource:TWIADevice;
  i,
  newIndex: Integer;

begin
    btAcquire.Enabled :=False;

    newIndex:= FWIA.SelectDeviceDialog;
    if (newIndex > -1) then
    begin
      FWIA.SelectedDeviceIndex:= newIndex;

      btAcquire.Enabled := (FWia.SelectedDevice <> nil);
    end;

    (*
    //Load source manager
    Twain.SourceManagerLoaded :=True;

    //Allow user to select source
    Twain.SelectSource;
    if Assigned(Twain.SelectedSource) then
    begin
      TwainSource :=Twain.SelectedSource;

      TwainSource.Loaded:=True;

      TwainCap.PaperFeedingSet:=TwainSource.GetPaperFeeding;
      capRet :=TwainSource.GetPaperSizeSet(paperCurrent, TwainCap.PaperSizeDefault, TwainCap.PaperSizeSet);
      capRet :=TwainSource.GetIBitDepth(bitCurrent, TwainCap.BitDepthDefault, TwainCap.BitDepthArray);
      TwainCap.BitDepthArraySize :=Length(TwainCap.BitDepthArray);
      capRet :=TwainSource.GetIPixelType(pixelCurrent, TwainCap.PixelTypeDefault, TwainCap.PixelType);
      capRet :=TwainSource.GetIXResolution(resolutionCurrent, TwainCap.ResolutionDefault, TwainCap.ResolutionArray);
      TwainCap.ResolutionArraySize :=Length(TwainCap.ResolutionArray);

      btAcquire.Enabled :=True;
    end;
    *)
end;

procedure TFormWIADemo.btAcquireClick(Sender: TObject);
var
   aPath: String;
   WIASource: TWIADevice;
   SelectedItemIndex,
   NewItemIndex: Integer;
   c, p: Integer;
   capRet: Boolean;
   t: TVarType;

begin
  WIASource:= FWia.SelectedDevice;
//  WIASource:= FWia.Devices[0];

  if (WIASource <> nil) then
  try
   //WIASource.SelectItem('Feeder');
    SelectedItemIndex:= WIASource.SelectedItemIndex;
    NewItemIndex:= SelectedItemIndex;

    //Select Scanner Setting to use
    if TWIASettingsSource.Execute(WIASource, NewItemIndex, initCurrent, WIAParams) then
    begin
      if (NewItemIndex <> SelectedItemIndex) then
      begin
        //Sub Item Changed do Something
      end;

      aPath:= ExtractFilePath(ParamStr(0));

      with WIAParams do
      begin
        capRet:= WIASource.SetResolution(Resolution, Resolution);
        capRet:= WIASource.SetPaperSize(PaperSize);

        (*
        { #todo 10 -oMaxM : The PaperSizes does not works, use the XEntent = pagesize * xres / 1000 }
        p:= 200;
        capRet:= WIASource.SetProperty(WIA_IPS_XRES, VT_I4, p);
        capRet:= WIASource.SetProperty(WIA_IPS_YRES, VT_I4, p);

        capRet:= WIASource.GetProperty(WIA_IPS_MAX_HORIZONTAL_SIZE, t, p);
        capRet:= WIASource.GetProperty(WIA_IPS_MAX_VERTICAL_SIZE, t, p);

        p:= Trunc(8267 * 200 / 1000);
        capRet:= WIASource.SetProperty(WIA_IPS_XEXTENT, VT_I4, p);
        p:=0;
        capRet:= WIASource.SetProperty(WIA_IPS_XPOS, VT_I4, p);
        p:= Trunc(11692 * 200 / 1000);
        capRet:= WIASource.SetProperty(WIA_IPS_YEXTENT, VT_I4, p);
        *)

(*        capRet :=TwainSource.SetIPixelType(PixelType);

        capRet :=TwainSource.SetIXResolution(Resolution);
        capRet :=TwainSource.SetIYResolution(Resolution);
        capRet :=TwainSource.SetContrast(Contrast);
        capRet :=TwainSource.SetBrightness(Brightness);
*)
      end;

      c:= WIASource.Download(aPath, 'test_wia.bmp');

      if (c>0)
      then begin
             MessageDlg('Downloaded '+IntToStr(c)+' Files', mtInformation, [mbOk], 0);
             ImageHolder.Picture.Bitmap.LoadFromFile(aPath+'test_wia.bmp');
           end
      else MessageDlg('NO Files Downloaded ', mtError, [mbOk], 0);
     end;

  finally
    WIASettingsSource.Free; WIASettingsSource:= Nil;
  end;

  (*
    //Show Select Form and User Select a device
    Twain.SelectedSource;

    if Assigned(Twain.SelectedSource) then
    begin
      aPath:=ExtractFilePath(ParamStr(0))+'test_0.bmp';
      if FileExists(aPath)
      then DeleteFile(aPath);

      TwainSource :=Twain.SelectedSource;
      TwainSource.Loaded := True;

      try
        //Select Scanner Setting to use
        if TTwainSettingsSource.Execute(True, TwainCap, TwainParams) then
        begin
          with TwainParams do
          begin
            capRet :=TwainSource.SetPaperFeeding(PaperFeed);
            capRet :=TwainSource.SetDuplexEnabled(False);
            capRet :=TwainSource.SetPaperSize(PaperSize);
            capRet :=TwainSource.SetIPixelType(PixelType);

            capRet :=TwainSource.SetIXResolution(Resolution);
            capRet :=TwainSource.SetIYResolution(Resolution);
            capRet :=TwainSource.SetContrast(Contrast);
            capRet :=TwainSource.SetBrightness(Brightness);
          end;
          capRet :=TwainSource.SetIndicators(True);

          //Load source, select transference method and enable
          TwainSource.TransferMode:=ttmNative;

          if cbNativeCapture.Checked
          then begin
                 Twain.OnTwainAcquireNative:=TwainAcquireNative;
                 Twain.OnTwainAcquire:=nil;
               end
          else begin
                 Twain.OnTwainAcquireNative:=nil;
                 Twain.OnTwainAcquire:=TwainAcquire;
               end;

          TwainSource.EnableSource(cbShowUI.Checked, cbModalCapture.Checked, Application.ActiveFormHandle);
        end;
      finally
        TwainSettingsSource.Free; TwainSettingsSource:= Nil;
      end;
    end;
    *)
end;

(*
function TFormWIADemo.getTwain: TDelphiTwain;
begin
  //Create Twain
  if (rTwain = nil) then
  begin
    rTwain := TDelphiTwain.Create;
    //Load Twain Library dynamically
    if not(rTwain.LoadLibrary) then ShowMessage('Twain is not installed.');
  end;

  Result :=rTwain;
end;

procedure TFormWIADemo.TwainAcquireNative(Sender: TObject; const Index: Integer;
  nativeHandle: TW_UINT32; var Cancel: Boolean);
begin
  try
     WriteBitmapToFile('test_0.bmp', nativeHandle);
     ImageHolder.Picture.Bitmap.LoadFromFile('test_0.bmp');
     Cancel := True;//Only want one image
  except
  end;
end;

procedure TFormWIADemo.TwainAcquire(Sender: TObject; const Index: Integer;
  Image: TBitmap; var Cancel: Boolean);
begin
  try
     ImageHolder.Picture.Bitmap.Assign(Image);
     ImageHolder.Picture.Bitmap.SaveToFile('test_0.bmp');
     Cancel := True;//Only want one image
  except
  end;
end;
*)

end.

