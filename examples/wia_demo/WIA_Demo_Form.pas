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
  WIASource:TWIADevice;
  i,
  newIndex: Integer;

begin
    btAcquire.Enabled :=False;

    //Open the Select Device Dialog
    newIndex:= FWia.SelectDeviceDialog;
    if (newIndex > -1) then
    begin
      //Select the Device
      FWia.SelectedDeviceIndex:= newIndex;

      WIASource:= FWia.SelectedDevice;

      btAcquire.Enabled := (WIASource <> nil) and WiaSource.GetParamsCapabilities(WIACap);

      WIAParams:= WIACopyDefaultValues(WIACap);
      //WIAParams:= WIACopyCurrentValues(WIACap);
    end;
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
    if TWIASettingsSource.Execute(WIASource, NewItemIndex, initParams, WIAParams) then
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
        capRet:= WIASource.SetBrightness(Brightness);
        capRet:= WIASource.SetContrast(Contrast);

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
end;

end.

