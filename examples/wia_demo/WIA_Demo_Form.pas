unit WIA_Demo_Form;

{$ifdef fpc}
  {$mode delphi}
{$endif}

{$H+}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, ComCtrls, Spin,
  WIA, WIA_LH, WiaDef, WIA_PaperSizes, WIA_SettingsForm;

type
  { TFormWIADemo }

  TFormWIADemo = class(TForm)
    btAcquire: TButton;
    btSelect: TButton;
    btNative: TButton;
    cbTest: TCheckBox;
    cbEnumLocal: TCheckBox;
    edTests: TEdit;
    ImageHolder: TImage;
    Label1: TLabel;
    lbSelected: TLabel;
    lbProgress: TLabel;
    Panel1: TPanel;
    progressBar: TProgressBar;
    edPages: TSpinEdit;
    procedure btNativeClick(Sender: TObject);
    procedure btSelectClick(Sender: TObject);
    procedure btAcquireClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    lres: HResult;
    FWia: TWIAManager;
    WIACap: TWIAParamsCapabilities;
    WIAParams: TArrayWIAParams;
    WIAformat,
    WIAformatDef: TWiaImageFormat;
    WIAFormatSet: TWiaImageFormatSet;
    WIAformatStr: String;

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
  WIAParams:= nil;
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
    lbProgress.Caption:= 'Downloading '+AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count)+' file';
  end;
  WIA_TRANSFER_MSG_END_OF_STREAM: begin
    lbProgress.Caption:= 'Downloaded '+AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count)+' file';
  end;
  WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
  end;
  WIA_TRANSFER_MSG_DEVICE_STATUS: begin
  end;
  WIA_TRANSFER_MSG_NEW_PAGE: begin
    //lbProgress.Caption:= AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count+1);
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
  propType: TVarType;
  capRet: Boolean;

begin
    btAcquire.Enabled :=False; btNative.Enabled :=False;
    lbSelected.Caption:= '';

    FWia.EnumAll:= not(cbEnumLocal.Checked);

    //Open the Select Device Dialog
    newIndex:= FWia.SelectDeviceDialog;
    if (newIndex > -1) then
    begin
      //Select the Device
      FWia.SelectedDeviceIndex:= newIndex;

      WIASource:= FWia.SelectedDevice;
      if (WIASource <> nil) then
      begin
        btAcquire.Enabled := True; ///WiaSource.GetParamsCapabilities(WIACap);
        lbSelected.Caption:= 'Selected : '+WIASource.Name;

        //SetLength(WIAParams, 1);
        //WIAParams[0]:= WIACopyDefaultValues(WIACap);
        //WIAParams[0]:= WIACopyCurrentValues(WIACap);
        WIAParams:= nil;

        if cbTest.Checked then
        begin
          capRet:= WIASource.GetImageFormat(WIAformat, WIAformatDef, WIAFormatSet);
          if capRet then MessageDlg('ImageFormat: '+
                                    #13#10+'Current: '+WiaImageFormatDescr[WIAformat]+
                                    #13#10+'Default: '+WiaImageFormatDescr[WIAformatDef],
                                    mtInformation, [mbOk], 0)
                    else MessageDlg('ImageFormat Error', mtInformation, [mbOk], 0);

        end;
      end
      else MessageDlg('Error Connecting Device', mtError, [mbOk], 0);
    end;
    btNative.Enabled :=btAcquire.Enabled;
end;

procedure TFormWIADemo.btNativeClick(Sender: TObject);
var
   aPath: String;
   WIASource: TWIADevice;
   c: Integer;
   DownloadedFiles: TStringArray;

begin
  WIASource:= FWia.SelectedDevice;
  if (Sender <> nil) then WIASource.SelectedItemIndex:= 0;
  aPath:= ExtractFilePath(ParamStr(0));
  c:= WiaSource.DownloadNativeUI(Self.Handle, False, aPath, 'test_wia', DownloadedFiles);
  if (c>0)
  then begin
           MessageDlg('Downloaded '+IntToStr(c)+' Files', mtInformation, [mbOk], 0);

           ImageHolder.Picture.Bitmap.LoadFromFile(DownloadedFiles[0]);
         end
    else MessageDlg('NO Files Downloaded ', mtError, [mbOk], 0);
  DownloadedFiles:= nil;
end;

procedure TFormWIADemo.btAcquireClick(Sender: TObject);
var
   aPath, aExt: String;
   WIASource: TWIADevice;
   SelectedItemIndex,
   NewItemIndex: Integer;
   c, p: Integer;
   capRet: Boolean;
   propType: TVarType;
   x, y, w, h:Integer;
   aFormat: TWIAImageFormat;

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
        SelectedItemIndex:= NewItemIndex;

      end;
      WIASource.SelectedItemIndex:= SelectedItemIndex;

      aPath:= ExtractFilePath(ParamStr(0));

      if WIAParams[SelectedItemIndex].NativeUI
      then begin
             btNativeClick(nil);
           end
      else begin
             if cbTest.Checked then
             begin
               capRet:= WIASource.GetProperty(WIA_IPS_XPOS, propType, x);
               capRet:= WIASource.GetProperty(WIA_IPS_YPOS, propType, y);
               capRet:= WIASource.GetProperty(WIA_IPS_XEXTENT, propType, w);
               capRet:= WIASource.GetProperty(WIA_IPS_YEXTENT, propType, h);
               if capRet then MessageDlg('Paper Position Before SetParams: '+
                                  #13#10+'x: '+IntToStr(x)+
                                  #13#10+'y: '+IntToStr(y)+
                                  #13#10+'w: '+IntToStr(w)+
                                  #13#10+'h: '+IntToStr(h),
                                  mtInformation, [mbOk], 0)
             end;

//             WIAParams[SelectedItemIndex].DocHandling:= [wdhDuplex];
//             WIAParams[SelectedItemIndex].Resolution:= 200;
//             WIAParams[SelectedItemIndex].DataType:= wdtGRAYSCALE;

{ #todo 10 -oMaxM : When switching from pages=1 to pages=0 my Brother scanner returns an image
             shrunk by a factor N within the correct sized page.
             N= (Resolution when pages=0) / (Resolution when pages=1). I'm stumped. }
             WIASource.SetPages(edPages.Value);
             //capRet:= WIASource.GetPages(c, x, y, w, h);

             WIASource.SetParams(WIAParams[SelectedItemIndex]);

             if cbTest.Checked then
             begin
               capRet:= WIASource.GetProperty(WIA_IPS_XPOS, propType, x);
               capRet:= WIASource.GetProperty(WIA_IPS_YPOS, propType, y);
               capRet:= WIASource.GetProperty(WIA_IPS_XEXTENT, propType, w);
               capRet:= WIASource.GetProperty(WIA_IPS_YEXTENT, propType, h);
               if capRet then MessageDlg('Paper Position After SetParams: '+
                                    #13#10+'x: '+IntToStr(x)+
                                    #13#10+'y: '+IntToStr(y)+
                                    #13#10+'w: '+IntToStr(w)+
                                    #13#10+'h: '+IntToStr(h),
                                    mtInformation, [mbOk], 0);
               if (edTests.Text<>'') then
               begin
                 capRet:= WIASource.SetBitDepth(StrToInt(edTests.Text));
                 if not(capRet) then raise Exception.Create('SetBitDepth');
               end;
             end;

             { #todo 5 -oMaxM : Move To Download? }
             if WIAParams[SelectedItemIndex].DataType in [wdtRAW_RGB..wdtRAW_CMYK]
             then capRet:= WIASource.SetImageFormat(wifRAW)
             else capRet:= WIASource.SetImageFormat(wifBMP);
             if not(capRet) then raise Exception.Create('SetImageFormat');

             if WIAParams[SelectedItemIndex].DataType in [wdtRAW_RGB..wdtRAW_CMYK]
             then begin
                    aFormat:= wifRAW;
                    aExt:= '.raw'
                  end
             else begin
                    aFormat:= wifBMP;
                    aExt:= '.bmp';
                  end;

             c:= WIASource.Download(aPath, 'test_wia', aExt,
                                    aFormat (*, WIAParams[SelectedItemIndex].DocHandling*));

             if cbTest.Checked then
             begin
                 capRet:= WIASource.GetProperty(WIA_IPS_XPOS, propType, x);
                 capRet:= WIASource.GetProperty(WIA_IPS_YPOS, propType, y);
                 capRet:= WIASource.GetProperty(WIA_IPS_XEXTENT, propType, w);
                 capRet:= WIASource.GetProperty(WIA_IPS_YEXTENT, propType, h);
                 if capRet then MessageDlg('Paper Position After Download: '+
                                      #13#10+'x: '+IntToStr(x)+
                                      #13#10+'y: '+IntToStr(y)+
                                      #13#10+'w: '+IntToStr(w)+
                                      #13#10+'h: '+IntToStr(h),
                                      mtInformation, [mbOk], 0);

             end;

             if (c>0)
             then begin
                    MessageDlg('Downloaded '+IntToStr(c)+' Files', mtInformation, [mbOk], 0);
                    ImageHolder.Picture.Bitmap.LoadFromFile(aPath+'test_wia'+aExt);
                  end
             else MessageDlg('NO Files Downloaded ', mtError, [mbOk], 0);
          end;
     end;

  finally
    WIASettingsSource.Free; WIASettingsSource:= Nil;
  end;
end;

end.

