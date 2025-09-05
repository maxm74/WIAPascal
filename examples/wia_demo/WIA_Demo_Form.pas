(******************************************************************************
*                FreePascal \ Delphi WIA Implementation                       *
*                                                                             *
*  FILE: WIA_Demo_Form.pas                                                    *
*                                                                             *
*  VERSION:     1.0.1                                                         *
*                                                                             *
*  DESCRIPTION:                                                               *
*    WIA Demo Form.                                                           *
*    This is a Package demo not a full scanning application,                  *
*    see the DigIt project on my GitHub repository for that.                  *
*                                                                             *
*******************************************************************************
*                                                                             *
*  (c) 2025 Massimo Magnano                                                   *
*                                                                             *
*  See changelog.txt for Change Log                                           *
*                                                                             *
*******************************************************************************)
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
    btBrowse: TButton;
    cbTest: TCheckBox;
    cbEnumLocal: TCheckBox;
    edPath: TEdit;
    edTests: TEdit;
    ImageHolder: TImage;
    Label1: TLabel;
    Label2: TLabel;
    lbSelected: TLabel;
    lbProgress: TLabel;
    Panel1: TPanel;
    progressBar: TProgressBar;
    edPages: TSpinEdit;
    procedure btBrowseClick(Sender: TObject);
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

function SelectPath(var ADir: String): Boolean;
var
   selDir: TSelectDirectoryDialog;

begin
  try
     selDir:= TSelectDirectoryDialog.Create(FormWIADemo);
     Result:= selDir.Execute;
     if Result then ADir:= selDir.FileName;

  finally
    selDir.Free;
  end;
end;

{$else}
  {$R *.dfm}

uses FileCtrl;

function SelectPath(var ADir: String): Boolean;
begin
  Result:= SelectDirectory(ADir, [sdAllowCreate], 0);
end;
{$endif}



{ TFormWIADemo }

procedure TFormWIADemo.FormCreate(Sender: TObject);
begin
    FWia :=TWIAManager.Create;
    FWia.OnAfterDeviceTransfer:= DeviceTransferEvent;
    edPath.Text:= ExtractFilePath(ParamStr(0));
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
    progressBar.Position:= 100;
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
        btAcquire.Enabled := True;
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
  aPath:= edPath.Text;
  c:= WiaSource.DownloadNativeUI(Self.Handle, False, aPath, 'wia_demo', DownloadedFiles);
  if (c>0)
  then begin
           MessageDlg('Downloaded '+IntToStr(c)+' Files', mtInformation, [mbOk], 0);

           ImageHolder.Picture.Bitmap.LoadFromFile(DownloadedFiles[0]);
         end
    else MessageDlg('NO Files Downloaded ', mtError, [mbOk], 0);
  DownloadedFiles:= nil;
end;

procedure TFormWIADemo.btBrowseClick(Sender: TObject);
var
   ADir: String;

begin
  if SelectPath(ADir) then edPath.Text:= ADir;
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

  if (WIASource <> nil) then
  try
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

      aPath:= edPath.Text;

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

             if WIAParams[SelectedItemIndex].DataType in [wdtRAW_RGB..wdtRAW_CMYK]
             then begin
                    aFormat:= wifRAW;
                    aExt:= '.raw'
                  end
             else begin
                    aFormat:= wifBMP;
                    aExt:= '.bmp';
                  end;

             c:= WIASource.Download(aPath, 'wia_demo', aExt, aFormat);

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
                    MessageDlg('Downloaded '+IntToStr(c)+' Files on'#13#10+aPath, mtInformation, [mbOk], 0);
                    ImageHolder.Picture.Bitmap.LoadFromFile(aPath+'wia_demo'+aExt);
                  end
             else MessageDlg('NO Files Downloaded ', mtError, [mbOk], 0);
          end;
     end;

  finally
    WIASettingsSource.Free; WIASettingsSource:= Nil;
  end;
end;

end.

