program wia_console;

{$APPTYPE CONSOLE}

{$ifdef fpc}
  {$mode delphi}
{$endif}

{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Windows, ActiveX, Classes, SysUtils,
  WIA, WIA_LH, WiaDef, WIA_UI_Common, WIA_PaperSizes;

resourcestring
  rsNoWIADevicePresent = 'No Wia Devices present...';
  rsWIADeviceSelect = 'Select a Device to use :';
  rsExcCannotGetCapabilities = 'Cannot Get Capabilities for Source Item %s';


const
  Red           = 4;
  White         = 15;

type
  { TWIADemoApplication }
  TWIADemoApplication = class(TObject)
  protected
    lres: HResult;
    FWia: TWIAManager;
    WIASource: TWIADevice;
    WIAParams: TArrayWIAParams;   //for semplicity i only use the [0] array item
    WIACap: TWIAParamsCapabilities;
    WIAformat,
    WIAformatDef: TWiaImageFormat;
    WIAFormatSet: TWiaImageFormatSet;
    WIAformatStr: String;

    procedure Run;
    procedure Acquire;

    function DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                                 lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
  end;


procedure TextColor(AColor: Word);
var
  hConsole: THandle;
begin
  hConsole := GetStdHandle(STD_OUTPUT_HANDLE);
  SetConsoleTextAttribute(hConsole, AColor);
end;

function AddCharR(C: AnsiChar; const S: string; N: Integer): string;
var
  l : Integer;

begin
  Result:= S;
  for l:=Length(S) to N do Result:= Result+C;
end;


// Procedure to draw the progress bar
const
    BarWidth = 40; // Length of the progress bar in characters

procedure DrawProgressBar(ACaption:String; AProgress, ATotal: Integer);
var
   FilledWidth: Integer;
   Percent: Double;
   i: Integer;
   Bar: String;
begin
  Percent := (AProgress / ATotal) * 100;
  FilledWidth := Round((AProgress / ATotal) * BarWidth);

  // Construct the bar as string
  Bar := '';
  for i := 1 to BarWidth do
  begin
    if i <= FilledWidth then
      Bar := Bar + '#'
    else
      Bar := Bar + ' ';
  end;

  // \r returns the cursor to the start of the line to update it in-place
  Write(#13, ACaption+' [', Bar, '] ', FormatFloat('0.0', Percent), '%');
end;

function InArray(AValue: Integer; AArray: TArrayInteger): Boolean;
var
   i: Integer;

begin
  Result:= False;
  for i:= Low(AArray) to High(AArray) do
    if (AArray[i] = AValue) then begin Result:= True; Break; end;
end;

procedure UpdateWIAParams(AWIASource: TWIADevice; var AWIAParams: TWIAParams; var AWIACap: TWIAParamsCapabilities);
begin
  if (AWIASource <> nil) and (AWIASource.SelectedItem <> nil) then
  begin
    if AWIASource.GetParamsCapabilities(AWIACap) then
    begin
      if (AWIASource.SelectedItem^.ItemCategory = wicFEEDER)
      then AWIAParams:= WIACopyDefaultValues(AWIACap, waHCenter) //maybe WIACopyCurrentValues but current paper does not work
      else AWIAParams:= WIACopyDefaultValues(AWIACap);
    end
    else begin
           TextColor(Red);
           Writeln('Error: '+Format(rsExcCannotGetCapabilities, [AWIASource.SelectedItem^.Name]));
           TextColor(White);
         end;
  end;
end;

function SelectDevice(AWIAManager: TWIAManager): Integer;
var
  ch: String;
  dIndex,
  numDevices: Integer;


  function FillList: Integer;
  var
     i: Integer;
     curItem: String;
     curDevice: TWIADevice;

  begin
    Result:= AWIAManager.DevicesCount;
    if (Result > 0)
    then for i:=0 to Result-1 do
         begin
           curDevice :=AWIAManager.Devices[i];

           //Add Name Manufacturer and Type
           curItem:= curDevice.Name+' | '+curDevice.Manufacturer+' | '+WIADeviceType(curDevice.Type_);

           //if is Current Selected Scanner set selectedIndex
           if (AWIAManager.SelectedDevice <> nil) and (AWIAManager.SelectedDevice.ID = curDevice.ID)
           then begin
                  curItem:= '=> ['+IntToStr(i)+'] '+curItem
                end
           else curItem:= '   ['+IntToStr(i)+'] '+curItem;

           Writeln(curItem);
       end
    else begin
           TextColor(Red);
           Writeln(rsNoWIADevicePresent);
           TextColor(White);
         end;
  end;

begin
  Result:= -1;
  repeat
    Writeln(rsWIADeviceSelect);
    numDevices:= FillList;

    Readln(ch);
    ch:= UpperCase(ch);
    Case ch[1] of
    'B': Break;
    else begin
           try
              dIndex:= StrToInt(ch);
              if (dIndex>=0) and (dIndex<numDevices)
              then begin
                     Result:= dIndex;
                     Break;
                   end
              else raise Exception.Create('');
           except
             TextColor(Red);
             if (numDevices > 0)
             then Writeln('Enter a valid Choice [B, 0..'+IntToStr(numDevices-1)+']')
             else Writeln('Enter a valid Choice [B, <any other to List>]');
             TextColor(White);
           end;
         end;
    end;
  until False;
end;


function Settings(AWIASource: TWIADevice;
                  var ASelectedItemIndex: Integer;
                  AInitItemValues: TInitialItemValues;
                  var AParams: TArrayWIAParams;
                  AOnInitDefaultValues: TInitDefaultValuesEvent=nil): Boolean;
var
   ch: String;
   i, iIndex: Integer;
   curItem: PWIAItem;
   WIACap: TWIAParamsCapabilities;

   function SelectItem: Integer;
   var
      i: Integer;
      ch: Char;
      dIndex,
      numItems: Integer;
      curItemStr: String;
      curItem: PWIAItem;

   begin
     Result:= -1;
     Writeln;
     Writeln('Select a Sub Item to use :');
     repeat
       numItems:= AWIASource.ItemCount;

       for i:=0 to numItems-1 do
       begin
         curItem:= AWIASource.Items[i];

         if (curItem <> nil) then
         begin
           //Add Name and Type
           curItemStr:= curItem^.Name+' | '+WiaItemCategoryDescr[curItem^.ItemCategory];

           //if Current is Selected Item set Result
           if (i = AWIASource.SelectedItemIndex)
           then begin
                  Result:= i;
                  curItemStr:= '=> ['+IntToStr(i)+'] '+curItemStr
                end
           else curItemStr:= '   ['+IntToStr(i)+'] '+curItemStr;

           Writeln(curItemStr);
         end;
      end;

      Writeln('B: Back');

      Readln(ch);
      ch:= UpCase(ch);
      Case ch of
      'B': Break;
      else begin
             try
                dIndex:= StrToInt(ch);
                if (dIndex>=0) and (dIndex<numItems)
                then begin
                       Result:= dIndex;
                       Break;
                     end
                else raise Exception.Create('');
             except
               TextColor(Red);
               if (numItems > 0)
               then Writeln('Enter a valid Choice [B, 0..'+IntToStr(numItems-1)+']')
               else Writeln('Enter a valid Choice [B, <any other to List>]');
               TextColor(White);
             end;
           end;
       end;
     until False;
   end;


   procedure Papers;
   var
      paperI: TWIAPaperType;

   begin
     Writeln;
     Writeln('Paper Type:');
     i:=0;
     for paperI in WIACap.PaperTypeSet do
     if (paperI <> wptCUSTOM) then
     begin
       if (paperI = AParams[0].PaperType) then Write('=>') else Write('  ');

       Write(AddCharR(' ', '['+IntToStr(Integer(paperI))+'] '+
                      PaperTypeNameAndSize(WIASettings_Unit_cm, paperI), 40));
       inc(i);
       if not(Odd(i)) then Writeln;
     end;

     repeat
       Writeln('B: Back to Settings Menu');

       Readln(ch);
       ch:= UpperCase(ch);
       Case ch[1] of
       'B': Break;
       else begin
              try
                 AParams[0].PaperType:= TWIAPaperType(StrToInt(ch));
                 Break;

              except
                TextColor(Red);
                Writeln('Enter a valid Choice [B, 0..255]');
                TextColor(White);
              end;
            end;
       end;
     until False;
   end;

   procedure DataTypes;
   var
      dataI: TWIADataType;

   begin
     Writeln;
     Writeln('Image Data Type:');
     i:=0;
     for dataI in WIACap.DataTypeSet do
     if not(dataI in [wdtDITHER, wdtCOLOR_DITHER]) then
     begin
       if (dataI = AParams[0].DataType) then Write('=>') else Write('  ');

       Case dataI of
       wdtAUTO: Write(AddCharR(' ', '['+IntToStr(Integer(dataI))+'] '+rsAutoType, 40));
       else Write(AddCharR(' ', '['+IntToStr(Integer(dataI))+'] '+WIADataTypeDescr[dataI], 40));
       end;

       inc(i);
       if not(Odd(i)) then Writeln;
     end;

     repeat
       Writeln('B: Back to Settings Menu');

       Readln(ch);
       ch:= UpperCase(ch);
       Case ch[1] of
       'B': Break;
       else begin
              try
                 AParams[0].DataType:= TWIADataType(StrToInt(ch));
                 Break;

              except
                TextColor(Red);
                Writeln('Enter a valid Choice [B, 0..100]');
                TextColor(White);
              end;
            end;
       end;
     until False;
   end;

   procedure Orientation;
   begin
     Writeln;
     Writeln('Orientation:');
     if (AParams[0].Rotation in [wrLandscape, wrRot270])
     then Writeln(AddCharR(' ', '  [0] '+rsPortrait, 40)+'=>[1] '+rsLandscape)
     else Writeln(AddCharR(' ', '=>[0] '+rsPortrait, 40)+'  [1] '+rsLandscape);

     repeat
       Writeln('B: Back to Settings Menu');

       Readln(ch);
       ch:= UpperCase(ch);
       Case ch[1] of
       'B': Break;
       else begin
              try
                 if StrToBool(ch)
                 then AParams[0].Rotation:= wrLandscape
                 else AParams[0].Rotation:= wrPortrait;
                 Break;

              except
                TextColor(Red);
                Writeln('Enter a valid Choice [B, 0..100]');
                TextColor(White);
              end;
            end;
       end;
     until False;
   end;

   procedure Resolution;
   var
      i, dIndex: Integer;

   begin
     Writeln;
     Writeln('Resolution:');
     repeat
        if WiaCap.ResolutionRange
        then Writeln('Select a Value from '+IntToStr(WiaCap.ResolutionArray[WIA_RANGE_MIN])+' to '+IntToStr(WiaCap.ResolutionArray[WIA_RANGE_MAX])+':')
        else begin
               Writeln('Select a Value in List:');
               for i:=0 to Length(WiaCap.ResolutionArray)-1 do
               begin
                 if (WiaCap.ResolutionArray[i] = AParams[0].Resolution) then Write('=>') else Write('  ');

                 Write(AddCharR(' ', '['+IntToStr(WiaCap.ResolutionArray[i])+']', 10));

                 if ((i+1) mod 4=0) then Writeln;
               end;
               if ((i+1) mod 4<>0) then Writeln;
             end;

       Writeln('B: Back to Settings Menu');

       Readln(ch);
       ch:= UpperCase(ch);
       Case ch[1] of
       'B': Break;
       else begin
              try
                 dIndex:= StrToInt(ch);

                 if WiaCap.ResolutionRange
                 then begin
                        if not( (dIndex>= WiaCap.ResolutionArray[WIA_RANGE_MIN]) and
                              (dIndex<= WiaCap.ResolutionArray[WIA_RANGE_MAX]))
                        then raise Exception.Create('');
                      end
                 else if not(InArray(dIndex, WiaCap.ResolutionArray))
                      then raise Exception.Create('');

                 AParams[0].Resolution:= dIndex;
                 Break;

              except
                TextColor(Red);
                Writeln('Enter a valid Choice [B, <Int Number>]');
                TextColor(White);
              end;
            end;
       end;
     until False;
   end;

begin
  Result:= False;
  if not(AWIASource.GetParamsCapabilities(WIACap)) then
  begin
    Writeln;
    TextColor(Red);
    Writeln('Error: '+Format(rsExcCannotGetCapabilities, [AWIASource.SelectedItem^.Name]));
    TextColor(White);
  end
  else
  begin
  repeat
      Writeln;
      curItem:= AWIASource.SelectedItem;
      iIndex:= AWIASource.SelectedItemIndex;

      Write('Settings of '+AWIASource.Name);

      if (curItem <> nil) then
      begin
        Writeln(' - '+curItem^.Name);

        Writeln('P: Paper Type [ '+PaperTypeNameAndSize(WIASettings_Unit_cm, AParams[0].PaperType)+' ]');

        Write('O: Landscape/Portrait [');
        if (AParams[0].Rotation in [wrLandscape, wrRot270])
        then Writeln('Landscape]')
        else Writeln('Portrait]');

        Writeln('D: Image Data Type ['+WIADataTypeDescr[AParams[0].DataType]+']');

        Writeln('R: Resolution ['+IntToStr(AParams[0].Resolution)+']');
      end
      else Writeln;

      if (AWIASource.ItemCount > 1) then Writeln('I: Change Selected Item');
      Writeln('B: Back to Main Menu');

      Readln(ch);
      ch:= UpperCase(ch);
      Case ch[1] of
      'I': begin
             ASelectedItemIndex:= SelectItem;

             //Select the Item
             AWIASource.SelectedItemIndex:= ASelectedItemIndex;

             if (iIndex <> ASelectedItemIndex) then UpdateWIAParams(AWIASource, AParams[0], WIACap);
           end;

      'P': Papers;
      'O': Orientation;
      'D': DataTypes;
      'R': Resolution;
      'B': Break;
      end;
    until False;
    Result:= True;
  end;
end;

{ TWIADemoApplication }

procedure TWIADemoApplication.Run;
var
  ch: String;
  dIndex,
  iIndex: Integer;

begin
  SetLength(WIAParams, 1);

  WIA_UI_Common.WIASelectDialogFunc:= @SelectDevice;
  WIA_UI_Common.WIASettingsDialogFunc:= @Settings;

  TextColor(White);
  Writeln('WIA Console Demo (c) 2026 Massimo Magnano');
  try
     WIASource:= FWia.SelectedDevice;
     UpdateWIAParams(WIASource, WIAParams[0], WIACap);

     repeat
       Writeln;
       Writeln('D: Select a Device');

       if (WIASource <> nil) then
       begin
         if (WIASource.SelectedItem <> nil)
         then begin
                Writeln('A: Acquire from '+WIASource.Name+' - '+WIASource.SelectedItem^.Name);
                Writeln('S: Settings');
              end
         else Writeln('A: Acquire');
       end;
       Writeln('X: Exit');

       Readln(ch);
       ch:= UpperCase(ch);
       Case ch[1] of
       'D': begin
              dIndex:= FWia.SelectDeviceDialog;
              if (dIndex>=0) then
              begin
                //Select the Device
                FWia.SelectedDeviceIndex:= dIndex;
                WIASource:= FWia.SelectedDevice;
                UpdateWIAParams(WIASource, WIAParams[0], WIACap);
              end;
            end;
       'A': Acquire;
       'S': if (WIASource<>nil) then WIASource.SettingsDeviceDialog(iIndex, initParams, WIAParams);
       'X': Break;
       end;
     until False;

  finally
    // stop program loop
  end;
end;

procedure TWIADemoApplication.Acquire;
var
   aPath, aExt: String;
   c: Integer;
   aFormat: TWIAImageFormat;

begin
  if (WIASource <> nil) then
  try
    //Set Scanner Setting to use
    if (WIASource.SelectedItem <> nil) then WIASource.SetParams(WIAParams[0]);

    aPath:= '';

    aFormat:= wifBMP;
    aExt:= '.bmp';

    c:= WIASource.Download(aPath, 'wia_console', aExt, aFormat);

    Writeln;

    if (c>0)
    then Writeln('Downloaded '+IntToStr(c)+' Files')
    else begin
           TextColor(Red);
           Writeln('Error: NO Files Downloaded');
           TextColor(White);
         end;

  finally
  end;
end;

constructor TWIADemoApplication.Create;
begin
  inherited Create;

  FWia :=TWIAManager.Create;
  FWia.OnAfterDeviceTransfer:= DeviceTransferEvent;
end;

destructor TWIADemoApplication.Destroy;
begin
  WIAParams:= nil;
  if (FWIA<>nil) then FWia.Free;

  inherited Destroy;
end;

function TWIADemoApplication.DeviceTransferEvent(AWiaManager: TWIAManager; AWiaDevice: TWIADevice;
                                          lFlags: LONG; pWiaTransferParams: PWiaTransferParams): Boolean;
begin
  Result:= True;

  if (pWiaTransferParams <> nil) then
  Case pWiaTransferParams^.lMessage of
  WIA_TRANSFER_MSG_STATUS: begin
    DrawProgressBar('Downloading '+AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count)+' file',
                    pWiaTransferParams^.lPercentComplete, 100);
  end;
  WIA_TRANSFER_MSG_END_OF_STREAM: begin
    Write(#13, AddCharR(' ',
                    'Downloaded '+AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count)+' file', 80));
  end;
  WIA_TRANSFER_MSG_END_OF_TRANSFER: begin
    DrawProgressBar('Downloading '+AWiaDevice.Download_FileName+AWiaDevice.Download_Ext+'  : '+IntToStr(AWiaDevice.Download_Count)+' file',
                    100, 100);
  end;
  WIA_TRANSFER_MSG_DEVICE_STATUS: begin
  end;
  WIA_TRANSFER_MSG_NEW_PAGE: begin
  end
  else begin
         //Writeln('WIA_TRANSFER_MSG_'+IntToHex(pWiaTransferParams^.lMessage)+' : '+IntToStr(pWiaTransferParams^.lPercentComplete)+'% err='+IntToHex(pWiaTransferParams^.hrErrorStatus));
       end;
  end;
end;


var
  Application: TWIADemoApplication;
  hr : HRESULT;

{$R *.res}

begin
  //initialize COM  really only Delphi need this
  hr := ActiveX.CoInitialize(nil);
  if (FAILED (hr)) then begin
    TextColor(Red);
    Writeln('Failed to initialize COM');
    TextColor(White);
    exit;
  end;

  Application:= TWIADemoApplication.Create;
  Application.Run;
  Application.Free;

  ActiveX.CoUninitialize();
end.

