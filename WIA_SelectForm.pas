(******************************************************************************
*                FreePascal \ Delphi WIA Implementation                       *
*                                                                             *
*  FILE: WIA_SelectForm.pas                                                   *
*                                                                             *
*  VERSION:     1.0.1                                                         *
*                                                                             *
*  DESCRIPTION:                                                               *
*    WIA Select Device Dialog.                                                *
*                                                                             *
*******************************************************************************
*                                                                             *
*  (c) 2025 Massimo Magnano                                                   *
*                                                                             *
*  See changelog.txt for Change Log                                           *
*                                                                             *
*******************************************************************************)

unit WIA_SelectForm;

{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  ComCtrls, StdCtrls, WIA;

resourcestring
  rsNoWIADevicePresent = 'No Wia Devices present...';

type
  { TWIASelectForm }
  TWIASelectForm = class(TForm)
    btCancel: TBitBtn;
    btRefresh: TBitBtn;
    btOk: TBitBtn;
    lvSources: TListView;
    Panel1: TPanel;
    panelButtons: TPanel;
    procedure btRefreshClick(Sender: TObject);
  private
    WIAManager: TWIAManager;

    procedure FillList;

  public
     class function Execute(AWIAManager: TWIAManager): Integer;
  end;

var
  WIASelectForm: TWIASelectForm = nil;

implementation

{$ifdef FPC}
  {$R *.lfm}
{$else}
  {$R *.dfm}
{$endif}


{ TWIASelectForm }

procedure TWIASelectForm.btRefreshClick(Sender: TObject);
begin
  WIAManager.RefreshDeviceList;
  FillList;
  lvSources.SetFocus;
end;

procedure TWIASelectForm.FillList;
var
   i,
   txtW,
   numDevices,
   selectedIndex: Integer;
   curItem: TListItem;
   curDevice: TWIADevice;

begin
  selectedIndex:=-1;
  lvSources.Clear;

  numDevices:= WIAManager.DevicesCount;
  if (numDevices > 0)
  then begin
         for i:=0 to numDevices-1 do
         begin
           curDevice :=WIAManager.Devices[i];
           curItem :=lvSources.Items.Add;

           //Add Name Colums and increase width if necessary (since MinWidth/AutoSize don't work as expected)
           curItem.Caption:=curDevice.Name;
           txtW:= lvSources.Canvas.TextWidth(curDevice.Name)+16;
           if (lvSources.Columns[0].Width < txtW) then lvSources.Columns[0].Width:= txtW;

           //Add Manufacturer
           curItem.SubItems.Add(curDevice.Manufacturer);
           txtW:= lvSources.Canvas.TextWidth(curDevice.Manufacturer)+16;
           if (lvSources.Columns[1].Width < txtW) then lvSources.Columns[1].Width:= txtW;

           //Add Type
           curItem.SubItems.Add(WIADeviceType(curDevice.Type_));
           txtW:= lvSources.Canvas.TextWidth(WIADeviceType(curDevice.Type_))+16;
           if (lvSources.Columns[2].Width < txtW) then lvSources.Columns[2].Width:= txtW;

           //if is Current Selected Scanner set selectedIndex
           if (WIAManager.SelectedDevice <> nil) and (WIAManager.SelectedDevice.ID = curDevice.ID)
           then selectedIndex :=curItem.Index;
         end;

         //Select Current Scanner
         if (selectedIndex > -1)
         then lvSources.ItemIndex :=selectedIndex
         else lvSources.ItemIndex :=0;
       end
  else MessageDlg(rsNoWIADevicePresent, mtError, [mbOk], 0);
end;

class function TWIASelectForm.Execute(AWIAManager: TWIAManager): Integer;
begin
  Result :=-1;
  if (WIASelectForm = nil)
  then WIASelectForm :=TWIASelectForm.Create(nil);

  if (WIASelectForm <> nil) then
  with WIASelectForm do
  try
    WIAManager:= AWIAManager;
    FillList;

    if (ShowModal = mrOk)
    then Result:= lvSources.ItemIndex;

  finally
     WIASelectForm.Free; WIASelectForm:= nil;
  end;
end;

end.

