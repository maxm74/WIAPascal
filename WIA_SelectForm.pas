unit WIA_SelectForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Grids, ExtCtrls, Buttons,
  ComCtrls, WIA, LResources;

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

{$R *.lfm}

{ TWIASelectForm }

procedure TWIASelectForm.btRefreshClick(Sender: TObject);
begin
  WIAManager.ClearDeviceList;
  FillList;
end;

procedure TWIASelectForm.FillList;
var
   i, selectedIndex: Integer;
   curItem: TListItem;
   curDevice: TWIADevice;

begin
  selectedIndex:=-1;
  lvSources.Clear;

  for i:=0 to WIAManager.DevicesCount-1 do
  begin
    curDevice :=WIAManager.Devices[i];
    curItem :=lvSources.Items.Add;
    curItem.Caption:=curDevice.Name;
    //curItem.SubItems.Add(curDevice.Type_);
    curItem.SubItems.Add(curDevice.Manufacturer);

    //if is Current Selected Scanner set selectedIndex
    if (WIAManager.SelectedDevice <> nil) and (WIAManager.SelectedDevice.ID = curDevice.ID)
    then selectedIndex :=curItem.Index;
  end;

  //Select Current Scanner
  if (selectedIndex>-1)
  then lvSources.ItemIndex :=selectedIndex
  else lvSources.ItemIndex :=0;
end;

class function TWIASelectForm.Execute(AWIAManager: TWIAManager): Integer;
begin
  Result :=-1;
  if (WIASelectForm = nil)
  then WIASelectForm :=TWIASelectForm.Create(nil);

  with WIASelectForm do
  begin
    {$IFDEF DELPHI_2006_UP}
      PopupMode := pmAuto;
    {$ENDIF}

    WIAManager:= AWIAManager;
    FillList;

    if (lvSources.Items.Count = 0)
    then MessageDlg('WIA', 'No Wia Device present...', mtError, [mbOk], 0)
    else if (ShowModal = mrOk)
         then Result:= lvSources.ItemIndex;
  end;
end;

end.

