{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit WIAPascal_pkg;

{$warn 5023 off : no warning about unused units}
interface

uses
  wia_lh, WIA_TLB, WiaDef, WiaDevD, LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('WIA_TLB', @WIA_TLB.Register);
end;

initialization
  RegisterPackage('WIAPascal_pkg', @Register);
end.
