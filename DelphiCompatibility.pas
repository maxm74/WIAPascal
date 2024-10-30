unit DelphiCompatibility;

interface

{$ifndef fpc}
type
{$ifdef CPUX64}
  PtrInt = Int64;
  PtrUInt = UInt64;
{$else}
  PtrInt = longint;
  PtrUInt = Longword;
{$endif}
const
  AllowDirectorySeparators : set of AnsiChar = ['\','/'];
  DirectorySeparator = '\';
{$endif}

implementation

end.
