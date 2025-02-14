(****************************************************************************
*
*  FILE: DelphiCompatibility.pas
*
*  VERSION:     1.0
*
*  DESCRIPTION:
*    Types and Defines to make code compatible between FreePascal and Delphi.
*
*****************************************************************************
*
*  2024 Massimo Magnano
*
*
*****************************************************************************)
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
