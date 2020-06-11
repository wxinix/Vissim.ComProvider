unit Vissim.ComProvider.Utilities;

interface

uses
  WinApi.Windows;

function GetVissimMainWindowHandle(unk: IUnknown): HWND; stdcall;
  external 'Vissim.ComProvider.Utilities.dll' name 'GetVissimMainWindowHandle';

procedure HideVissim(aVissim: IUnknown); stdcall;
  external 'Vissim.ComProvider.Utilities.dll' name 'HideVissim';

procedure ShowVissim(aVissim: IUnknown); stdcall;
  external 'Vissim.ComProvider.Utilities.dll' name 'ShowVissim';

implementation

end.
