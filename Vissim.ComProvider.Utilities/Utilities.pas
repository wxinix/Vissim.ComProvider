// MIT License
// Copyright (c) Wuping Xin 2020.
//
// Permission is hereby  granted, free of charge, to any  person obtaining a copy
// of this software and associated  documentation files (the "Software"), to deal
// in the Software  without restriction, including without  limitation the rights
// to  use, copy,  modify, merge,  publish, distribute,  sublicense, and/or  sell
// copies  of  the Software,  and  to  permit persons  to  whom  the Software  is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE  IS PROVIDED "AS  IS", WITHOUT WARRANTY  OF ANY KIND,  EXPRESS OR
// IMPLIED,  INCLUDING BUT  NOT  LIMITED TO  THE  WARRANTIES OF  MERCHANTABILITY,
// FITNESS FOR  A PARTICULAR PURPOSE AND  NONINFRINGEMENT. IN NO EVENT  SHALL THE
// AUTHORS  OR COPYRIGHT  HOLDERS  BE  LIABLE FOR  ANY  CLAIM,  DAMAGES OR  OTHER
// LIABILITY, WHETHER IN AN ACTION OF  CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE  OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

namespace Vissim.ComProvider.Utilities;

uses
  rtl, RemObjects.Elements.System;

type
  {$HIDE H6} {$HIDE H7} {$HIDE H8}
  // A hack to sniff the PID out of the header. Only works for PID (<= 65535) of local server.
  ComObjRefHeader = packed record
    Signature: DWORD; 	        // Signature "MEOW", Offset 0, Size 4
    Flag: DWORD;                // Flag indicating the kind of structure. 1: Standard
    IID: array[0..15] of Byte;  // Interface Identifier
    ReservedFlags: DWORD;       // Flags reserved for the system
    RefCount: DWORD;            // Reference count
    OXID: array[0..7] of Byte;  // Object Exporter Identifier
    OID: array[0..7] of Byte;   // Object Identifier
    IPID: array[0..15] of Byte; // Interface Pointer Identifier
  end;
  {$SHOW H6} {$SHOW H7} {$SHOW H8}

type
  EnumWindowsData = record
    Pid: DWORD;
    Hnd: HWND;
  end;

method Succeeded(aStatus: HRESULT): Boolean;
begin
  result := aStatus and HRESULT($80000000) = 0;
end;

method Failed(aStatus: HRESULT): Boolean;
begin
  result := aStatus and HRESULT($80000000) <> 0;
end;

method CoGetServerPID(const aUnk: IUnknown; var aPid: DWORD): HRESULT;
begin
  if not assigned(aUnk) then exit E_INVALIDARG;

  // Check if not standard proxy. The ComObjRefHeader packet must be in the standard format.
  var proxyManager: IUnknown;
  result := aUnk.QueryInterface(@IID_IProxyManager, ^^Void(@proxyManager));
  if Failed(result) then exit;

  // Marshall the interface to get a new OBJREF. The CreateStreamOnHGlobalfunction creates a stream
  // object that uses an HGLOBAL memory handle to store the stream contents. This object is the
  // OLE-provided implementation of the IStream interface.
  var marshalStream: IStream;
  result := CreateStreamOnHGlobal(
    nil,           // HGLOBAL memory handle; nil to allocate a new one.
    True,          // Whether automatically free the handle when releasing stream object.
    @marshalStream // Address of IStream variable for the new stream object.
   );
  if Failed(Result) then exit;

  // Writes the passed-in unk interface object into a stream with the data required to
  // initialize a proxy object in seperate client process.
  result := CoMarshalInterface(
      marshalStream,
      @IID_IUnknown,
      aUnk,
      DWORD(MSHCTX.MSHCTX_INPROC),      // Unmarshaling to be done in another apartment in the same process.
      nil,                              // Destination context.
      DWORD(MSHLFLAGS.MSHLFLAGS_NORMAL) // Marshaling for intf pointer passed from one process to another.
      );
  if Failed(result) then exit;

  // Using the created stream, restore to a raw pointer.
  var hg: HGLOBAL;
  result := GetHGlobalFromStream(marshalStream, @hg);

  if Succeeded(result) then begin
    try
      result := RPC_E_INVALID_OBJREF; // Start out pessimistic
      var objRefHdr: ^ComObjRefHeader := ^ComObjRefHeader(GlobalLock(hg));

      if assigned(objRefHdr) then begin // Verify that the signature is MEOW
        if objRefHdr^.Signature = $574F454D then begin
          aPid := objRefHdr^.IPID[4] + objRefHdr^.IPID[5] shl 8; // Make WORD for PID.
          result := S_OK;
        end;
      end;
    finally
      GlobalUnlock(hg);
    end;
  end;

  // Rewind stream and release marshaled data to keep refcount in order. The Seek method changes
  // the seek pointer to a new location. The new location is relative to either the beginning of
  // the stream (STREAM_SEEK_SET), the end of the stream (STREAM_SEEK_END), or the current seek 
  // pointer (STREAM_SEEK_CUR).
  var libNewPosition: ULARGE_INTEGER;
  var dlibMove: LARGE_INTEGER := new LARGE_INTEGER(QuadPart := 0); // or default(LARGE_INTEGER)

  marshalStream.Seek(
    dlibMove,                           // Offset to ref point, which is specified by 2nd arg.
    DWORD(STREAM_SEEK.STREAM_SEEK_SET), // Use stream begin as the origin.
    @libNewPosition                     // New position of the seek pointer.
    );

  // Destroys a previously marshaled data packet.
  CoReleaseMarshalData(marshalStream);
end;

method EnumWindowsCallback(aHnd: HWND; aLParam: LPARAM): Boolean; stdcall;
begin
  var data := ^EnumWindowsData(aLParam);
  var pid: DWORD := 0;

  // Retrieves the id of the thread that created the specified window and, optionally,
  // the id of the parent process.
  GetWindowThreadProcessId(aHnd, @pid); // Discard the returned thread id.

  if ((data^.Pid <> pid) or (not IsMainWindow(aHnd))) then begin
    result := true;
  end else begin
    if not assigned(data.Hnd) then data.Hnd := aHnd;
    result := false; 
  end;
end;

method IsMainWindow(aHnd: HWND): Boolean;
begin
  // This is good enough for searching Vissim main window
  result := (GetWindow(aHnd, GW_OWNER) = nil) and IsWindowVisible(aHnd);
end;

method FindMainWindow(aPid: DWORD): HWND;
begin
  if aPid = 0 then exit nil;
  var data: EnumWindowsData := new EnumWindowsData (Pid := aPid, Hnd := nil);
  EnumWindows(@EnumWindowsCallback, LPARAM(@data));
  result := data.Hnd;
end;

type
  VissimComProviderHelper = public class
  private
    class var fHwndCache: Dictionary<NativeUInt, HWND> := nil;
    
  public
    class method ShowVissim(aVissim: IUnknown);
    begin
      var hnd := GetVissimMainWindowHandle(aVissim);
      if assigned(hnd) then ShowWindow(hnd, SW_NORMAL); 
    end;

    class method HideVissim(aVissim: IUnknown);
    begin
      var hnd := GetVissimMainWindowHandle(aVissim);
      if assigned(hnd) then ShowWindow(hnd, SW_HIDE); 
    end;

    class method GetVissimMainWindowHandle(aVissim: IUnknown): HWND;
    begin
      var key := NativeUInt(^Void(aVissim));

      if fHwndCache.ContainsKey(key) then begin
        result := fHwndCache[key];
      end else begin   
        var pid: DWORD := GetVissimSeverPID(aVissim);
        result := FindMainWindow(pid);
        if assigned(result) then fHwndCache.Add(key, result);        
      end;
    end;

    class method GetVissimSeverPID(aVissim: IUnknown): DWORD;
    begin
      result := 0;
      CoGetServerPID(aVissim, var result);
    end;

    class method Initialize;
    begin
      fHwndCache := new Dictionary<NativeUInt, HWND>;
    end;
  end;

end.

