//
//  File thumbnail module.
//
// Type: Windows thumbnail.
// Author: 2023 Wei-Lun Huang
// Description: Extract file thumbnail.
//
// Features: Please see the functions in interface.
//
// Last modified date: Oct 9, 2023.


{ Example:
uses FileThumbnail;

procedure GetFileThumb;
var
  Thumb: TFileThumb;
  Icon: TIcon;
begin
  if GetShellIcon(Icon, JvFilenameEdit1.FileName, _IconBest.TypicallyPixels) then
    try
      Image1.Picture.Assign(Icon);
    finally
      Icon.Free;
    end
  else
    Image1.Picture.Assign(nil);
end;
}

unit Winapi.FileThumbnail;

{$IF NOT DEFINED(DEBUG)}
  {$R-,Q-,V-} // RangeChecks, OverFlowChecks, VarsStringChecks
  {$T-,B-,U-} // TypedAddress, BoolEval, SafeDivide
  {$INLINE ON} // ON|OFF|AUTO
  {$O+} // Optimization
{$ENDIF}


interface

uses
  Winapi.Windows, Winapi.ActiveX, Winapi.ShellApi, Winapi.commCtrl, Winapi.ShlObj,
  System.SysUtils, Vcl.Graphics,

// Necessary System.Win.ComObj.InitComObj()
// Queue CoInitializeEx into InitProc to execute when the initiation.
  System.Win.ComObj;

// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
//
// Begin Windows API interface definition.
//

//
// IShellItemImageFactory interface.
//
// Most of them are already defined in Winapi.ShlObj.
//
// IShellItemImageFactory::GetImage method flags.
// https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellitemimagefactory-getimage
const
//  SIIGBF_RESIZETOFIT    = $00000000; // Defined in Winapi.ShlObj.
//  SIIGBF_BIGGERSIZEOK   = $00000001; // ...
//  SIIGBF_MEMORYONLY     = $00000002; // ...
//  SIIGBF_ICONONLY       = $00000004; // ...
//  SIIGBF_THUMBNAILONLY  = $00000008; // ...
//  SIIGBF_INCACHEONLY    = $00000010; // ...
  SIIGBF_CROPTOSQUARE   = $00000020;
  SIIGBF_WIDETHUMBNAILS = $00000040;
  SIIGBF_ICONBACKGROUND = $00000080;
  SIIGBF_SCALEUP        = $00000100;
// End of the IShellItemImageFactory definition.


//
// IImageList interface.
//
// https://learn.microsoft.com/en-us/windows/win32/api/commoncontrols/nn-commoncontrols-iimagelist
const
  IID_IImageList = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';
type
  IImageList = interface(IUnknown)
    [IID_IImageList]
    function Add(Image, Mask: HBITMAP; var Index: Integer): HRESULT; stdcall;
    function ReplaceIcon(IndexToReplace: Integer; Icon: HICON; var Index: Integer): HRESULT; stdcall;
    function SetOverlayImage(iImage: Integer; iOverlay: Integer): HRESULT; stdcall;
    function Replace(Index: Integer; Image, Mask: HBITMAP): HRESULT; stdcall;
    function AddMasked(Image: HBITMAP; MaskColor: COLORREF; var Index: Integer): HRESULT; stdcall;
    function Draw(var DrawParams: TImageListDrawParams): HRESULT; stdcall;
    function Remove(Index: Integer): HRESULT; stdcall;
    function GetIcon(Index: Integer; Flags: UINT; var Icon: HICON): HRESULT; stdcall;
    function GetImageInfo(Index: Integer; var ImageInfo: TImageInfo): HRESULT; stdcall;
    function Copy(iDest: Integer; SourceList: IUnknown; iSource: Integer; Flags: UINT): HRESULT; stdcall;
    function Merge(i1: Integer; List2: IUnknown; i2, dx, dy: Integer; ID: TGUID; out ppvOut): HRESULT; stdcall;
    function Clone(ID: TGUID; out ppvOut): HRESULT; stdcall;
    function GetImageRect(Index: Integer; out rc: TRect): HRESULT; stdcall;
    function GetIconSize(var cx, cy: Integer): HRESULT; stdcall;
    function SetIconSize(cx, cy: Integer): HRESULT; stdcall;
    function GetImageCount(var Count: Integer): HRESULT; stdcall;
    function SetImageCount(NewCount: UINT): HRESULT; stdcall;
    function SetBkColor(BkColor: COLORREF; var OldColor: COLORREF): HRESULT; stdcall;
    function GetBkColor(var BkColor: COLORREF): HRESULT; stdcall;
    function BeginDrag(iTrack, dxHotSpot, dyHotSpot: Integer): HRESULT; stdcall;
    function EndDrag: HRESULT; stdcall;
    function DragEnter(hWndLock: HWND; x, y: Integer): HRESULT; stdcall;
    function DragLeave(hWndLock: HWND): HRESULT; stdcall;
    function DragMove(x, y: Integer): HRESULT; stdcall;
    function SetDragCursorImage(Image: IUnknown; iDrag, dxHotSpot, dyHotSpot: Integer): HRESULT; stdcall;
    function DragShowNoLock(fShow: BOOL): HRESULT; stdcall;
    function GetDragImage(var CurrentPos, HotSpot: TPoint; ID: TGUID; out ppvOut): HRESULT; stdcall;
    function GetItemFlags(i: Integer; var dwFlags: DWORD): HRESULT; stdcall;
    function GetOverlayImage(iOverlay: Integer; var iIndex: Integer): HRESULT; stdcall;
  end;
// End of the IImageList definition.

// End Windows API interface definition.
//
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =



// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
//
// Begin function implemented.
//

// About TIconSizeMode, see the parameter iImageList of SHGetImageList function.
// https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shgetimagelist
  TIconSizeType = (
    _IconSmall,  // Typically 16x16 pixels
    _IconLarge,  // Typically 32x32 pixels
    _IconExtra,  // Typically 48x48 pixels
    _IconJumbo); // Typically 256x256 pixels

  TSysImageList = record
// The record will automatically handle the IImageList reference count and
// release it automatically, so Obj does not require additional processing.
    Obj: IImageList; // IImageList, specify nil to release the interface.
    Size: TSize;     // Icon actual size.
  end;
  PSysImageList = ^TSysImageList;

  TSysImageLists = array[TIconSizeType] of TSysImageList;

  TIconSizeTypeHelper = record helper for TIconSizeType
  protected
// These functions obtain default(const values).
    function GetFlag: Integer; inline;       // Read cIconTypically.iImageList
    function GetId: string; inline;          // Read cIconTypically.Id
    function GetString: string; inline;      // Read cIconTypically.Name
    function GetTypicallySize: Word; inline; // Read cIconTypically.Size
    function GetTypicallyPixels: TSize;      // Read cIconTypically.Size and create TSize

// Will try to call CatchImageList(TIconSizeType) to get the image size of the ImageList.
    function GetSysImageList: TSysImageList; // Read gImageLists
    function GetImageList: IImageList;       // Read gImageLists.Obj
    function GetActualPixels: TSize;         // Read gImageLists.Size
  public
    property Flag: Integer read GetFlag;
    property Id: string read GetId;
    property Str: string read GetString;
    property TypicallySize: Word read GetTypicallySize;
    property TypicallyPixels: TSize read GetTypicallyPixels;
    property SysImageList: TSysImageList read GetSysImageList;
    property ImageList: IImageList read GetImageList;
    property ActualPixels: TSize read GetActualPixels;
  end;

  TIconSizeTypes = set of TIconSizeType;

  TImageListRetrieves = set of (
    ILR_ActualSize, // After obtaining all list of icons of sizes, choose the one with the highest actual resolution.
    ILR_NoKeep,     // Unkeep the interface for image lists of other sizes.
    ILR_Renew);     // Release the existing list and get the latest list.

  TExtractIconFlags = set of (EIF_SimulateDoc, EIF_IgnoreLarge, EIF_IgnoreSmall);

  TImageSize = (ImageSize_Non, ImageSize_Large, ImageSize_Small);

  TIconSizeInfo = record
    iImageList: Integer; // This flag is parameter iImageList of SHGetImageList.
    Size: Word;          // Icon typically size.
    Id: string;          // Name defined of image list flag in Windows SDK.
    Name: string;        // Simple name.
  end;
  PIconSizeInfo = ^TIconSizeInfo;

const
  cColorDepth = 32; // Color depth, default 32.

// TIconSizeMode
  _IconBase = _IconSmall;
  _IconBest = _IconJumbo; // SHIL_LAST

//  _ILR_Default: TImageListRetrieves = [ILR_ActualSize];

// The size is normally pixels.
// However, the icon size can be change by system setting or by user.
// when processing the CatchImageList series of functions,
// the actual size will be copied to the implementation variable gImageLists.
  cIconTypically: array[TIconSizeType] of TIconSizeInfo = (
    (iImageList: SHIL_SMALL;      Size:  16; ID: 'SHIL_SMALL';      Name: 'Small'),
    (iImageList: SHIL_LARGE;      Size:  32; ID: 'SHIL_LARGE';      Name: 'Large'),
    (iImageList: SHIL_EXTRALARGE; Size:  48; ID: 'SHIL_EXTRALARGE'; Name: 'Extra'),
    (iImageList: SHIL_JUMBO;      Size: 256; ID: 'SHIL_JUMBO';      Name: 'Jumbo')
  );

// Get the mode by ImageList category flag. (see value cIconTypically.iImageList).
function SysImageListFlagToIconSizeType(Flag: Integer): TIconSizeType; inline;

//
// Get pixels of icon.
//
function GetIconTypicallySizeByFlag(out Size: Word; Flag: Integer): Boolean; overload;
function GetIconActualSizeByFlag(out Size: TSize; Flag: Integer): Boolean; overload;
function GetIconSizeByFlag(out TypicallySize: Word; out ActualSize: TSize; Flag: Integer): Boolean; overload;

//
// Convert.
//
function ConvertToIcon(hbmColor: HBITMAP; hbmMask: HBITMAP = 0): TIcon; overload;
// If the size of the image is smaller than the specified size(ASize),
// the base of the image is expanded (not enlarged the image) and the image is centered.
function ConvertToIcon(hbmColor: HBITMAP; hbmMask: HBITMAP; ASize: TSize): TIcon; overload;

//
// Extract thumbnail, file association icon or file icon.
//
// GetShellIcon, get image by the IShellItemImageFactory interface. (Like Windows file explorer)

// GetShellIcon - Get the shell icon image.
function GetShellIcon(out Bitmap: HBITMAP; const AFilename: string; ASize: TSize; Flags: UINT = SIIGBF_RESIZETOFIT): Boolean; overload;
// GetShellIcon out TIcon, this is implemented with GetSysIcon out Bitmap.
function GetShellIcon(out Bitmap: TBitmap; const AFilename: string; ASize: TSize; Flags: UINT = SIIGBF_RESIZETOFIT): Boolean; overload;
function GetShellIcon(out Icon: TIcon; const AFilename: string; ASize: TSize; Flags: UINT = SIIGBF_RESIZETOFIT): Boolean; overload;

//
// Extract thumbnail from video or image file only.
//
// GetThumbnail - Get image thumbnail of the content form video or image file.
function GetThumbnail(out Bitmap: HBITMAP; const AFilename: string; ASize: TSize; ColorDepth: DWORD = cColorDepth): Boolean; overload;
function GetThumbnail(out Bitmap: TBitmap; const AFilename: string; ASize: TSize; ColorDepth: DWORD = cColorDepth): Boolean; overload;
function GetThumbnail(out Icon: TIcon; const AFilename: string; ASize: TSize; ColorDepth: DWORD = cColorDepth): Boolean; overload;


// Extract file icon image only.
//
// GetFileIcon - Extract icon by index in a file.
function GetFileIcon(out Icon: HICON; const AFilename: string; Index: Integer = 0; AMaxSize: Word = 0; AMinSize: Word = 0; ExtractFlags: TExtractIconFlags = []): TImageSize; overload;
function GetFileIcon(out Icon: TIcon; const AFilename: string; Index: Integer = 0; AMaxSize: Word = 0; AMinSize: Word = 0; ExtractFlags: TExtractIconFlags = []): TImageSize; overload;
function GetFileIcon(out Bitmap: TBitmap; const AFilename: string; Index: Integer = 0; AMaxSize: Word = 0; AMinSize: Word = 0; ExtractFlags: TExtractIconFlags = []): TImageSize; overload;

//
// Extract file association icon image only.

// GetSysImageListIcon (Get system image list icon):
// Usually files or systems have 16 and 32 standard size icon resources. if only want to
// get 32 or 16 standard size file icons, this should be the fastest and most appropriate.
// because it is image list through the system icons cache.
//
// This is use image list of file association of system. if the size of the icon is smaller
// than the Icon size in the image list, the image will be fixed to the upper left of the Icon.
//
// Through SHGetFileInfo and SHGetImageList API's.
// So, GetSysImageListIcon will call CatchImageList to SHGetImageList to get the image list with the file association.
function GetSysImageListIcon(out Icon: HICON; const AFilename: string; SizeType: TIconSizeType = _IconBest): Boolean; overload;
function GetSysImageListIcon(out Icon: TIcon; const AFilename: string; SizeType: TIconSizeType = _IconBest): Boolean; overload;
function GetSysImageListIcon(out Bitmap: TBitmap; const AFilename: string; SizeType: TIconSizeType = _IconBest): Boolean; overload;

//
// Check if the system image list has been obtained.
//
function SysImageListLoaded(IconSizeType: TIconSizeType): Boolean; overload;
function SysImageListLoaded: TIconSizeTypes; overload;

//
// [Automatically]
// The following is the automatically triggered functions, no additional calls are necessary.
//

//
// [Automatically] ImageList interface operation.
//
function GetImageListInterface(IconSizeType: TIconSizeType): IImageList; inline;
procedure FreeSysImageList(var SysImageList: TSysImageList); overload; inline;
procedure FreeSysImageList(IconSizeType: TIconSizeType); overload;
procedure FreeSysImageListAll;

//
// [Automatically] Catch icon list of the ImageList interface from system.
//
// See the variable gImageLists and gImageListMode in implementation.
function CatchImageList(SizeType: TIconSizeType = _IconBest; ForceLatest: Boolean = False): Boolean; //overload;
function CatchImageListAll(ForceLatest: Boolean = False): TIconSizeTypes; overload;
function CatchImageListAll(out ActualBest: TIconSizeType; ForceLatest: Boolean = False): TIconSizeTypes; overload;
function CatchImageListBest(out ASysImageList: TSysImageList; Flags: TImageListRetrieves = [ILR_ActualSize]): Boolean;

//
// [Automatically] Get ImageList form variable gImageLists in implementation.
//
function SysImageList(SizeType: TIconSizeType; ExceptNil: Boolean = True): TSysImageList; overload;
function SysImageList: TSysImageList; overload;
function SysImageLists: TSysImageLists; inline;

implementation

//uses
//  Debug;

var
// The TSysImageList.Obj type is an interface,
// so it cannot be used as a general variable type or a memory block filled
// with 0 (erased) when released. It needs to be specified as nil to make the
// interface automatically released.
  gImageLists: array[TIconSizeType] of TSysImageList; // TSysImageList.Obj
  gImageListMode: TIconSizeType = _IconBest; // IImageList size type.


//
// public functions
//

function SysImageListFlagToIconSizeType(Flag: Integer): TIconSizeType;
begin
  case Flag of
    SHIL_LARGE     : Result := _IconLarge;
    SHIL_SMALL     : Result := _IconSmall;
    SHIL_EXTRALARGE: Result := _IconExtra;
    SHIL_JUMBO     : Result := _IconJumbo;
    else raise Exception.Create('Invalid flag of ImageList type.');// Abort;
  end;
end;

function GetIconTypicallySizeByFlag(out Size: Word; Flag: Integer): Boolean;
var
  SizeType: TIconSizeType;
begin
  try
    SizeType := SysImageListFlagToIconSizeType(Flag);
  except
    Exit(False);
  end;
  Size := cIconTypically[SizeType].Size;
  Result := True;
end;

function GetIconActualSizeByFlag(out Size: TSize; Flag: Integer): Boolean;
var
  SizeType: TIconSizeType;
begin
  try
    SizeType := SysImageListFlagToIconSizeType(Flag);
  except
    Exit(False);
  end;
  Size := gImageLists[SizeType].Size;
  Result := True;
end;

function GetIconSizeByFlag(out TypicallySize: Word; out ActualSize: TSize; Flag: Integer): Boolean;
var
  SizeType: TIconSizeType;
begin
  try
    SizeType := SysImageListFlagToIconSizeType(Flag);
  except
    Exit(False);
  end;
  TypicallySize := cIconTypically[SizeType].Size;
  ActualSize := gImageLists[SizeType].Size;
  Result := True;
end;

function GetPixels(Size: TSize): Integer;
begin
  if (Size.cx < Word.MinValue) or (Size.cx > Word.MaxValue) or
     (Size.cy < Word.MinValue) or (Size.cy > Word.MaxValue) then
    raise Exception.Create('Size exceeds range of type Word');
  Result := Size.cx * Size.cy;
end;

function SysImageListLoaded(IconSizeType: TIconSizeType): Boolean;
begin
  Result := Assigned(gImageLists[IconSizeType].Obj);
end;

function SysImageListLoaded: TIconSizeTypes;
var
  SizeType: TIconSizeType;
begin
  Result := [];
  for SizeType := _IconBase to _IconBest do
    if Assigned(gImageLists[SizeType].Obj) then
      Include(Result, SizeType);
end;

function GetImageListInterface(IconSizeType: TIconSizeType): IImageList;
begin
  if SHGetImageList(cIconTypically[IconSizeType].iImageList, IImageList, Pointer(Result)) <> S_OK then
    Result := nil;
end;

procedure FreeSysImageList(var SysImageList: TSysImageList);
begin
  SysImageList.Obj := nil;
  SysImageList.Size := TSize.Create(0, 0);
end;

procedure FreeSysImageList(IconSizeType: TIconSizeType);
begin
  FreeSysImageList(gImageLists[IconSizeType]);
end;

procedure FreeSysImageListAll;
var
  SizeType: TIconSizeType;
begin
  for SizeType := Low(TIconSizeType) to High(TIconSizeType) do
    FreeSysImageList(gImageLists[SizeType]);
end;

function CatchImageList(SizeType: TIconSizeType; ForceLatest: Boolean): Boolean;
var
  pImageList: PSysImageList;
  Size: TSize;
  procedure RenewSize(pItem: PSysImageList); inline;
  var
    N: TSize;
  begin
    if pItem.Obj.GetIconSize(N.cx, N.cy) = S_OK then
      pItem.Size := N;
  end;
begin
  pImageList := @gImageLists[SizeType];
  if Assigned(pImageList.Obj) then
  begin
    if ForceLatest then
    begin
      FreeSysImageList(pImageList^);
      pImageList.Obj := GetImageListInterface(SizeType);
      if not Assigned(pImageList.Obj) then
        Exit(False);
    end
  end
  else
  begin
    pImageList.Obj := GetImageListInterface(SizeType);
    if not Assigned(pImageList.Obj) then
      Exit(False);
  end;

  if ForceLatest then
  begin
    RenewSize(pImageList);
    Exit(True);
  end;

  Size := pImageList.Size;
  if (Size.cx <= 0) or (Size.cy <= 0) then
    RenewSize(pImageList);
  Result := True;
end;

function CatchImageListAll(ForceLatest: Boolean): TIconSizeTypes;
var
  SizeType: TIconSizeType;
begin
  Result := [];
  for SizeType := _IconBest downto _IconBase do
    if CatchImageList(SizeType, ForceLatest) then
      Include(Result, SizeType);
end;

function CatchImageListAll(out ActualBest: TIconSizeType; ForceLatest: Boolean): TIconSizeTypes;
var
  SizeType: TIconSizeType;
  Size, LastSize: TSize;
begin
  Result := [];
  LastSize := TSize.Create(0, 0);
  for SizeType := _IconBest downto _IconBase do
  begin
    if not CatchImageList(SizeType, ForceLatest) then
      Continue;

    Size := gImageLists[SizeType].Size;
    if GetPixels(Size) > GetPixels(LastSize) then
      ActualBest := SizeType;
    LastSize := Size;

    Include(Result, SizeType);
  end;

  if (LastSize.cx > 0) and (LastSize.cy > 0) then
    Exit;

  for SizeType := _IconBest downto _IconBase do
    if SizeType in Result then
    begin
      ActualBest := SizeType;
      Break;
    end;
end;



function CatchImageListBest(out ASysImageList: TSysImageList; Flags: TImageListRetrieves): Boolean;
var
  SizeType: TIconSizeType;
begin
  gImageListMode := _IconBest;
  if ILR_ActualSize in Flags then
  begin
    if CatchImageListAll(gImageListMode, ILR_Renew in Flags) = [] then
      Exit(False);
    if ILR_NoKeep in Flags then
      for SizeType := Low(TIconSizeType) to High(TIconSizeType) do
        if SizeType <> gImageListMode then
          FreeSysImageList(SizeType);
    Result := True;
  end
  else
  begin
    Result := CatchImageList(_IconBest, ILR_Renew in Flags);
    if not Result then
      Exit(False);
  end;
  ASysImageList := gImageLists[gImageListMode];
end;

function SysImageList(SizeType: TIconSizeType; ExceptNil: Boolean): TSysImageList;
begin
  if (SizeType < Low(TIconSizeType)) or (SizeType > High(TIconSizeType)) then
    raise Exception.Create('Out of index.');

  Result := gImageLists[SizeType];

  if not Assigned(Result.Obj) and ExceptNil then
    raise Exception.Create('ImageList does not exist.');
end;

function SysImageList: TSysImageList;
begin
  Result := SysImageList(gImageListMode);
end;

function SysImageLists: TSysImageLists;
var
  SizeType: TIconSizeType;
begin
  for SizeType := Low(TIconSizeType) to High(TIconSizeType) do
    Result[SizeType] := SysImageList(SizeType);
end; 


function ConvertToIcon(hbmColor: HBITMAP; hbmMask: HBITMAP): TIcon;
var
  IconInfo: TIconInfo;
  Icon: HICON;
  DIB: TDIBSection;
  Mask: TBitmap;   
begin
// English:
// Regarding the DestroyIcon, CreateIconIndirect, and the TIcon.Handle property:
// Internal process of TIcon.Handle property setting:
// TIcon.Handle property >> TIcon.SetHandle function >> TIcon.NewImage function >>
// Replaces TIcon.FImage by creating a new TIconImage with the value specified by TIconImage.FHandle.
// When TIcon is destroyed, TIcon.FImage.Release will be executed, which will
// trigger DestroyIcon to release the value specified by TIcon.Handle.
// Therefore, the HICON set by TIcon.Handle will be released when the TIcon is released,
// without the need for an additional DestroyIcon.
// 中文：
// 關於 CreateIconIndirect 相應的 DestroyIcon，與 TIcon.Handle 屬性設定：
// TIcon.Handle 屬性設定的內部流程：
// TIcon.Handle >> TIcon.SetHandle >> TIcon.NewImage >> 
// 取代 TIcon.FImage 為建立新的 TIconImage 且該 TIconImage.FHandle 為 TIcon.Handle 指定的值。
// 當 TIcon 進行摧毀時會執行 TIcon.FImage.Release，進而引發 DestroyIcon 來釋放 TIcon.Handle 指定的值。   
// 因此，TIcon.Handle 設定的 HICON 會在 TIcon 釋放時一並釋放，不必額外 DestroyIcon。


// see ICONINFO structure
// https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-iconinfo
  IconInfo.fIcon := True; // True 表示輸出 Icon，False 表示輸出 Cursor
  IconInfo.xHotspot := 0; // Icon 則不需要，如 Cursor 則水平中心點
  IconInfo.yHotspot := 0; // Icon 則不需要，如 Cursor 則垂直中心點
  if hbmMask <> 0 then
  begin // Create an Icon with a specified mask.
    IconInfo.hbmMask  := hbmMask;
    IconInfo.hbmColor := hbmColor;
    Icon := CreateIconIndirect(IconInfo);
    if Icon = 0 then
      Exit(nil);
        
    Result := TIcon.Create;
    Result.Handle := Icon;
    Exit;
  end;

  FillChar(DIB, SizeOf(DIB), 0);
  // Get Bitmap image information.
  if GetObject(hbmColor, SizeOf(DIB), @DIB) = 0 then
    Exit(nil);

  // Create a new mask image.
  Mask := TBitmap.Create(DIB.dsBm.bmWidth, DIB.dsBm.bmHeight);
  try
    Mask.Monochrome := True;

    //
    // Draw a rectangular mask image with the bmColor image size.
    //
    Mask.Canvas.Lock;
    try
      Mask.Canvas.Brush.Color := clBlack;
      Mask.Canvas.FillRect(TRect.Create(0, 0, DIB.dsBm.bmWidth, DIB.dsBm.bmHeight));
    finally
      Mask.Canvas.Unlock;
    end;

    //
    // Create an Icon with a new mask.
    //
    IconInfo.hbmMask  := Mask.Handle;
    IconInfo.hbmColor := hbmColor;
    Icon := CreateIconIndirect(IconInfo);
    if Icon = 0 then
      Exit(nil);
        
    Result := TIcon.Create;
    Result.Handle := Icon;   
  finally
    Mask.Free;
  end;
end;

function ConvertToIcon(hbmColor: HBITMAP; hbmMask: HBITMAP; ASize: TSize): TIcon;
var                  
  bSize: Boolean;     
  Size: TSize;
  Bmp, BmpColor, BmpMask: TBitmap;
  Point: TPoint;
begin
  if ASize.IsZero then
    Exit(ConvertToIcon(hbmColor, hbmMask));

  Bmp := TBitmap.Create;
  try 
    Bmp.Handle := hbmColor;

    //
    // Checks whether the size is less than the specified size.(meaning, ASize is minimum size)
    //
    bSize := False;
    if ASize.Width < Bmp.Width then   
      Size.cx := Bmp.Width   
    else          
    begin
      bSize := True;
      Size.cx := ASize.Width
    end;
    if ASize.Height < Bmp.Height then  
      Size.cy := Bmp.Height   
    else    
    begin
      bSize := True;
      Size.cy := ASize.Height;
    end;

    if not bSize then
    begin // Create Icon at original size.
      Exit(ConvertToIcon(hbmColor, hbmMask));
    end;            
            
    Point := TPoint.Create((Size.cx - Bmp.Width) shr 1, (Size.cy - Bmp.Height) shr 1);

    BmpColor := TBitmap.Create;
    try
      BmpColor.Assign(Bmp);
      // Copy the image of TBitmap.Handle and disassociate it from the
      // system handle to make the hBitmap image content independent.
      BmpColor.HandleType := bmDDB;

      BmpColor.Palette := Bmp.Palette;  
      BmpColor.SetSize(Size.cx, Size.cy);
      BmpColor.Canvas.Lock;
      try            
        if BmpColor.Transparent and (BmpColor.TransparentMode = tmFixed) then
          BmpColor.Canvas.Brush.Color := Bmp.TransparentColor
        else
          BmpColor.Canvas.Brush.Color := clDefault;
        BmpColor.Canvas.FillRect(TRect.Create(0, 0, BmpColor.Width, BmpColor.Height));
        BmpColor.Canvas.Draw((Size.cx - Bmp.Width) shr 1, (Size.cy - Bmp.Height) shr 1, Bmp);
      finally
        BmpColor.Canvas.Unlock;
      end;

      if hbmMask = 0 then
      begin // Creates the Icon by a mask at the new size.
        BmpMask := TBitmap.Create(Size.cx, Size.cy);
        try
          // Set a mask image by the BmpColor image.
          BmpMask.Handle := BmpColor.MaskHandle;

          // Copy the image of TBitmap.Handle and disassociate it from the
          // system handle to make the hBitmap image content independent.
          BmpMask.HandleType := bmDDB;

          BmpMask.Monochrome := True; // Convert the image to single color.

          //
          // Fill the base
          //
          BmpMask.Canvas.Lock;
          try
            BmpMask.Canvas.Brush.Color := clBlack;
            BmpMask.Canvas.FillRect(TRect.Create(Point.X, Point.Y, Point.X + Bmp.Width, Point.Y + Bmp.Height));
          finally
            BmpMask.Canvas.Unlock;
          end;

          // Create Icon by new mask.
          Result := ConvertToIcon(BmpColor.Handle, BmpMask.Handle);
        finally
          BmpMask.Free;
        end;
      end
      else
      begin // Create the Icon by a mask by enlarging the base of the old mask with the new size..
        // 尚未測試
        Bmp.Handle := hbmMask;
        BmpMask := TBitmap.Create;
        try
          // Copy old mask image.
          BmpMask.Assign(Bmp);

          // Content copied via TBitmap.Assign still maintains the association (DIB),
          // so...
          // Copy the image of TBitmap.Handle and disassociate it from the
          // system handle to make the hBitmap image content independent.
          BmpMask.HandleType := bmDDB;

          BmpMask.Palette := Bmp.Palette;  
          BmpMask.SetSize(Size.cx, Size.cy);
          BmpMask.Monochrome := True;

          //
          // Fill the base then copy the old mask image to new mask image.
          //
          BmpMask.Canvas.Lock;
          try                                      
            BmpMask.Canvas.Brush.Color := clWhite;          
            BmpMask.Canvas.FillRect(TRect.Create(0, 0, BmpMask.Width, BmpMask.Height));
            BmpMask.Canvas.Draw((Size.cx - Bmp.Width) shr 1, (Size.cy - Bmp.Height) shr 1, Bmp);
          finally
            BmpMask.Canvas.Unlock;
          end;

          // Create Icon by new mask.
          Result := ConvertToIcon(BmpColor.Handle, BmpMask.Handle);    
        finally
          BmpMask.Free;
        end;
      end;
    finally       
      BmpColor.Free;
    end;
  finally
    Bmp.Free;
  end;
end;

function GetShellIcon(out Bitmap: HBITMAP; const AFilename: string; ASize: TSize; Flags: UINT): Boolean;
var
  ImageFactory: IShellItemImageFactory;
begin
  if SHCreateItemFromParsingName(PChar(AFilename), nil, IShellItemImageFactory, ImageFactory) = S_OK then
    Exit(ImageFactory.GetImage(ASize, Flags, Bitmap) = S_OK);
  Result := False;
end;

function GetShellIcon(out Bitmap: TBitmap; const AFilename: string; ASize: TSize; Flags: UINT): Boolean;
var
  hBmp: HBITMAP;
begin
  Result := GetShellIcon(hBmp, AFilename, ASize, Flags);
  if not Result then
    Exit;
  Bitmap := TBitmap.Create;
  Bitmap.TransparentMode := tmAuto;
  Bitmap.Handle := hBmp;
end;

function GetShellIcon(out Icon: TIcon; const AFilename: string; ASize: TSize; Flags: UINT): Boolean;
var
  Bitmap: HBITMAP;
begin
  Result := GetShellIcon(Bitmap, AFilename, ASize, Flags);
  if Result then
    Icon := ConvertToIcon(Bitmap, 0, ASize);
end;

function GetThumbnail(out Bitmap: HBITMAP; const AFilename: string; ASize: TSize; ColorDepth: DWORD): Boolean;
type
  TBufferISF = record      // For ISF parameter buffer.
    pItemIDL: PItemIDList;
    Attrib: DWORD;
    Eaten: DWORD;
  end;
  TBufferLocation = record // For IExtractImage::GetLocation parameter buffer.
    GLResult: HResult;
    Size: TSize;
    Flags: Cardinal;
    Priority: Cardinal;
  end;
  TBufferMultiple = record // Merge the parameter buffer spaces.
    case Integer of
      0: (ISF: TBufferISF);
      1: (IGL: TBufferLocation);
  end;
var
  Path, Name: String;
  DesktopISF, FolderISF: IShellFolder;
  IExtractImg: IExtractImage;
  Buff: TBufferMultiple;
  CharBuf: array[0..MAX_PATH] of WideChar; // Originally it was 0..2047.
//  MemAlloc: IMalloc;
begin
  Result := False;

  Path := ExtractFilePath(AFilename);
  Name := ExtractFileName(AFilename);
  if Name.IsEmpty then
    Exit;

//  if (SHGetMalloc(MemAlloc) <> NOERROR) or (MemAlloc = nil) or
//   (NOERROR <> SHGetDesktopFolder(DesktopISF)) then
//    Exit;

  if SHGetDesktopFolder(DesktopISF) <> S_OK then
    Exit;
  try
    if DesktopISF.ParseDisplayName(0, nil, PChar(Path), Buff.ISF.Eaten, Buff.ISF.pItemIDL, Buff.ISF.Attrib) <> S_OK then
      Exit;
    try
      if DesktopISF.BindToObject(Buff.ISF.pItemIDL, nil, IShellFolder, FolderISF) <> S_OK then
        Exit;
    finally
      CoTaskMemFree(Buff.ISF.pItemIDL);
    end;
    if FolderISF.ParseDisplayName(0, nil, PChar(Name), Buff.ISF.Eaten, Buff.ISF.pItemIDL, Buff.ISF.Attrib) <> S_OK then
      Exit;
    try
      if FolderISF.GetUIObjectOf(0, 1, Buff.ISF.pItemIDL, IExtractImage, nil, IExtractImg) <> S_OK then
        Exit;
    finally
      CoTaskMemFree(Buff.ISF.pItemIDL);
    end;
  finally
    FolderISF := nil;
    DesktopISF := nil;
  end;
  if not Assigned(IExtractImg) then
    Exit;
  try
//    FillChar(CharBuf, SizeOf(CharBuf), 0);
    Buff.IGL.Size := ASize;
    Buff.IGL.Priority := ITSAT_MIN_PRIORITY;
    Buff.IGL.Flags := IEIFLAG_SCREEN or IEIFLAG_OFFLINE or IEIFLAG_QUALITY;

    // IExtractImage::Extract must call IExtractImage::GetLocation prior to calling Extract.
    Buff.IGL.GLResult := IExtractImg.GetLocation(CharBuf, SizeOf(CharBuf), Buff.IGL.Priority, Buff.IGL.Size, ColorDepth, Buff.IGL.Flags);
    if Buff.IGL.GLResult <> NOERROR then
      Exit;

    if IExtractImg.Extract(Bitmap) = NOERROR then
      Result := Bitmap <> 0
    else
      Result := False;
  finally
    IExtractImg := nil;
  end;
end;

function GetThumbnail(out Bitmap: TBitmap; const AFilename: string; ASize: TSize; ColorDepth: DWORD): Boolean;
var
  hBmp: HBITMAP;
begin
  Result := GetThumbnail(hBmp, AFilename, ASize, ColorDepth);
  if not Result then
    Exit;
  Bitmap := TBitmap.Create;
  Bitmap.Handle := hBmp;
end;

function GetThumbnail(out Icon: TIcon; const AFilename: string; ASize: TSize; ColorDepth: DWORD): Boolean;
var
  hBmp: HBITMAP;
begin
  Result := GetThumbnail(hBmp, AFilename, ASize, ColorDepth);
  if Result then
    Icon := ConvertToIcon(hBmp, 0, ASize);
end;

function GetFileIcon(out Icon: HICON; const AFilename: string; Index: Integer; AMaxSize: Word; AMinSize: Word; ExtractFlags: TExtractIconFlags): TImageSize;
var
  Flags: UINT;
  hLargeIcon, hSmallIcon: HICON;
begin
  Result := ImageSize_Non;
  if EIF_SimulateDoc in ExtractFlags then
    Flags := GIL_SIMULATEDOC
  else
    Flags := 0;
  hLargeIcon := 0;
  hSmallIcon := 0;
  try
    if SHDefExtractIcon(PChar(AFilename), Index, Flags, hLargeIcon, hSmallIcon, AMinSize shl 16 or AMaxSize) <> S_OK then
      Exit;
    if not(EIF_IgnoreLarge in ExtractFlags) and (hLargeIcon <> 0) then
    begin
      Icon := hLargeIcon;
      hLargeIcon := 0;
      Exit(ImageSize_Large);
    end;
    if not(EIF_IgnoreSmall in ExtractFlags) and (hSmallIcon <> 0) then
    begin
      Icon := hSmallIcon;
      hSmallIcon := 0;
      Exit(ImageSize_Small);
    end;
  finally
    if hLargeIcon <> 0 then
      DestroyIcon(hLargeIcon);
    if hSmallIcon <> 0 then
      DestroyIcon(hSmallIcon);
  end;
end;

function GetFileIcon(out Icon: TIcon; const AFilename: string; Index: Integer; AMaxSize: Word; AMinSize: Word; ExtractFlags: TExtractIconFlags): TImageSize;
var
  IconHandle: HIcon;
begin
  Result := GetFileIcon(IconHandle, AFilename, Index, AMaxSize, AMinSize, ExtractFlags);
  if Result = ImageSize_Non then
    Exit;
  Icon := TIcon.Create;
  Icon.Handle := IconHandle;
end;

function GetFileIcon(out Bitmap: TBitmap; const AFilename: string; Index: Integer; AMaxSize: Word; AMinSize: Word; ExtractFlags: TExtractIconFlags): TImageSize;
var
  Icon: TIcon;
begin
  Result := GetFileIcon(Icon, AFilename, Index, AMaxSize, AMinSize, ExtractFlags);
  if Result = ImageSize_Non then
    Exit;
  try
    Bitmap := TBitmap.Create;
    Bitmap.Assign(Icon);
  finally
    Icon.Free;
  end;
end;

function GetSysImageListIcon(out Icon: HICON; const AFilename: string; SizeType: TIconSizeType): Boolean;
var
  ShInfo: TShFileInfo;
  GFIResult: IImageList;
begin
  Result := False;
  if not CatchImageList(SizeType, False) then
    Exit;

  GFIResult := IImageList(SHGetFileInfo(PChar(AFilename), FILE_ATTRIBUTE_NORMAL,
    ShInfo, SizeOf(ShInfo), SHGFI_SYSICONINDEX));
  if GFIResult = nil then
    Exit;

  ShInfo.hIcon := 0;
  if SysImageList(SizeType).Obj.GetIcon(ShInfo.iIcon, ILD_PRESERVEALPHA, ShInfo.hIcon) <> S_OK then
    Exit;
  if ShInfo.hIcon = 0 then
    Exit;

  Icon := ShInfo.hIcon;
  Result := True;
end;

function GetSysImageListIcon(out Icon: TIcon; const AFilename: string; SizeType: TIconSizeType): Boolean;
var
  h: HICON;
begin
  Result := GetSysImageListIcon(h, AFilename, SizeType);
  if not Result then
    Exit;
  Icon := TIcon.Create;
  Icon.Handle := h;
end;

function GetSysImageListIcon(out Bitmap: TBitmap; const AFilename: string; SizeType: TIconSizeType): Boolean;
var
  Icon: TIcon;
begin
  Result := GetSysImageListIcon(Icon, AFilename, SizeType);
  if not Result then
    Exit;
  try
    Bitmap := TBitmap.Create;
    Bitmap.Assign(Icon);
  finally
    Icon.Free;
  end;
end;

{ TIconSizeModeHelper }

function TIconSizeTypeHelper.GetFlag: Integer;
begin
  Result := cIconTypically[Self].iImageList;
end;

function TIconSizeTypeHelper.GetId: string;
begin
  Result := cIconTypically[Self].Id;
end;

function TIconSizeTypeHelper.GetString: string;
begin
  Result := cIconTypically[Self].Name;
end;

function TIconSizeTypeHelper.GetTypicallySize: Word;
begin
  Result := cIconTypically[Self].Size;
end;

function TIconSizeTypeHelper.GetTypicallyPixels: TSize;
var
  N: Word;
begin
  N := GetTypicallySize;
  Result := TSize.Create(N, N);
end;

function TIconSizeTypeHelper.GetSysImageList: TSysImageList;
begin
  CatchImageList(Self, False);
  Result := gImageLists[Self];
end;

function TIconSizeTypeHelper.GetImageList: IImageList;
begin
  CatchImageList(Self, False);
  Result := gImageLists[Self].Obj;
end;

function TIconSizeTypeHelper.GetActualPixels: TSize;
begin
  CatchImageList(Self, False);
  Result := gImageLists[Self].Size;
end;

initialization
  FillChar(gImageLists, SizeOf(gImageLists), 0);

finalization
  FreeSysImageListAll;

end.
